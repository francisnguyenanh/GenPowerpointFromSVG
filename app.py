"""
app.py
------
Ứng dụng Flask: Tạo file PPTX từ mã SVG.

Routes:
  GET  /                  — Trang chính (giao diện nhập liệu)
  GET  /api/prompt        — Lấy prompt template hiện tại từ file prompts.json
  POST /api/prompt        — Lưu prompt template mới
  POST /generate          — Nhận SVG, xử lý, và trả về file PPTX để tải xuống
"""

import json
import os
from pathlib import Path
from flask import Flask, render_template, request, send_file, jsonify

from svg_processor import extract_slides_from_svg, validate_svg_input
from pptx_builder import build_pptx_from_slides, build_pptx_from_slides_with_master
from svg_fixer import fix_svg
from master_handler import parse_master_info

# ─── Khởi tạo ứng dụng Flask ────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 100 * 1024 * 1024  # Giới hạn 16 MB mỗi request

# ─── Error handlers trả JSON thay vì HTML ───────────────────────────────────
@app.errorhandler(413)
def request_too_large(e):
    return jsonify({"success": False, "error": "File quá lớn (tối đa 16MB)."}), 413

@app.errorhandler(500)
def internal_error(e):
    return jsonify({"success": False, "error": f"Lỗi server nội bộ: {str(e)}"}), 500

# Đường dẫn file prompt
PROMPTS_FILE = Path(__file__).parent / "prompts.json"

# Thư mục lưu SVG từng slide
OUTPUT_SVG_DIR = Path(__file__).parent / "output_svg"
OUTPUT_SVG_DIR.mkdir(exist_ok=True)

# Thư mục chứa file master PPTX cố định
INPUT_DIR = Path(__file__).parent / "input"
INPUT_DIR.mkdir(exist_ok=True)

# ─── Cache master PPTX (tự động load từ input/) ──────────────────────────────
_master_cache: dict = {}   # { "bytes": bytes, "info": dict, "filename": str }


def _auto_load_master() -> bool:
    """
    Tự động tìm và load file .pptx đầu tiên trong thư mục input/.
    Trả về True nếu load thành công, False nếu không tìm thấy.
    """
    pptx_files = sorted(INPUT_DIR.glob("*.pptx"))
    if not pptx_files:
        app.logger.warning("Không tìm thấy file .pptx nào trong input/")
        return False
    pptx_path = pptx_files[0]
    try:
        pptx_bytes = pptx_path.read_bytes()
        master_info = parse_master_info(pptx_bytes)
        _master_cache["bytes"]    = pptx_bytes
        _master_cache["info"]     = master_info
        _master_cache["filename"] = pptx_path.name
        app.logger.info("Đã load master slide: %s", pptx_path.name)
        return True
    except Exception as exc:
        app.logger.error("Lỗi load master từ input/: %s", exc)
        return False


# ─── Hàm tiện ích ───────────────────────────────────────────────────────────

def save_slides_to_output_svg(slides: list, topic: str = "slide") -> list:
    """
    Lưu từng slide SVG thành file riêng vào thư mục output_svg/.
    Tên file: output_svg/{safe_topic}_slide_{index:02d}.svg

    Trả về danh sách đường dẫn đã lưu.
    """
    safe_topic = "".join(
        c if c.isalnum() or c in ("-", "_") else "_"
        for c in (topic or "slide").strip()
    ).strip("_")[:40] or "slide"

    saved = []
    for slide in slides:
        idx = slide.get("index", slide.get("id", 1))
        svg_content = slide.get("svg", "")
        if not svg_content:
            continue
        filepath = OUTPUT_SVG_DIR / f"{safe_topic}_slide_{int(idx):02d}.svg"
        try:
            filepath.write_text(svg_content, encoding="utf-8")
            saved.append(filepath)
        except Exception as exc:
            app.logger.warning("Không lưu được SVG slide %s: %s", idx, exc)
    return saved


def load_prompts() -> dict:
    """Tải dữ liệu từ file prompts.json."""
    if not PROMPTS_FILE.exists():
        return {"ai_prompt_template": ""}
    try:
        with open(PROMPTS_FILE, "r", encoding="utf-8") as f:
            return json.load(f)
    except json.JSONDecodeError:
        return {"ai_prompt_template": ""}


def save_prompts(data: dict) -> bool:
    """Lưu dữ liệu vào file prompts.json."""
    try:
        with open(PROMPTS_FILE, "w", encoding="utf-8") as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as exc:
        app.logger.error("Lỗi lưu prompts.json: %s", exc)
        return False


# ─── Route: Trang chính ──────────────────────────────────────────────────────
@app.route("/", methods=["GET"])
def index():
    """Hiển thị trang giao diện chính."""
    return render_template("index.html")


# ─── Route: Lấy prompt ──────────────────────────────────────────────────────
@app.route("/api/prompt", methods=["GET"])
def get_prompt():
    """
    Trả về cả 2 prompt template:
      - prompt_new    : prompt tạo PPTX mới (từ key ai_prompt_template)
      - prompt_master : prompt map vào master slide
    """
    prompts = load_prompts()
    return jsonify({
        "prompt_new":    prompts.get("ai_prompt_template", ""),
        "prompt_master": prompts.get("prompt_master", ""),
    })


# ─── Route: Lưu prompt ──────────────────────────────────────────────────────
@app.route("/api/prompt", methods=["POST"])
def save_prompt():
    """
    Lưu prompt template theo mode.
    Body JSON: { "mode": "new"|"master", "template": "..." }
    """
    try:
        data = request.get_json() or {}
        mode     = data.get("mode", "new")          # "new" hoặc "master"
        template = data.get("template", "").strip()

        if not template:
            return jsonify({"success": False, "error": "Template không được để trống."}), 400

        prompts = load_prompts()
        if mode == "master":
            prompts["prompt_master"] = template
        else:  # "new" hoặc default
            prompts["ai_prompt_template"] = template

        if save_prompts(prompts):
            return jsonify({"success": True, "message": f"Đã lưu prompt [{mode}] thành công."})
        else:
            return jsonify({"success": False, "error": "Lỗi lưu file."}), 500

    except Exception as exc:
        app.logger.error("Lỗi lưu prompt: %s", exc)
        return jsonify({"success": False, "error": str(exc)}), 500

# ─── Route: Kiểm tra & Sửa lỗi SVG ───────────────────────────────────────────
@app.route("/api/fix-svg", methods=["POST"])
def fix_svg_route():
    """
    Nhận mã SVG thô, chạy pipeline kiểm tra và tự động sửa lỗi.
    Trả về JSON gồm:
      - success     : bool
      - fixed_svg   : cỗuỗi SVG đã sửa
      - fixes       : danh sách những cải thiện đã làm
      - warnings    : cảnh báo
      - errors      : lỗi không sửa được
      - slide_count : số slide tìm thấy
    """
    data = request.get_json() or {}
    raw_svg = data.get("svg_code", "").strip()

    if not raw_svg:
        return jsonify({"success": False, "errors": ["Mã SVG không được để trống."]}), 400

    try:
        result = fix_svg(raw_svg)
        return jsonify({
            "success":     result.success,
            "fixed_svg":   result.fixed_svg,
            "fixes":       result.fixes,
            "warnings":    result.warnings,
            "errors":      result.errors,
            "slide_count": result.slide_count,
        })
    except Exception as exc:
        app.logger.error("Lỗi fix-svg: %s", exc, exc_info=True)
        return jsonify({"success": False, "errors": [str(exc)]}), 500

# ─── Route: Tạo file PPTX ────────────────────────────────────────────────────
@app.route("/generate", methods=["POST"])
def generate():
    """
    Nhận mã SVG từ form, tách thành các slide, và trả về file PPTX.

    Form fields:
      - svg_code  : Chuỗi mã SVG đầy đủ (bắt buộc)
      - topic     : Chủ đề (dùng để đặt tên file tải xuống)
    """
    # Lấy dữ liệu từ form
    svg_code = request.form.get("svg_code", "").strip()
    topic    = request.form.get("topic", "presentation").strip()

    # ── Bước 1: Kiểm tra đầu vào ──────────────────────────────────────────
    is_valid, error_msg = validate_svg_input(svg_code)
    if not is_valid:
        # Trả về lỗi dạng JSON để JS xử lý
        return jsonify({"success": False, "error": error_msg}), 400

    # ── Bước 2: Tách SVG thành danh sách các slide ────────────────────────
    slides = extract_slides_from_svg(svg_code)

    if not slides:
        return jsonify({
            "success": False,
            "error": (
                "Không tìm thấy slide nào. "
                "Hãy kiểm tra lại mã SVG và đảm bảo có các thẻ "
                "&lt;g id='slide_1'&gt;, &lt;g id='slide_2'&gt;..."
            )
        }), 400

    # ── Bước 3: Lưu từng slide SVG ra output_svg/ ────────────────────────
    try:
        saved = save_slides_to_output_svg(slides, topic)
        app.logger.info("Đã lưu %d slide SVG vào output_svg/", len(saved))
    except Exception as exc:
        app.logger.warning("save_slides_to_output_svg thất bại (non-critical): %s", exc)

    # ── Bước 4: Tạo file PPTX ─────────────────────────────────────────────
    try:
        pptx_buffer = build_pptx_from_slides(slides)
    except Exception as exc:
        app.logger.error("Lỗi khi tạo PPTX: %s", exc, exc_info=True)
        return jsonify({
            "success": False,
            "error": f"Lỗi khi tạo file PPTX: {str(exc)}"
        }), 500

    # ── Bước 5: Trả về file để tải xuống ──────────────────────────────────
    # Tạo tên file an toàn từ chủ đề (loại ký tự đặc biệt)
    safe_topic = "".join(
        c if c.isalnum() or c in (" ", "-", "_") else "_"
        for c in topic
    ).strip().replace(" ", "_")[:50]

    filename = f"{safe_topic or 'presentation'}.pptx"

    return send_file(
        pptx_buffer,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=filename,
    )


# ─── Route: Trạng thái master slide ─────────────────────────────────────────
@app.route("/api/master-status", methods=["GET"])
def get_master_status():
    """
    Kiểm tra xem master slide đã được load từ input/ hay chưa.
    Trả về: { loaded: bool, filename: str }
    """
    loaded   = bool(_master_cache.get("bytes"))
    filename = _master_cache.get("filename", "")
    return jsonify({"loaded": loaded, "filename": filename})


# ─── Route: Lấy prompt master ────────────────────────────────────────────────
@app.route("/api/master-prompt", methods=["GET"])
def get_master_prompt():
    """
    Trả về prompt_master từ prompts.json (không inject schema).
    """
    prompts = load_prompts()
    prompt  = prompts.get("prompt_master", "")
    return jsonify({
        "success":  True,
        "prompt":   prompt,
        "filename": _master_cache.get("filename", ""),
        "loaded":   bool(_master_cache.get("bytes")),
    })


# ─── Route: Tạo PPTX với Master Slide ───────────────────────────────────────
@app.route("/generate-with-master", methods=["POST"])
def generate_with_master():
    """
    Nhận SVG + dùng master PPTX đã upload, trả về file PPTX map vào master.
    Form fields: svg_code, topic
    Yêu cầu: phải upload master trước qua /api/upload-master
    """
    if not _master_cache.get("bytes"):
        return jsonify({
            "success": False,
            "error": "Không tìm thấy master slide. Vui lòng đặt file .pptx vào thư mục input/."
        }), 400

    svg_code = request.form.get("svg_code", "").strip()
    topic    = request.form.get("topic", "presentation").strip()

    is_valid, error_msg = validate_svg_input(svg_code)
    if not is_valid:
        return jsonify({"success": False, "error": error_msg}), 400

    slides = extract_slides_from_svg(svg_code)
    if not slides:
        return jsonify({"success": False, "error": "Không tìm thấy slide nào."}), 400

    # Lưu từng slide SVG ra output_svg/
    try:
        saved = save_slides_to_output_svg(slides, topic)
        app.logger.info("Đã lưu %d slide SVG vào output_svg/", len(saved))
    except Exception as exc:
        app.logger.warning("save_slides_to_output_svg thất bại (non-critical): %s", exc)

    try:
        pptx_buffer = build_pptx_from_slides_with_master(
            slides,
            _master_cache["bytes"],
            _master_cache["info"],
        )
    except Exception as exc:
        app.logger.error("Lỗi generate-with-master: %s", exc, exc_info=True)
        return jsonify({"success": False, "error": f"Lỗi tạo PPTX: {str(exc)}"}), 500

    safe_topic = "".join(
        c if c.isalnum() or c in (" ", "-", "_") else "_" for c in topic
    ).strip().replace(" ", "_")[:50]
    filename = f"{safe_topic or 'presentation'}_master.pptx"

    return send_file(
        pptx_buffer,
        mimetype="application/vnd.openxmlformats-officedocument.presentationml.presentation",
        as_attachment=True,
        download_name=filename,
    )


# ─── Chạy ứng dụng ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    with app.app_context():
        _auto_load_master()
    app.run(debug=True, host="0.0.0.0", port=5000)
