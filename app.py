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
from master_analyzer import analyze_master
from prompt_injector import inject_master_into_prompt

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

# ─── Cache master PPTX trong memory ─────────────────────────────────────────
# TODO: Upgrade lên Flask session + temp file nếu cần multi-user production
_master_cache: dict = {}   # { "bytes": bytes, "info": dict, "filename": str, "schema": dict, "dynamic_prompt": str }


# ─── Hàm tiện ích ───────────────────────────────────────────────────────────
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

    # ── Bước 3: Tạo file PPTX ─────────────────────────────────────────────
    try:
        pptx_buffer = build_pptx_from_slides(slides)
    except Exception as exc:
        app.logger.error("Lỗi khi tạo PPTX: %s", exc, exc_info=True)
        return jsonify({
            "success": False,
            "error": f"Lỗi khi tạo file PPTX: {str(exc)}"
        }), 500

    # ── Bước 4: Trả về file để tải xuống ──────────────────────────────────
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


# ─── Route: Upload Master Slide ─────────────────────────────────────────────
@app.route("/api/upload-master", methods=["POST"])
def upload_master():
    """
    Nhận file PPTX upload, parse master info + phân tích sâu schema,
    inject schema vào prompt master, lưu vào cache tạm.
    Form field: master_file (file .pptx)
    Trả về: { success, filename, layouts, theme_colors, default_font, dynamic_prompt }
    """
    if "master_file" not in request.files:
        return jsonify({"success": False, "error": "Không tìm thấy file"}), 400

    file = request.files["master_file"]
    if not file.filename or not file.filename.lower().endswith(".pptx"):
        return jsonify({"success": False, "error": "Chỉ chấp nhận file .pptx"}), 400

    pptx_bytes = file.read()
    if len(pptx_bytes) > 100 * 1024 * 1024:  # Giới hạn 100MB
        return jsonify({"success": False, "error": "File quá lớn (tối đa 100MB)"}), 400

    # ── Bước 1: Parse master info (layouts cơ bản) ────────────────────────
    try:
        master_info = parse_master_info(pptx_bytes)
    except Exception as exc:
        app.logger.error("Lỗi parse master: %s", exc, exc_info=True)
        return jsonify({"success": False, "error": f"Không thể đọc file PPTX: {str(exc)}"}), 400

    # ── Bước 2: Phân tích sâu master schema (non-critical) ────────────────
    schema = {}
    try:
        schema = analyze_master(pptx_bytes)
    except Exception as exc:
        app.logger.warning("Phân tích master schema thất bại (non-critical): %s", exc)

    # ── Bước 3: Inject schema vào prompt master ───────────────────────────
    dynamic_prompt = ""
    try:
        prompts = load_prompts()
        prompt_template = prompts.get("prompt_master", "")
        if prompt_template and schema:
            dynamic_prompt = inject_master_into_prompt(prompt_template, schema)
    except Exception as exc:
        app.logger.warning("Inject prompt thất bại (non-critical): %s", exc)

    # ── Bước 4: Lưu vào cache ─────────────────────────────────────────────
    _master_cache["bytes"]          = pptx_bytes
    _master_cache["info"]           = master_info
    _master_cache["filename"]       = file.filename
    _master_cache["schema"]         = schema
    _master_cache["dynamic_prompt"] = dynamic_prompt

    return jsonify({
        "success":        True,
        "filename":       file.filename,
        "layouts":        [{"index": l["index"], "name": l["name"]} for l in master_info["layouts"]],
        "theme_colors":   master_info["theme_colors"],
        "default_font":   master_info["default_font"],
        "dynamic_prompt": dynamic_prompt,
    })


# ─── Route: Lấy dynamic prompt theo master đã upload ────────────────────────
@app.route("/api/master-prompt", methods=["GET"])
def get_master_prompt():
    """
    Trả về dynamic prompt đã được inject schema của master slide đang được cache.
    Nếu chưa upload master → 404.
    """
    if not _master_cache.get("bytes"):
        return jsonify({"success": False, "error": "Chưa upload master slide."}), 404

    dynamic_prompt = _master_cache.get("dynamic_prompt", "")

    # Nếu cha có dynamic_prompt (schema build lỗi) → thử regenerate
    if not dynamic_prompt:
        try:
            prompts = load_prompts()
            prompt_template = prompts.get("prompt_master", "")
            schema = _master_cache.get("schema", {})
            if prompt_template:
                dynamic_prompt = inject_master_into_prompt(prompt_template, schema)
                _master_cache["dynamic_prompt"] = dynamic_prompt
        except Exception as exc:
            app.logger.warning("Regenerate dynamic prompt thất bại: %s", exc)

    return jsonify({
        "success":        True,
        "filename":       _master_cache.get("filename", ""),
        "dynamic_prompt": dynamic_prompt,
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
            "error": "Chưa upload master slide. Vui lòng upload file .pptx trước."
        }), 400

    svg_code = request.form.get("svg_code", "").strip()
    topic    = request.form.get("topic", "presentation").strip()

    is_valid, error_msg = validate_svg_input(svg_code)
    if not is_valid:
        return jsonify({"success": False, "error": error_msg}), 400

    slides = extract_slides_from_svg(svg_code)
    if not slides:
        return jsonify({"success": False, "error": "Không tìm thấy slide nào."}), 400

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
    app.run(debug=True, host="0.0.0.0", port=5000)
