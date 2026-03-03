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
from pptx_builder import build_pptx_from_slides
from svg_fixer import fix_svg

# ─── Khởi tạo ứng dụng Flask ────────────────────────────────────────────────
app = Flask(__name__)
app.config["MAX_CONTENT_LENGTH"] = 16 * 1024 * 1024  # Giới hạn 16 MB mỗi request

# Đường dẫn file prompt
PROMPTS_FILE = Path(__file__).parent / "prompts.json"


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
    """Trả về prompt template hiện tại từ file prompts.json."""
    prompts = load_prompts()
    return jsonify(prompts)


# ─── Route: Lưu prompt ──────────────────────────────────────────────────────
@app.route("/api/prompt", methods=["POST"])
def save_prompt():
    """Nhận prompt mới từ client và lưu vào file prompts.json."""
    try:
        data = request.get_json() or {}
        
        # Kiểm tra dữ liệu
        if "ai_prompt_template" not in data:
            return jsonify({"success": False, "error": "Thiếu trường 'ai_prompt_template'"}), 400
        
        # Lưu vào file
        if save_prompts(data):
            return jsonify({"success": True, "message": "Đã lưu prompt thành công"})
        else:
            return jsonify({"success": False, "error": "Lỗi lưu file"}), 500
            
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


# ─── Chạy ứng dụng ───────────────────────────────────────────────────────────
if __name__ == "__main__":
    app.run(debug=True, host="0.0.0.0", port=5000)
