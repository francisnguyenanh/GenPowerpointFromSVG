"""
prompt_injector.py
------------------
Nhận master schema từ master_analyzer.analyze_master() và inject thông tin
đó vào trong prompt template, tạo ra dynamic prompt có đầy đủ:
  - Bảng màu theme
  - Font heading/body
  - Danh sách layouts với data-layout values
  - Tọa độ + kích thước các placeholder zones (px, 1280×720)
  - Cấp độ bullet/font-size cho body placeholder

Chiến lược inject:
  1. Tìm placeholder {{MASTER_CONTEXT}} trong template → thay thế trực tiếp
  2. Nếu không có → tìm trigger phrase "hãy tạo mã svg"/"create svg" → prepend
  3. Nếu không có trigger → prepend vào đầu prompt
"""

import re


# ─── Helpers để tạo từng phần của context block ─────────────────────────────

def _build_colors_block(schema: dict) -> str:
    """Tạo bảng màu theme."""
    colors: dict = schema.get("theme", {}).get("colors", {})
    if not colors:
        return ""
    lines = ["### 🎨 Theme Colors:"]
    for name, hex_val in colors.items():
        lines.append(f"  - {name}: {hex_val}")
    return "\n".join(lines)


def _build_font_block(schema: dict) -> str:
    """Tạo thông tin font heading/body."""
    fonts: dict = schema.get("theme", {}).get("fonts", {})
    if not fonts:
        return ""
    heading = fonts.get("heading", "Calibri Light")
    body    = fonts.get("body",    "Calibri")
    return (
        "### 🔤 Fonts:\n"
        f"  - Heading/Title: \"{heading}\"\n"
        f"  - Body/Content:  \"{body}\""
    )


def _build_layouts_block(schema: dict) -> str:
    """Tạo danh sách layouts với data-layout values."""
    layouts: list = schema.get("layouts", [])
    if not layouts:
        return ""
    lines = ["### 📐 Available Layouts (dùng trong thuộc tính data-layout của <g>):"]
    seen = set()
    for layout in layouts:
        val  = layout.get("data_layout_value", "content")
        name = layout.get("name", f"Layout {layout.get('index',0)}")
        if val not in seen:
            lines.append(f'  - data-layout="{val}"  →  "{name}"')
            seen.add(val)
    return "\n".join(lines)


def _build_placeholder_zone_rules(schema: dict) -> str:
    """
    Tạo quy tắc placeholder zones cho từng layout.
    Canvas chuẩn: 1280×720 px.
    """
    layouts: list = schema.get("layouts", [])
    if not layouts:
        return ""

    blocks = ["### 📦 Placeholder Zones (px, canvas 1280×720):"]

    # Nhóm layouts theo data_layout_value (tránh lặp)
    seen = set()
    for layout in layouts:
        val  = layout.get("data_layout_value", "content")
        name = layout.get("name", "")
        phs  = layout.get("placeholders", [])
        if not phs:
            continue

        if val in seen:
            continue
        seen.add(val)

        blocks.append(f'\n  [data-layout="{val}"]  ← "{name}"')

        for ph in phs:
            ph_type = ph.get("type", "body")
            l, t    = ph.get("left_px", 0), ph.get("top_px", 0)
            w, h    = ph.get("width_px", 0), ph.get("height_px", 0)
            fs      = ph.get("font_size_pt", 18)
            bold    = " bold" if ph.get("font_bold") else ""
            color   = f", color: {ph['font_color']}" if ph.get("font_color") else ""

            blocks.append(
                f"    • {ph_type:20s}: "
                f"x={l:4d} y={t:4d} w={w:4d} h={h:4d}  "
                f"[font: {fs}pt{bold}{color}]"
            )

            # Bullet levels cho body
            bullet_levels = ph.get("bullet_levels", [])
            for lvl in bullet_levels:
                lvl_num    = lvl.get("level", 1)
                lvl_sz     = lvl.get("font_size_pt", 18)
                lvl_indent = lvl.get("indent_px", 0)
                blocks.append(
                    f"        ↳ level {lvl_num}: {lvl_sz}pt, "
                    f"indent {lvl_indent}px → dùng data-level=\"{lvl_num}\""
                )

    return "\n".join(blocks)


def build_master_context_section(schema: dict) -> str:
    """
    Gộp tất cả các block thành một master context section hoàn chỉnh
    để inject vào prompt.
    """
    meta       = schema.get("meta", {})
    w_px       = meta.get("slide_width_px",  1280)
    h_px       = meta.get("slide_height_px", 720)

    parts = [
        "=" * 60,
        "🎯 MASTER SLIDE CONTEXT — ĐÃ ĐƯỢC TỰ ĐỘNG PHÂN TÍCH",
        f"Canvas chuẩn: {w_px}×{h_px} px (viewBox=\"0 0 {w_px} {h_px}\")",
        "=" * 60,
    ]

    for builder in [_build_colors_block, _build_font_block,
                    _build_layouts_block, _build_placeholder_zone_rules]:
        try:
            block = builder(schema)
            if block:
                parts.append(block)
        except Exception:
            continue

    parts.append("=" * 60)
    return "\n\n".join(parts)


# ─── Hàm inject chính ───────────────────────────────────────────────────────

def inject_master_into_prompt(prompt_template: str, schema: dict) -> str:
    """
    Inject master context vào prompt template.

    Chiến lược (ưu tiên theo thứ tự):
    1. Thay thế placeholder {{MASTER_CONTEXT}} nếu có.
    2. Prepend trước trigger phrase tiếng Việt/Anh nếu tìm thấy.
    3. Prepend vào đầu prompt.

    Parameters
    ----------
    prompt_template : str  — Prompt gốc (thường lấy từ prompts.json)
    schema          : dict — Output của analyze_master()

    Returns
    -------
    str — Prompt đã được inject master context
    """
    if not prompt_template:
        return ""
    if not schema:
        return prompt_template

    try:
        context_section = build_master_context_section(schema)
    except Exception:
        return prompt_template  # nếu build lỗi, trả về prompt gốc

    # ── Chiến lược 1: {{MASTER_CONTEXT}} placeholder ─────────────────────
    if "{{MASTER_CONTEXT}}" in prompt_template:
        return prompt_template.replace("{{MASTER_CONTEXT}}", context_section)

    # ── Chiến lược 2: trigger phrase ──────────────────────────────────────
    triggers = [
        r"hãy\s+tạo\s+mã\s+svg",
        r"tạo\s+mã\s+svg",
        r"create\s+svg",
        r"generate\s+svg",
        r"sinh\s+mã\s+svg",
    ]
    for pattern in triggers:
        match = re.search(pattern, prompt_template, flags=re.IGNORECASE)
        if match:
            insert_pos = match.start()
            return (
                prompt_template[:insert_pos]
                + context_section + "\n\n"
                + prompt_template[insert_pos:]
            )

    # ── Chiến lược 3: prepend vào đầu ─────────────────────────────────────
    return context_section + "\n\n" + prompt_template
