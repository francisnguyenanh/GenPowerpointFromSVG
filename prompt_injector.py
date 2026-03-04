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
    """Tạo thông tin font heading/body (bao gồm tên gốc nếu là alias)."""
    fonts: dict = schema.get("theme", {}).get("fonts", {})
    if not fonts:
        return ""
    heading     = fonts.get("heading",     "Calibri Light")
    body        = fonts.get("body",        "Calibri")
    heading_raw = fonts.get("heading_raw", "")
    body_raw    = fonts.get("body_raw",    "")

    lines = [
        "### 🔤 Fonts:",
        f'  - Heading/Title: "{heading}"'
        + (f'  (raw: "{heading_raw}")' if heading_raw and heading_raw != heading else ""),
        f'  - Body/Content:  "{body}"'
        + (f'  (raw: "{body_raw}")' if body_raw and body_raw != body else ""),
    ]

    # Cảnh báo nếu là Japanese/alias font
    alias_flags = {"+mj-ea", "+mn-ea", "+mj-lt", "+mn-lt"}
    if heading_raw.lower() in alias_flags or body_raw.lower() in alias_flags:
        lines.append(
            "  ⚠️  Font theme là alias (+mj-ea / +mn-ea). "
            "Dùng fallback: Yu Gothic Light / Meiryo trong SVG."
        )

    return "\n".join(lines)


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

            font_name = ph.get("font_name", "")
            align     = ph.get("align", "")
            font_str  = f"font: {fs}pt{bold}"
            if font_name:
                font_str += f", {font_name}"
            if align:
                font_str += f", align={align}"
            font_str += color

            blocks.append(
                f"    • {ph_type:20s}: "
                f"x={l:4d} y={t:4d} w={w:4d} h={h:4d}  "
                f"[{font_str}]"
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


def _build_content_zone_block(schema: dict) -> str:
    """
    Tạo bảng content_zone (vùng AI tự do thiết kế) cho mỗi layout.
    Nhóm theo data_layout_value để tránh lặp.
    """
    layouts: list = schema.get("layouts", [])
    if not layouts:
        return ""

    lines = ["### 🖼️ Content Zones — Vùng tự do thiết kế (px):",
             "  (Không bị che bởi title/header/footer)",
             "  Tất cả hình ảnh, icon, ilustration, diagram nên nằm trong vùng này."]
    seen = set()
    for layout in layouts:
        val  = layout.get("data_layout_value", "content")
        name = layout.get("name", "")
        zone = layout.get("content_zone")
        if zone is None or val in seen:
            continue
        seen.add(val)
        x, y = zone.get("x", 0), zone.get("y", 0)
        w, h = zone.get("w", 0), zone.get("h", 0)
        lines.append(
            f'  [{val:20s}] "{name}": '
            f"x={x:4d} y={y:4d} w={w:4d} h={h:4d}"
        )
    return "\n".join(lines)


def build_master_context_section(schema: dict) -> str:
    """
    Gộp tất cả các block thành một master context section hoàn chỉnh
    để inject vào prompt.
    """
    meta  = schema.get("meta", {})
    w_px  = meta.get("slide_width_px",  1280)
    h_px  = meta.get("slide_height_px", 720)
    w_emu = meta.get("slide_width_emu",  0)
    h_emu = meta.get("slide_height_emu", 0)

    size_note = ""
    if w_emu and h_emu:
        size_note = f" (EMU: {w_emu}×{h_emu})"

    parts = [
        "=" * 60,
        "🎯 MASTER SLIDE CONTEXT — ĐÃ ĐƯỢC TỰ ĐỘNG PHÂN TÍCH",
        f"Canvas chuẩn: {w_px}×{h_px} px (viewBox=\"0 0 {w_px} {h_px}\"){size_note}",
        "=" * 60,
    ]

    for builder in [_build_colors_block, _build_font_block,
                    _build_layouts_block, _build_placeholder_zone_rules,
                    _build_content_zone_block]:
        try:
            block = builder(schema)
            if block:
                parts.append(block)
        except Exception:
            continue

    parts.append("=" * 60)
    return "\n\n".join(parts)


# ─── Hàm inject chính ───────────────────────────────────────────────────────

def inject_master_into_prompt(
    prompt_template: str,
    schema: dict,
    topic: str = "",
    num_slides: int = 8,
    language: str = "Tiếng Việt",
    style: str = "Hiện đại, chuyên nghiệp"
) -> str:
    """
    Inject master context vào prompt template, đồng thời thay thế các biến {topic}, {num_slides}, {language}, {style}.

    Parameters
    ----------
    prompt_template : str  — Prompt gốc (thường lấy từ prompts.json)
    schema          : dict — Output của analyze_master()
    topic           : str  — Chủ đề
    num_slides      : int  — Số lượng slide
    language        : str  — Ngôn ngữ
    style           : str  — Phong cách trình bày

    Returns
    -------
    str — Prompt đã được inject master context và thay thế biến
    """
    if not prompt_template:
        return ""
    if not schema:
        return prompt_template

    # Thay thế các biến trong template
    prompt_filled = prompt_template
    prompt_filled = prompt_filled.replace("{topic}", topic or "[Chủ đề]")
    prompt_filled = prompt_filled.replace("{num_slides}", str(num_slides or 8))
    prompt_filled = prompt_filled.replace("{language}", language or "Tiếng Việt")
    prompt_filled = prompt_filled.replace("{style}", style or "Hiện đại, chuyên nghiệp")

    try:
        context_section = build_master_context_section(schema)
    except Exception:
        return prompt_filled  # nếu build lỗi, trả về prompt đã thay biến

    # ── Chiến lược 1: {{MASTER_CONTEXT}} placeholder ─────────────────────
    if "{{MASTER_CONTEXT}}" in prompt_filled:
        return prompt_filled.replace("{{MASTER_CONTEXT}}", context_section)

    # ── Chiến lược 2: trigger phrase ──────────────────────────────────────
    triggers = [
        r"hãy\s+tạo\s+mã\s+svg",
        r"tạo\s+mã\s+svg",
        r"create\s+svg",
        r"generate\s+svg",
        r"sinh\s+mã\s+svg",
    ]
    for pattern in triggers:
        match = re.search(pattern, prompt_filled, flags=re.IGNORECASE)
        if match:
            insert_pos = match.start()
            return (
                prompt_filled[:insert_pos]
                + context_section + "\n\n"
                + prompt_filled[insert_pos:]
            )

    # ── Chiến lược 3: prepend vào đầu ─────────────────────────────────────
    return context_section + "\n\n" + prompt_filled
