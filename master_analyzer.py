"""
master_analyzer.py
------------------
Phân tích sâu file PPTX master slide, trích xuất toàn bộ thông tin cần thiết
để inject vào prompt cho AI gen SVG phù hợp với cấu trúc master.

Thông tin trích xuất:
  - Tất cả layouts: tên, data-layout value tương ứng, danh sách placeholders
  - Mỗi placeholder: idx, type, label, tọa độ px (1280×720), font size, màu
  - Theme colors: dk1, lt1, dk2, lt2, accent1–6
  - Fonts: heading font, body font
  - Background: solid/gradient/inherit cho mỗi layout
"""

import io
from lxml import etree
from pptx import Presentation
from pptx.util import Pt

# ─── Tỉ lệ chuyển đổi EMU → px (chuẩn 1280×720) ────────────────────────────
_EMU_TO_PX_X = 1280 / 9144000
_EMU_TO_PX_Y = 720  / 5143500

# ─── Map tên layout → data-layout value dùng trong SVG ──────────────────────
_LAYOUT_NAME_MAP = {
    "title slide":          "title-slide",
    "title and content":    "content",
    "two content":          "two-column",
    "comparison":           "two-column",
    "section header":       "section-header",
    "blank":                "blank",
    "content with caption": "content",
    "picture with caption": "content",
    "title only":           "section-header",
}

# ─── Map placeholder idx → type string ──────────────────────────────────────
_PH_IDX_TYPE_MAP = {
    0:  "center_title",
    1:  "body",
    2:  "subtitle",
    3:  "title",
    10: "date",
    11: "footer",
    12: "slide_number",
    13: "object",
    14: "picture",
}


def _emu_to_px(emu_val, axis: str = "x") -> int:
    """Chuyển đổi EMU sang pixel (làm tròn)."""
    ratio = _EMU_TO_PX_X if axis == "x" else _EMU_TO_PX_Y
    return round(int(emu_val or 0) * ratio)


def _extract_theme_data(prs: Presentation) -> dict:
    """
    Đọc theme colors và font names từ XML nội bộ của slide master.
    Trả về dict với keys: fonts {heading, body}, colors {dk1, lt1, ...accent1-6}.
    Tất cả exception được catch silently → fallback về default values.
    """
    fonts  = {"heading": "Calibri Light", "body": "Calibri"}
    colors = {}

    try:
        master_part = prs.slide_master.part
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"

        for rel in master_part.rels.values():
            if "theme" not in rel.reltype:
                continue
            try:
                theme_xml = etree.fromstring(rel.target_part.blob)
            except Exception:
                continue

            ns = {"a": ns_a}

            # ── Fonts ──────────────────────────────────────────────────────
            try:
                major = theme_xml.find(".//a:fontScheme/a:majorFont/a:latin", ns)
                minor = theme_xml.find(".//a:fontScheme/a:minorFont/a:latin", ns)
                if major is not None and major.get("typeface"):
                    fonts["heading"] = major.get("typeface")
                if minor is not None and minor.get("typeface"):
                    fonts["body"] = minor.get("typeface")
            except Exception:
                pass

            # ── Colors ─────────────────────────────────────────────────────
            try:
                wanted = {"dk1", "lt1", "dk2", "lt2",
                          "accent1", "accent2", "accent3",
                          "accent4", "accent5", "accent6"}
                for node in theme_xml.findall(".//a:clrScheme/*", ns):
                    try:
                        tag = etree.QName(node.tag).localname
                        if tag not in wanted:
                            continue
                        child = list(node)
                        if not child:
                            continue
                        hex_val = (child[0].get("val") or child[0].get("lastClr", "")).upper()
                        if hex_val:
                            colors[tag] = f"#{hex_val.lstrip('#')}"
                    except Exception:
                        continue
            except Exception:
                pass

            break  # chỉ cần theme đầu tiên

    except Exception:
        pass  # fallback về giá trị mặc định

    return {"fonts": fonts, "colors": colors}


def _extract_background(layout) -> dict:
    """
    Trích xuất thông tin background của một layout.
    Trả về {"type": "solid"|"gradient"|"inherit", "colors": [hex...]}.
    Tất cả exception được catch silently.
    """
    try:
        fill = layout.background.fill
        if fill.type is not None:
            fill_type_str = str(fill.type)
            if "SOLID" in fill_type_str or fill_type_str == "1":
                try:
                    hex_color = f"#{fill.fore_color.rgb}"
                    return {"type": "solid", "colors": [hex_color]}
                except Exception:
                    pass
    except Exception:
        pass

    try:
        ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
        grad = layout.element.find(f".//{{{ns_a}}}gradFill")
        if grad is not None:
            stops = grad.findall(f".//{{{ns_a}}}gs")
            hex_colors = []
            for stop in stops:
                try:
                    srgb = stop.find(f"{{{ns_a}}}srgbClr")
                    if srgb is not None:
                        val = srgb.get("val", "")
                        if val:
                            hex_colors.append(f"#{val.upper()}")
                except Exception:
                    continue
            if hex_colors:
                return {"type": "gradient", "colors": hex_colors}
    except Exception:
        pass

    return {"type": "inherit", "colors": []}


def _extract_bullet_levels(ph) -> list:
    """
    Đọc font size theo từng cấp indent trong body placeholder.
    Trả về list[{"level": int, "font_size_pt": int, "indent_px": int}].
    Tất cả exception được catch silently → fallback về defaults.
    """
    ns_a = "http://schemas.openxmlformats.org/drawingml/2006/main"
    levels = []

    try:
        tx_body = ph._element.find(f"{{{ns_a}}}txBody")
        if tx_body is None:
            raise ValueError("no txBody")

        for i in range(1, 4):
            tag = f"lvl{i}pPr"
            sz_pt = 18  # default
            try:
                node = tx_body.find(f".//{{{ns_a}}}{tag}")
                if node is not None:
                    def_rpr = node.find(f"{{{ns_a}}}defRPr")
                    if def_rpr is not None:
                        sz_raw = def_rpr.get("sz")
                        if sz_raw:
                            sz_pt = int(sz_raw) // 100
            except Exception:
                pass
            levels.append({
                "level":        i,
                "font_size_pt": sz_pt,
                "indent_px":    (i - 1) * 32,
            })
    except Exception:
        levels = [
            {"level": 1, "font_size_pt": 18, "indent_px": 0},
            {"level": 2, "font_size_pt": 16, "indent_px": 32},
            {"level": 3, "font_size_pt": 14, "indent_px": 64},
        ]

    return levels


def analyze_master(pptx_bytes: bytes) -> dict:
    """
    Hàm chính: phân tích PPTX và trả về master schema đầy đủ.
    Tất cả exception nội bộ được catch silently → luôn trả về dict hợp lệ.

    Returns dict:
    {
      "meta": {
        "slide_width_emu": int, "slide_height_emu": int,
        "slide_width_px": 1280, "slide_height_px": 720
      },
      "theme": {
        "fonts": {"heading": str, "body": str},
        "colors": {"dk1": "#hex", "lt1": "#hex", "accent1": "#hex", ...}
      },
      "layouts": [
        {
          "index": int,
          "name": str,
          "data_layout_value": str,
          "background": {"type": str, "colors": [str]},
          "placeholders": [
            {
              "idx": int,
              "type": str,
              "label": str,
              "left_px": int, "top_px": int,
              "width_px": int, "height_px": int,
              "left_emu": int, "top_emu": int,
              "width_emu": int, "height_emu": int,
              "font_size_pt": int,
              "font_bold": bool,
              "font_color": str,
              "bullet_levels": [...]   # chỉ có khi type == "body"
            }
          ]
        }
      ]
    }
    """
    try:
        prs = Presentation(io.BytesIO(pptx_bytes))
    except Exception as exc:
        raise ValueError(f"Không thể đọc file PPTX: {exc}") from exc

    schema = {
        "meta": {
            "slide_width_emu":  int(prs.slide_width),
            "slide_height_emu": int(prs.slide_height),
            "slide_width_px":   1280,
            "slide_height_px":  720,
        },
        "theme":   _extract_theme_data(prs),
        "layouts": [],
    }

    for idx, layout in enumerate(prs.slide_layouts):
        try:
            name = (layout.name or f"Layout {idx}").strip()
            data_layout = _LAYOUT_NAME_MAP.get(name.lower(), "content")

            layout_info = {
                "index":             idx,
                "name":              name,
                "data_layout_value": data_layout,
                "background":        _extract_background(layout),
                "placeholders":      [],
            }

            for ph in layout.placeholders:
                try:
                    ph_idx  = ph.placeholder_format.idx
                    ph_type = _PH_IDX_TYPE_MAP.get(ph_idx, "body")

                    # ── Tọa độ ───────────────────────────────────────────
                    try:
                        left   = int(ph.left   or 0)
                        top    = int(ph.top    or 0)
                        width  = int(ph.width  or 0)
                        height = int(ph.height or 0)
                    except Exception:
                        left = top = width = height = 0

                    # ── Font từ text frame ────────────────────────────────
                    font_size_pt = 18
                    font_bold    = False
                    font_color   = ""
                    try:
                        tf = ph.text_frame
                        if tf.paragraphs and tf.paragraphs[0].runs:
                            run = tf.paragraphs[0].runs[0]
                            if run.font.size:
                                font_size_pt = int(run.font.size.pt)
                            font_bold = bool(run.font.bold)
                            try:
                                font_color = f"#{run.font.color.rgb}"
                            except Exception:
                                pass
                    except Exception:
                        pass

                    ph_data = {
                        "idx":          ph_idx,
                        "type":         ph_type,
                        "label":        ph.name or f"ph_{ph_idx}",
                        "left_emu":     left,  "top_emu":    top,
                        "width_emu":    width, "height_emu": height,
                        "left_px":      _emu_to_px(left,   "x"),
                        "top_px":       _emu_to_px(top,    "y"),
                        "width_px":     _emu_to_px(width,  "x"),
                        "height_px":    _emu_to_px(height, "y"),
                        "font_size_pt": font_size_pt,
                        "font_bold":    font_bold,
                        "font_color":   font_color,
                    }

                    if ph_type == "body":
                        ph_data["bullet_levels"] = _extract_bullet_levels(ph)

                    layout_info["placeholders"].append(ph_data)

                except Exception:
                    continue  # bỏ qua placeholder lỗi, tiếp tục

            schema["layouts"].append(layout_info)

        except Exception:
            continue  # bỏ qua layout lỗi, tiếp tục

    return schema
