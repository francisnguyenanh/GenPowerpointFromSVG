"""
master_handler.py
-----------------
Xử lý master slide từ file PPTX upload.
Đọc thông tin layout, theme, placeholder để map SVG semantic content vào đúng vị trí.

Các hàm chính:
  - parse_master_info(pptx_bytes)         → dict chứa layout, theme, font
  - find_best_layout(master_info, layout) → index layout phù hợp nhất
  - extract_svg_semantic_content(svg)     → dict nội dung từ data-role attributes
  - build_pptx_with_master(slides, ...)   → io.BytesIO file PPTX đã map vào master
"""

import io
import re
from lxml import etree
from pptx import Presentation
from pptx.util import Pt
from pptx.enum.text import PP_ALIGN


# ─── Namespace cho DrawingML theme XML ───────────────────────────────────────
_DML_NS  = {"a": "http://schemas.openxmlformats.org/drawingml/2006/main"}
_SVG_NS  = "http://www.w3.org/2000/svg"

# Mapping placeholder type integer → tên đọc được
_PH_TYPE_MAP = {
    1:  "title",
    2:  "body",
    3:  "center_title",
    4:  "subtitle",
    5:  "body",        # date
    6:  "footer",
    7:  "slide_number",
    10: "object",
    15: "title",
    None: "body",
}


# ═══════════════════════════════════════════════════════════════════════════
# 1. parse_master_info
# ═══════════════════════════════════════════════════════════════════════════

def parse_master_info(pptx_bytes: bytes) -> dict:
    """
    Đọc file PPTX và trả về thông tin layout, theme, placeholder.

    Args:
        pptx_bytes: Nội dung file .pptx dạng bytes.

    Returns:
        dict với keys: slide_width_emu, slide_height_emu, layouts, theme_colors, default_font
    """
    prs = Presentation(io.BytesIO(pptx_bytes))

    # ── Kích thước slide ────────────────────────────────────────────────────
    slide_width_emu  = int(prs.slide_width)
    slide_height_emu = int(prs.slide_height)

    # ── Theme colors ────────────────────────────────────────────────────────
    theme_colors = _extract_theme_colors(prs)

    # ── Default font ────────────────────────────────────────────────────────
    default_font = _extract_default_font(prs)

    # ── Layouts ─────────────────────────────────────────────────────────────
    layouts = []
    for idx, layout in enumerate(prs.slide_layouts):
        placeholders = []
        for ph in layout.placeholders:
            ph_type_int = ph.placeholder_format.type if ph.placeholder_format else None
            ph_type_str = _PH_TYPE_MAP.get(
                int(ph_type_int) if ph_type_int is not None else None, "body"
            )
            try:
                left   = int(ph.left)   if ph.left   is not None else 0
                top    = int(ph.top)    if ph.top    is not None else 0
                width  = int(ph.width)  if ph.width  is not None else 0
                height = int(ph.height) if ph.height is not None else 0
            except Exception:
                left = top = width = height = 0

            placeholders.append({
                "idx":        ph.placeholder_format.idx if ph.placeholder_format else 0,
                "type":       ph_type_str,
                "left_emu":   left,
                "top_emu":    top,
                "width_emu":  width,
                "height_emu": height,
            })

        layouts.append({
            "index":        idx,
            "name":         layout.name or f"Layout {idx}",
            "placeholders": placeholders,
        })

    return {
        "slide_width_emu":  slide_width_emu,
        "slide_height_emu": slide_height_emu,
        "layouts":          layouts,
        "theme_colors":     theme_colors,
        "default_font":     default_font,
    }


def _extract_theme_colors(prs: Presentation) -> dict:
    """Đọc màu theme từ slide master XML bằng XPath đúng namespace."""
    color_keys = ["dk1", "lt1", "dk2", "lt2",
                  "accent1", "accent2", "accent3", "accent4", "accent5", "accent6"]
    result = {k: "" for k in color_keys}

    try:
        # Lấy theme XML từ slide master đầu tiên
        master = prs.slide_masters[0]
        theme_element = master.theme_color_map  # nếu có
    except Exception:
        return result

    try:
        # Truy cập raw XML của theme qua part
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        tree = etree.fromstring(theme_part.blob)

        for key in color_keys:
            # Tìm node <a:dk1>, <a:lt1>... trong fmtScheme/fontScheme và clrScheme
            nodes = tree.xpath(
                f".//a:clrScheme/a:{key}/*",
                namespaces=_DML_NS
            )
            if nodes:
                node = nodes[0]
                # sysClr có lastClr, srgbClr có val
                color_val = node.get("lastClr") or node.get("val") or ""
                if color_val:
                    result[key] = f"#{color_val.upper()}"
    except Exception:
        pass  # Nếu không đọc được theme → trả về dict rỗng

    return result


def _extract_default_font(prs: Presentation) -> str:
    """Đọc font chữ mặc định từ theme hoặc slide master."""
    try:
        master = prs.slide_masters[0]
        theme_part = master.part.part_related_by(
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
        )
        tree = etree.fromstring(theme_part.blob)
        # Tìm latin font trong majorFont
        nodes = tree.xpath(
            ".//a:fontScheme/a:majorFont/a:latin",
            namespaces=_DML_NS
        )
        if nodes:
            typeface = nodes[0].get("typeface", "")
            if typeface:
                return typeface
        # Fallback: minorFont
        nodes = tree.xpath(
            ".//a:fontScheme/a:minorFont/a:latin",
            namespaces=_DML_NS
        )
        if nodes:
            typeface = nodes[0].get("typeface", "")
            if typeface:
                return typeface
    except Exception:
        pass
    return "Calibri"  # fallback mặc định


# ═══════════════════════════════════════════════════════════════════════════
# 2. find_best_layout
# ═══════════════════════════════════════════════════════════════════════════

# Map data-layout → danh sách keyword tìm trong tên layout (ưu tiên từ trái)
_LAYOUT_KEYWORDS = {
    "title-slide":    ["title slide", "title, slide", "title only"],
    "section-header": ["section header", "section", "divider"],
    "content":        ["title and content", "title, content", "content"],
    "two-column":     ["two content", "comparison", "two column", "2 content"],
    "big-stat":       ["title and content", "content"],
    "blank":          ["blank"],
}
_LAYOUT_FALLBACK_INDEX = {
    "title-slide":    0,
    "section-header": 2,
    "content":        1,
    "two-column":     3,
    "big-stat":       1,
    "blank":          6,
}


def find_best_layout(master_info: dict, svg_data_layout: str) -> int:
    """
    Map data-layout từ SVG sang layout index trong master PPTX.

    Args:
        master_info:     Kết quả từ parse_master_info().
        svg_data_layout: Giá trị thuộc tính data-layout trong SVG.

    Returns:
        Index của layout phù hợp nhất (int).
    """
    layout_key  = (svg_data_layout or "content").lower().strip()
    keywords    = _LAYOUT_KEYWORDS.get(layout_key, _LAYOUT_KEYWORDS["content"])
    fallback_idx = _LAYOUT_FALLBACK_INDEX.get(layout_key, 6)

    layouts = master_info.get("layouts", [])
    max_idx  = len(layouts) - 1

    for keyword in keywords:
        for layout in layouts:
            if keyword in layout["name"].lower():
                return layout["index"]

    # Fallback theo index cố định (clamp để tránh vượt giới hạn)
    return min(fallback_idx, max_idx) if max_idx >= 0 else 0


# ═══════════════════════════════════════════════════════════════════════════
# 3. extract_svg_semantic_content
# ═══════════════════════════════════════════════════════════════════════════

def extract_svg_semantic_content(slide_svg: str) -> dict:
    """
    Parse SVG XML để lấy nội dung từ các data-role attribute.
    Bỏ qua hoàn toàn <g data-role="decorative">.

    Args:
        slide_svg: Chuỗi SVG đầy đủ của một slide đơn.

    Returns:
        dict với layout, title, subtitle, content, content_left, content_right, footer.
    """
    result = {
        "layout":        "content",
        "title":         "",
        "subtitle":      "",
        "content":       [],
        "content_left":  [],
        "content_right": [],
        "footer":        "",
    }

    try:
        # Parse SVG — lax để xử lý SVG không hoàn hảo
        parser = etree.XMLParser(recover=True, encoding="utf-8")
        try:
            root = etree.fromstring(slide_svg.encode("utf-8"), parser)
        except Exception:
            return result

        # Hàm lấy text thuần từ một element (bao gồm text trong sub-elements)
        def _collect_text(el) -> str:
            parts = []
            for t in el.itertext():
                s = t.strip()
                if s:
                    parts.append(s)
            return " ".join(parts)

        # Hàm parse các content item trong một group
        def _parse_content_group(group_el) -> list:
            items = []
            # Tìm tất cả element con có data-type hoặc text
            for child in group_el.iter():
                # Chỉ lấy text/tspan/g/rect với text thực
                role = child.get("data-role", "")
                if role == "decorative":
                    continue
                dtype = child.get("data-type", "paragraph")
                level = int(child.get("data-level", "1") or "1")
                source = child.get("data-source", "")
                author = child.get("data-author", "")

                # Lấy text trực tiếp (không đệ quy để tránh trùng)
                text_parts = []
                if child.text and child.text.strip():
                    text_parts.append(child.text.strip())
                for sub in child:
                    if sub.get("data-role", "") == "decorative":
                        continue
                    if sub.text and sub.text.strip():
                        text_parts.append(sub.text.strip())
                    if sub.tail and sub.tail.strip():
                        text_parts.append(sub.tail.strip())

                text = " ".join(text_parts).strip()
                if text and child.tag != group_el.tag:
                    items.append({
                        "text":   text,
                        "type":   dtype,
                        "level":  level,
                        "source": source,
                        "author": author,
                    })

            # Nếu không có sub-items → lấy text của chính group
            if not items:
                text = _collect_text(group_el)
                if text:
                    items.append({
                        "text":   text,
                        "type":   "paragraph",
                        "level":  1,
                        "source": "",
                        "author": "",
                    })
            return items

        # Tìm <g id="slide_N"> — root của nội dung slide
        slide_g = None
        for el in root.iter():
            el_id = el.get("id", "")
            if re.match(r"^slide_\d+$", el_id):
                slide_g = el
                break

        if slide_g is None:
            slide_g = root  # Fallback: dùng root SVG

        # Lấy data-layout từ slide group
        result["layout"] = slide_g.get("data-layout", "content")

        # Duyệt tất cả <g data-role="..."> trong slide
        for group in slide_g.iter():
            tag_local = etree.QName(group.tag).localname if "{" in group.tag else group.tag
            if tag_local != "g":
                continue

            role = group.get("data-role", "").lower().strip()

            if role == "decorative":
                continue
            elif role == "title":
                result["title"] = _collect_text(group)
            elif role == "subtitle":
                result["subtitle"] = _collect_text(group)
            elif role == "footer":
                result["footer"] = _collect_text(group)
            elif role == "content":
                result["content"] = _parse_content_group(group)
            elif role == "content-left":
                result["content_left"] = _parse_content_group(group)
            elif role == "content-right":
                result["content_right"] = _parse_content_group(group)

    except Exception:
        pass  # Trả về dict mặc định nếu parse lỗi

    return result


# ═══════════════════════════════════════════════════════════════════════════
# 4. build_pptx_with_master
# ═══════════════════════════════════════════════════════════════════════════

def _set_placeholder_text(ph, content_items: list) -> None:
    """
    Điền nội dung từ list content_items vào một placeholder text frame.
    Không set font.name và màu — kế thừa từ master.
    """
    try:
        tf = ph.text_frame
        tf.clear()
        # Xóa hết paragraph cũ
        for para in tf.paragraphs[1:]:
            p_el = para._p
            p_el.getparent().remove(p_el)

        first = True
        for item in content_items:
            if first:
                para = tf.paragraphs[0]
                first = False
            else:
                para = tf.add_paragraph()

            para.level = max(0, (item.get("level", 1) or 1) - 1)  # python-pptx 0-based

            text = item.get("text", "")
            dtype = item.get("type", "paragraph") or "paragraph"

            # Biến đổi text theo type
            if dtype == "quote":
                text = f"\u275d {text} \u275e"
            # (stat-label, caption, bullet, paragraph giữ nguyên text)

            run = para.add_run()
            run.text = text

            # Font size theo type — KHÔNG set font.name
            if dtype == "stat-number":
                run.font.size = Pt(48)
                run.font.bold = True
            elif dtype == "caption":
                run.font.size = Pt(10)
            elif dtype == "quote":
                run.font.italic = True

    except Exception:
        pass  # Bỏ qua silently nếu placeholder không hỗ trợ


def _find_placeholder(slide, idx: int):
    """Tìm placeholder theo idx. Trả về None nếu không tồn tại."""
    for ph in slide.placeholders:
        try:
            if ph.placeholder_format.idx == idx:
                return ph
        except Exception:
            continue
    return None


def _find_placeholder_by_type(slide, type_str: str):
    """Tìm placeholder theo type string. Trả về None nếu không có."""
    type_map = {
        "title":    [0, 13],
        "body":     [1, 2],
        "footer":   [10, 11, 12],
        "subtitle": [1],
    }
    indices = type_map.get(type_str, [])
    for idx in indices:
        ph = _find_placeholder(slide, idx)
        if ph is not None:
            return ph
    return None


def build_pptx_with_master(
    slides: list[dict],
    pptx_bytes: bytes,
    master_info: dict,
) -> io.BytesIO:
    """
    Tạo PPTX mới dựa trên master từ pptx_bytes bằng cách map nội dung SVG
    vào đúng placeholder của từng layout.

    Args:
        slides:      Danh sách slide dict từ extract_slides_from_svg().
                     Mỗi item gồm: 'id', 'index', 'svg'.
        pptx_bytes:  Nội dung file master .pptx dạng bytes.
        master_info: Kết quả từ parse_master_info().

    Returns:
        io.BytesIO chứa file PPTX đã map xong.
    """
    if not slides:
        raise ValueError("Danh sách slides không được rỗng.")

    # ── Mở PPTX master ──────────────────────────────────────────────────────
    prs = Presentation(io.BytesIO(pptx_bytes))
    # KHÔNG thay đổi slide_width/slide_height — kế thừa từ master

    # ── Xóa toàn bộ slide hiện có ───────────────────────────────────────────
    try:
        xml_slides = prs.slides._sldIdLst
        for slide_id in list(xml_slides):
            xml_slides.remove(slide_id)
    except Exception:
        pass

    max_layout_idx = len(prs.slide_layouts) - 1

    # ── Thêm slide mới theo nội dung SVG ────────────────────────────────────
    for slide_data in slides:
        svg_text = slide_data.get("svg", "")

        # 1. Lấy nội dung semantic từ SVG
        semantic = extract_svg_semantic_content(svg_text)

        # 2. Tìm layout phù hợp
        layout_idx  = find_best_layout(master_info, semantic.get("layout", "content"))
        layout_idx  = min(layout_idx, max_layout_idx)
        layout      = prs.slide_layouts[layout_idx]

        # 3. Thêm slide mới
        slide = prs.slides.add_slide(layout)

        # 4. Map title (idx=0)
        title_ph = _find_placeholder(slide, 0)
        if title_ph is not None:
            try:
                title_ph.text = semantic.get("title", "")
            except Exception:
                pass

        # 5. Map subtitle (idx=1 trên title-slide layout)
        subtitle_text = semantic.get("subtitle", "")
        if subtitle_text:
            sub_ph = _find_placeholder(slide, 1)
            if sub_ph is not None:
                try:
                    sub_ph.text = subtitle_text
                except Exception:
                    pass

        # 6. Map content
        layout_name_lower = layout.name.lower()
        is_two_column = any(k in layout_name_lower for k in ["two content", "comparison", "two column", "2 content"])

        if is_two_column:
            # Map content_left → idx=1, content_right → idx=2
            left_items  = semantic.get("content_left", []) or semantic.get("content", [])
            right_items = semantic.get("content_right", [])
            ph_left  = _find_placeholder(slide, 1)
            ph_right = _find_placeholder(slide, 2)
            if ph_left is not None and left_items:
                _set_placeholder_text(ph_left, left_items)
            if ph_right is not None and right_items:
                _set_placeholder_text(ph_right, right_items)
        else:
            content_items = semantic.get("content", [])
            if content_items and not subtitle_text:
                # Tìm body placeholder (idx=1)
                body_ph = _find_placeholder(slide, 1)
                if body_ph is not None:
                    _set_placeholder_text(body_ph, content_items)

        # 7. Map footer
        footer_text = semantic.get("footer", "")
        if footer_text:
            # Thử idx=10, 11, 12 (footer placeholders)
            for footer_idx in [10, 11, 12]:
                footer_ph = _find_placeholder(slide, footer_idx)
                if footer_ph is not None:
                    try:
                        footer_ph.text = footer_text
                        break
                    except Exception:
                        continue

    # ── Lưu ra BytesIO ──────────────────────────────────────────────────────
    output = io.BytesIO()
    prs.save(output)
    output.seek(0)
    return output
