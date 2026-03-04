"""
svg_processor.py
----------------
Module xử lý và tách mã SVG thành các slide riêng lẻ.
Mỗi slide được nhận dạng bằng thẻ <g id="slide_n">.
"""

import re
from bs4 import BeautifulSoup


# Kích thước chuẩn cho mỗi slide (tỉ lệ 16:9)
SLIDE_WIDTH  = 1280
SLIDE_HEIGHT = 720

# Danh sách font dự phòng an toàn (phổ biến trên mọi hệ điều hành)
SAFE_FONT_FALLBACK = "Arial, Helvetica, sans-serif"

# Bản đồ thay thế font để tránh lỗi hiển thị trên server không có font
FONT_REPLACEMENTS = {
    "Georgia":          "Times New Roman, serif",
    "Verdana":          SAFE_FONT_FALLBACK,
    "Trebuchet MS":     SAFE_FONT_FALLBACK,
    "Tahoma":           SAFE_FONT_FALLBACK,
    "Comic Sans MS":    SAFE_FONT_FALLBACK,
    "Impact":           SAFE_FONT_FALLBACK,
    "Courier New":      "Courier New, Courier, monospace",
    "Courier":          "Courier New, Courier, monospace",
    "Times":            "Times New Roman, serif",
}


def normalize_fonts(svg_string: str) -> str:
    """
    Chuẩn hóa thuộc tính font-family trong SVG.
    Đảm bảo các font không có sẵn trên server được thay thế bằng font an toàn.
    """
    def replace_font(match):
        font_value = match.group(1)
        for original, replacement in FONT_REPLACEMENTS.items():
            # So sánh không phân biệt hoa thường
            if original.lower() in font_value.lower():
                return f'font-family="{replacement}"'
        # Nếu không tìm thấy trong bản đồ, thêm fallback sans-serif
        return f'font-family="{font_value}, sans-serif"'

    # Thay thế thuộc tính font-family dạng attribute
    svg_string = re.sub(
        r'font-family=["\']([^"\']+)["\']',
        replace_font,
        svg_string,
        flags=re.IGNORECASE
    )

    # Thay thế font-family trong inline style="..."
    def replace_style_font(match):
        style_content = match.group(1)
        def inner_replace(m):
            font_value = m.group(1)
            for original, replacement in FONT_REPLACEMENTS.items():
                if original.lower() in font_value.lower():
                    return f'font-family:{replacement}'
            return f'font-family:{font_value}, sans-serif'

        style_content = re.sub(
            r'font-family:\s*([^;"\'}]+)',
            inner_replace,
            style_content,
            flags=re.IGNORECASE
        )
        return f'style="{style_content}"'

    svg_string = re.sub(
        r'style="([^"]*font-family[^"]*)"',
        replace_style_font,
        svg_string,
        flags=re.IGNORECASE
    )

    return svg_string


def wrap_group_in_svg(group_tag, definitions: str = "") -> str:
    """
    Bọc nội dung của một thẻ <g> vào trong thẻ <svg> hoàn chỉnh.
    Giữ lại id, data-layout và các data-* attribute để downstream parser đọc được.

    Args:
        group_tag: Đối tượng BeautifulSoup đại diện cho thẻ <g>.
        definitions: Chuỗi nội dung <defs> từ SVG gốc (gradient, filter...).

    Returns:
        Chuỗi SVG hoàn chỉnh cho một slide.
    """
    inner_content = group_tag.decode_contents()

    # Giữ lại tất cả attribute quan trọng (id, data-*, transform, class)
    preserved_attrs = ""
    for attr, val in group_tag.attrs.items():
        if isinstance(val, list):   # BS4 trả list cho class
            val = " ".join(val)
        preserved_attrs += f' {attr}="{val}"'

    svg_string = (
        f'<svg xmlns="http://www.w3.org/2000/svg" '
        f'xmlns:xlink="http://www.w3.org/1999/xlink" '
        f'viewBox="0 0 {SLIDE_WIDTH} {SLIDE_HEIGHT}" '
        f'width="{SLIDE_WIDTH}" height="{SLIDE_HEIGHT}">\n'
        f'{definitions}\n'
        f'<g{preserved_attrs}>\n'
        f'{inner_content}\n'
        f'</g>\n'
        f'</svg>'
    )

    return svg_string


def extract_slides_from_svg(svg_source: str) -> list[dict]:
    """
    Phân tích mã SVG và trích xuất từng slide.

    Quy tắc nhận dạng slide:
      - Là thẻ <g> có thuộc tính id bắt đầu bằng 'slide_'
      - Ví dụ: <g id="slide_1">, <g id="slide_2">

    Args:
        svg_source: Chuỗi mã SVG đầy đủ được nhập từ người dùng.

    Returns:
        Danh sách dict, mỗi dict gồm:
          - 'id'   : id của thẻ <g> (vd: "slide_1")
          - 'index': số thứ tự slide (1-based)
          - 'svg'  : chuỗi SVG hoàn chỉnh cho slide đó
    """
    if not svg_source or not svg_source.strip():
        return []

    # Dùng lxml parser để xử lý SVG phức tạp
    soup = BeautifulSoup(svg_source, "lxml")

    # Trích xuất phần <defs> từ SVG gốc (gradient, clipPath, filter...)
    # để các slide con vẫn có thể tham chiếu đến
    defs_tag = soup.find("defs")
    definitions = str(defs_tag) if defs_tag else ""

    # Tìm tất cả thẻ <g> có id bắt đầu bằng 'slide_'
    slide_groups = soup.find_all(
        "g",
        id=lambda tag_id: tag_id and tag_id.startswith("slide_")
    )

    if not slide_groups:
        return []

    # Sắp xếp theo số thứ tự trong id (vd: slide_1 < slide_2 < slide_10)
    def sort_key(tag):
        match = re.search(r"slide_(\d+)", tag.get("id", ""))
        return int(match.group(1)) if match else 0

    slide_groups.sort(key=sort_key)

    slides = []
    for idx, group in enumerate(slide_groups, start=1):
        slide_id = group.get("id", f"slide_{idx}")

        # Tạo SVG hoàn chỉnh cho slide này
        svg_content = wrap_group_in_svg(group, definitions)

        # Chuẩn hóa font (xử lý vấn đề font chữ)
        svg_content = normalize_fonts(svg_content)

        slides.append({
            "id":          slide_id,
            "index":       idx,
            "svg":         svg_content,
            "data_layout": group.get("data-layout", ""),  # cache trước khi BS4 wrap mất
        })

    return slides


def validate_svg_input(svg_source: str) -> tuple[bool, str]:
    """
    Kiểm tra tính hợp lệ của mã SVG đầu vào.

    Returns:
        (True, "") nếu hợp lệ.
        (False, "thông báo lỗi") nếu không hợp lệ.
    """
    if not svg_source or not svg_source.strip():
        return False, "Mã SVG không được để trống."

    stripped = svg_source.strip().lower()

    # Kiểm tra có chứa thẻ <svg> không
    if "<svg" not in stripped:
        return False, "Mã SVG không hợp lệ: Không tìm thấy thẻ &lt;svg&gt;."

    # Kiểm tra có ít nhất một slide group không
    if 'id="slide_' not in svg_source and "id='slide_" not in svg_source:
        return (
            False,
            "Không tìm thấy thẻ &lt;g id='slide_n'&gt; nào. "
            "Hãy đảm bảo mã SVG có các group được đặt id theo định dạng 'slide_1', 'slide_2'...",
        )

    return True, ""
