"""
svg_fixer.py
------------
Module kiểm tra và tự động sửa lỗi phổ biến trong mã SVG do AI tạo ra.

Các lỗi được xử lý:
  1. Markdown code fence  — AI bọc SVG trong ```svg ... ```
  2. Text thừa xung quanh — trích xuất đúng thẻ <svg>...</svg>
  3. Thiếu xmlns          — thêm xmlns="http://www.w3.org/2000/svg"
  4. Thiếu viewBox        — thêm viewBox="0 0 1280 720"
  5. width/height không khớp — chuẩn hóa theo viewBox
  6. XML lỗi cú pháp      — dùng lxml recover mode để phục hồi
  7. id slide không chuẩn — chuẩn hóa slide1 / slide-1 / Slide_1 → slide_1
  8. id slide trùng lặp   — tự động đánh số lại
  9. Thiếu thẻ đóng       — lxml tự sửa qua recover
 10. Encoding đặc biệt    — đảm bảo UTF-8
"""

import re
from lxml import etree
from bs4 import BeautifulSoup


# ─── Constant ────────────────────────────────────────────────────────────────
SVG_NAMESPACE   = "http://www.w3.org/2000/svg"
DEFAULT_VIEWBOX = "0 0 1280 720"
DEFAULT_WIDTH   = "1280"
DEFAULT_HEIGHT  = "720"


# ─── Kiểu kết quả trả về ─────────────────────────────────────────────────────
class FixResult:
    """Kết quả sau khi chạy fixer."""
    def __init__(self):
        self.fixed_svg: str = ""          # SVG đã được sửa
        self.fixes: list[str] = []        # Danh sách các thay đổi đã thực hiện
        self.warnings: list[str] = []     # Cảnh báo (không sửa được tự động)
        self.errors: list[str] = []       # Lỗi nghiêm trọng
        self.slide_count: int = 0         # Số slide tìm thấy sau khi sửa
        self.success: bool = False

    def add_fix(self, msg: str):
        self.fixes.append(msg)

    def add_warning(self, msg: str):
        self.warnings.append(msg)

    def add_error(self, msg: str):
        self.errors.append(msg)


# ─── Bước 1: Tách SVG ra khỏi text thừa ────────────────────────────────────
def _strip_markdown_fences(text: str) -> str:
    """
    Xóa markdown code fence nếu AI bọc SVG trong:
      ```svg ... ``` hoặc ``` ... ``` hoặc `...`
    """
    # Dạng ```svg ... ``` hoặc ```xml ... ```
    match = re.search(
        r"```(?:svg|xml|html)?\s*\n?([\s\S]+?)\n?```",
        text,
        re.IGNORECASE,
    )
    if match:
        return match.group(1).strip()

    # Dạng backtick đơn
    match = re.search(r"`([\s\S]+)`", text)
    if match:
        return match.group(1).strip()

    return text


def _extract_svg_tag(text: str) -> tuple[str, bool]:
    """
    Trích xuất nội dung từ thẻ <svg ...> đến </svg>.
    Trả về (svg_string, was_extracted).
    """
    # Tìm thẻ mở <svg (có thể có namespace, attributes...)
    start = text.find("<svg")
    if start == -1:
        # Thử tìm dạng chữ thường/hoa khác
        lower = text.lower()
        start = lower.find("<svg")

    if start == -1:
        return text, False

    # Tìm thẻ đóng </svg> tương ứng (từ cuối về)
    end = text.rfind("</svg>")
    if end == -1:
        end = text.lower().rfind("</svg>")

    if end == -1:
        # Không có thẻ đóng — trả về từ <svg> đến hết
        return text[start:].strip(), True

    return text[start : end + 6].strip(), (start > 0 or end < len(text) - 6)


# ─── Bước 2: Sửa bằng lxml (recover mode) ────────────────────────────────────
def _parse_with_recovery(svg_text: str) -> tuple[etree._Element | None, bool]:
    """
    Parse SVG bằng lxml. Nếu lỗi, thử lại với recover=True.
    Trả về (root_element, was_recovered).
    """
    parser_strict = etree.XMLParser(recover=False, encoding="utf-8")
    try:
        root = etree.fromstring(svg_text.encode("utf-8"), parser_strict)
        return root, False
    except etree.XMLSyntaxError:
        pass

    # Thử recover
    parser_recover = etree.XMLParser(recover=True, encoding="utf-8")
    try:
        root = etree.fromstring(svg_text.encode("utf-8"), parser_recover)
        return root, True  # True = đã phải dùng recover
    except Exception:
        return None, False


# ─── Bước 3: Sửa namespace, viewBox, width/height ────────────────────────────
def _fix_svg_attributes(root: etree._Element, result: FixResult) -> etree._Element:
    """Sửa các thuộc tính cơ bản của thẻ <svg> gốc."""
    tag_local = etree.QName(root.tag).localname if "}" in root.tag else root.tag

    # ── xmlns ──────────────────────────────────────────────────────────────
    # lxml tự xử lý namespace, nhưng kiểm tra xem có đúng không
    current_ns = root.nsmap.get(None, "")
    if current_ns != SVG_NAMESPACE:
        # Không thể đổi nsmap trực tiếp trong lxml — đánh dấu cảnh báo
        result.add_warning(
            f"Namespace SVG không chuẩn ({current_ns or 'thiếu'}). "
            "Sẽ được thêm khi render lại."
        )

    # ── viewBox ────────────────────────────────────────────────────────────
    if not root.get("viewBox"):
        root.set("viewBox", DEFAULT_VIEWBOX)
        result.add_fix("✔ Thêm viewBox=\"0 0 1280 720\"")

    # ── width / height ─────────────────────────────────────────────────────
    vb = root.get("viewBox", DEFAULT_VIEWBOX)
    try:
        parts = vb.split()
        vb_w, vb_h = parts[2], parts[3]
    except (IndexError, ValueError):
        vb_w, vb_h = DEFAULT_WIDTH, DEFAULT_HEIGHT

    cur_w = root.get("width", "")
    cur_h = root.get("height", "")

    if cur_w and cur_w != vb_w:
        root.set("width", vb_w)
        result.add_fix(f"✔ Sửa width: {cur_w} → {vb_w}")
    elif not cur_w:
        root.set("width", vb_w)

    if cur_h and cur_h != vb_h:
        root.set("height", vb_h)
        result.add_fix(f"✔ Sửa height: {cur_h} → {vb_h}")
    elif not cur_h:
        root.set("height", vb_h)

    return root


# ─── Bước 4: Chuẩn hóa id của slide groups ────────────────────────────────
_SLIDE_ID_PATTERNS = [
    # (pattern, description)
    (re.compile(r"^[Ss]lide[-_\s]?(\d+)$"),   "slide_N hoặc slide-N hoặc slideN"),
    (re.compile(r"^[Ss]lide_([A-Za-z]+)$"),    "slide_tên"),
    (re.compile(r"^[Ss](\d+)$"),               "sN"),
    (re.compile(r"^page[-_]?(\d+)$", re.I),    "page_N"),
    (re.compile(r"^frame[-_]?(\d+)$", re.I),   "frame_N"),
]


def _normalize_slide_ids(root: etree._Element, result: FixResult) -> int:
    """
    Tìm và chuẩn hóa các id của thẻ <g> đại diện slide.
    Trả về số slide tìm thấy.
    """
    nsmap  = {"svg": SVG_NAMESPACE}
    # Tìm tất cả thẻ <g> ở cấp 1 bên trong <svg>
    # (hoặc cấp 2 nếu có wrapper)
    all_g = root.findall(".//{%s}g" % SVG_NAMESPACE) or root.findall(".//g")

    # Thu thập các <g> ứng viên có id liên quan đến slide
    candidates = []
    for g in all_g:
        g_id = g.get("id", "")
        if not g_id:
            continue
        # Đã đúng chuẩn slide_N
        if re.match(r"^slide_\d+$", g_id):
            candidates.append((g, g_id, g_id))
            continue
        # Thử khớp các pattern không chuẩn
        for pattern, _ in _SLIDE_ID_PATTERNS:
            m = pattern.match(g_id)
            if m:
                try:
                    n = int(m.group(1))
                except (IndexError, ValueError):
                    n = len(candidates) + 1
                candidates.append((g, g_id, f"slide_{n}"))
                break

    if not candidates:
        return 0

    # Kiểm tra id trùng lặp sau khi chuẩn hóa
    seen_ids: dict[str, int] = {}
    for g, old_id, new_id in candidates:
        if new_id in seen_ids:
            seen_ids[new_id] += 1
            new_id = f"{new_id}_{seen_ids[new_id]}"
        else:
            seen_ids[new_id] = 1

        if old_id != new_id:
            g.set("id", new_id)
            result.add_fix(f"✔ Đổi id <g>: \"{old_id}\" → \"{new_id}\"")

    return len(candidates)


# ─── Bước 5: Kiểm tra nội dung slide ────────────────────────────────────────
def _check_slide_content(root: etree._Element, result: FixResult, slide_count: int):
    """Phát hiện các cảnh báo về nội dung slide."""
    if slide_count == 0:
        result.add_error(
            "Không tìm thấy thẻ <g id='slide_N'> nào. "
            "Kiểm tra lại định dạng id của các group."
        )
        return

    # Kiểm tra slide rỗng
    all_g = root.findall(".//{%s}g" % SVG_NAMESPACE) or root.findall(".//g")
    empty_slides = []
    for g in all_g:
        g_id = g.get("id", "")
        if re.match(r"^slide_\d+$", g_id):
            content = "".join(g.itertext()).strip()
            children_count = len(list(g))
            if children_count == 0 and not content:
                empty_slides.append(g_id)

    if empty_slides:
        result.add_warning(
            f"Slide rỗng (không có nội dung): {', '.join(empty_slides)}"
        )

    # Kiểm tra nếu <defs> có gradient/filter nhưng không ai dùng
    defs = root.find("{%s}defs" % SVG_NAMESPACE) or root.find("defs")
    if defs is not None:
        def_ids = [c.get("id") for c in defs if c.get("id")]
        svg_text = etree.tostring(root, encoding="unicode")
        unused = [d for d in def_ids if f"#{d}" not in svg_text and f"url(#{d})" not in svg_text]
        if unused:
            result.add_warning(
                f"Các định nghĩa <defs> không được dùng: {', '.join(unused)}"
            )


# ─── Bước 6: Render lại SVG sạch ─────────────────────────────────────────────
def _serialize_svg(root: etree._Element) -> str:
    """
    Chuyển cây lxml thành chuỗi SVG.
    Đảm bảo khai báo namespace đúng.
    """
    # Đảm bảo namespace SVG trên root
    nsmap = dict(root.nsmap)
    if None not in nsmap or nsmap[None] != SVG_NAMESPACE:
        # Tạo lại root với đúng nsmap
        new_root = etree.Element(
            "{%s}svg" % SVG_NAMESPACE,
            nsmap={
                None:    SVG_NAMESPACE,
                "xlink": "http://www.w3.org/1999/xlink",
            }
        )
        # Copy attributes
        for k, v in root.attrib.items():
            new_root.set(k, v)
        # Copy children
        for child in root:
            new_root.append(child)
        root = new_root

    svg_bytes = etree.tostring(
        root,
        xml_declaration=False,
        encoding="unicode",
        pretty_print=True,
    )
    return svg_bytes


# ─── Entry point ─────────────────────────────────────────────────────────────
def fix_svg(raw_input: str) -> FixResult:
    """
    Pipeline tự động kiểm tra và sửa mã SVG.

    Args:
        raw_input: Mã SVG thô (có thể chứa text markdown fence, text thừa...)

    Returns:
        FixResult với fixed_svg, danh sách fixes, warnings, errors.
    """
    result = FixResult()

    if not raw_input or not raw_input.strip():
        result.add_error("Mã SVG trống.")
        return result

    text = raw_input.strip()

    # ── Bước 1: Bóc tách markdown / text thừa ─────────────────────────────
    stripped = _strip_markdown_fences(text)
    if stripped != text:
        result.add_fix("✔ Xóa markdown code fence (``` ... ```)")
        text = stripped

    svg_text, was_extracted = _extract_svg_tag(text)
    if was_extracted and "<svg" in svg_text and svg_text != text:
        result.add_fix("✔ Trích xuất thẻ <svg> từ văn bản thừa xung quanh")
    text = svg_text

    if "<svg" not in text.lower():
        result.add_error("Không tìm thấy thẻ <svg> trong mã đầu vào.")
        return result

    # ── Bước 2: Parse XML (có recover) ────────────────────────────────────
    root, was_recovered = _parse_with_recovery(text)
    if root is None:
        result.add_error(
            "Mã SVG bị lỗi XML nghiêm trọng, không thể phục hồi. "
            "Hãy thử copy lại toàn bộ mã từ AI."
        )
        return result

    if was_recovered:
        result.add_fix("✔ Phục hồi XML bị lỗi cú pháp (thẻ không đóng, ký tự đặc biệt...)")

    # ── Bước 3: Sửa attributes cơ bản ─────────────────────────────────────
    root = _fix_svg_attributes(root, result)

    # ── Bước 4: Chuẩn hóa slide id ────────────────────────────────────────
    slide_count = _normalize_slide_ids(root, result)
    result.slide_count = slide_count

    # ── Bước 5: Kiểm tra nội dung ─────────────────────────────────────────
    _check_slide_content(root, result, slide_count)

    # ── Bước 6: Render lại ────────────────────────────────────────────────
    result.fixed_svg = _serialize_svg(root)
    result.success    = len(result.errors) == 0

    if result.success and not result.fixes:
        result.fixes.append("✔ Mã SVG hợp lệ, không cần sửa chữa.")

    return result
