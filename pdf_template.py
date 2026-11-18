"""Fill text placeholders without rebuilding a PDF's existing page layout."""
from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

import pymupdf as fitz


@dataclass
class Field:
    page: int
    name: str
    rect: tuple
    baseline: float
    size: float
    color: int
    font: str
    left: float
    right: float
    align: str


def _font_key(name):
    return re.sub(r"[^a-z0-9]", "", name.split("+")[-1].lower())


def _edges(page):
    vertical, horizontal = [], []
    for drawing in page.get_drawings():
        for item in drawing["items"]:
            segments = []
            if item[0] == "l":
                segments.append((item[1], item[2]))
            elif item[0] == "re":
                r = item[1]
                segments.extend(((r.tl, r.tr), (r.tr, r.br), (r.br, r.bl), (r.bl, r.tl)))
            for a, b in segments:
                if abs(a.x - b.x) < 0.5:
                    vertical.append((a.x, min(a.y, b.y), max(a.y, b.y)))
                if abs(a.y - b.y) < 0.5:
                    horizontal.append((a.y, min(a.x, b.x), max(a.x, b.x)))
    return vertical, horizontal


def _field_width(page, rect, chars, start, end, edges):
    """Use visible table cells where possible; otherwise keep the text's origin."""
    vertical, horizontal = edges
    cy, cx = (rect.y0 + rect.y1) / 2, (rect.x0 + rect.x1) / 2
    left = [x for x, y0, y1 in vertical if y0 <= cy <= y1 and x <= rect.x0]
    right = [x for x, y0, y1 in vertical if y0 <= cy <= y1 and x >= rect.x1]
    above = [y for y, x0, x1 in horizontal if x0 <= cx <= x1 and y <= rect.y0]
    below = [y for y, x0, x1 in horizontal if x0 <= cx <= x1 and y >= rect.y1]
    if left and right and above and below:
        x0, x1 = max(left), min(right)
        if abs(cx - (x0 + x1) / 2) <= 5:
            return x0 + 3, x1 - 3, "center"
        if x1 - rect.x1 <= 8:
            return x0 + 3, x1 - 3, "right"
        return rect.x0, x1 - 3, "left"

    following = [c["bbox"][0] for c, _ in chars[end:]
                 if c["c"].strip() and c["bbox"][0] >= rect.x1 - 0.1]
    right_edge = min(following) - 2 if following else page.rect.x1 - 12
    other_text = "".join(c["c"] for c, _ in chars[:start] + chars[end:]).strip()
    if not other_text and rect.x0 > page.rect.width * 0.6:
        return page.rect.x0 + 12, rect.x1, "right"
    return rect.x0, max(rect.x1, right_edge), "left"


class PDFTemplate:
    """Prepare the original page once, then insert each row's field values.

    Static text, images, vector lines and page geometry are retained. Input must
    contain selectable {field} text, not a scan or an already-filled notice.
    """

    def __init__(self, pdf_bytes: bytes, columns):
        self.fields = []
        self.fonts = {}
        self._fallback = None
        font_xrefs = set()
        column_aliases = {}
        for column in columns:
            alias = unicodedata.normalize("NFKC", column)
            if alias in column_aliases and column_aliases[alias] != column:
                raise ValueError("Excel 存在规范化后重名的列，请避免同时使用全角、兼容字符和普通字符命名同一字段")
            column_aliases[alias] = column
        try:
            doc = fitz.open(stream=pdf_bytes, filetype="pdf")
        except Exception as exc:
            raise ValueError("PDF 模板无法读取，请从 Word 重新导出 PDF") from exc
        with doc:
            if doc.needs_pass:
                raise ValueError("PDF 模板不能加密，请导出不带密码的版本")
            if not 0 < len(doc) <= 50:
                raise ValueError("PDF 模板页数需为 1–50 页")
            if doc.get_sigflags() > 0:
                raise ValueError("请使用未数字签名的 PDF 模板；修改签名文件会使签名失效")
            for page_index, page in enumerate(doc):
                if any(annotation.type[0] == fitz.PDF_ANNOT_REDACT for annotation in page.annots() or []):
                    raise ValueError("PDF 模板包含尚未应用的涂黑标记，请先处理后重新导出")
                if page.rotation:
                    raise ValueError("请导出未旋转页面的 PDF 模板")
                for xref, _, _, name, *_ in page.get_fonts(full=True):
                    key = _font_key(name)
                    if xref in font_xrefs or not xref:
                        continue
                    font_xrefs.add(xref)
                    try:
                        buffer = doc.extract_font(xref)[3]
                        if buffer:
                            self.fonts.setdefault(key, []).append(fitz.Font(fontbuffer=buffer))
                    except Exception:
                        # Some PDF fonts cannot be embedded again. Values use a
                        # compatible Unicode font, reported to the caller.
                        pass
                edges = _edges(page)
                first_field = len(self.fields)
                flags = fitz.TEXTFLAGS_RAWDICT & ~fitz.TEXT_PRESERVE_IMAGES
                for block in page.get_text("rawdict", flags=flags)["blocks"]:
                    for line in block.get("lines", []):
                        chars = [(char, span) for span in line["spans"] for char in span["chars"]]
                        text = "".join(char["c"] for char, _ in chars)
                        matches = list(re.finditer(r"\{([^{}\r\n]+)\}", text))
                        if matches and line.get("dir", (1, 0)) != (1, 0):
                            raise ValueError("PDF 字段需为水平文字，暂不支持旋转或竖排字段")
                        for match in matches:
                            raw_name = match.group(1)
                            name = column_aliases.get(unicodedata.normalize("NFKC", raw_name))
                            if name is None:
                                raise ValueError(f"PDF 模板字段 {{{raw_name}}} 在 Excel 中没有对应列")
                            field_chars = chars[match.start():match.end()]
                            rect = fitz.Rect(field_chars[0][0]["bbox"])
                            for char, _ in field_chars[1:]:
                                rect |= fitz.Rect(char["bbox"])
                            char, span = field_chars[0]
                            x0, x1, align = _field_width(page, rect, chars, match.start(), match.end(), edges)
                            self.fields.append(Field(page_index, name, tuple(rect), char["origin"][1],
                                                     span["size"], span["color"], _font_key(span["font"]),
                                                     x0, x1, align))
                            # Inset a little to avoid touching adjacent label glyphs.
                            removal = fitz.Rect(rect.x0 + 0.01, rect.y0 + 0.01,
                                                rect.x1 - 0.01, rect.y1 - 0.01)
                            page.add_redact_annot(removal, fill=False, cross_out=False)
                if len(self.fields) > first_field:
                    # Remove placeholder text ONLY: never remove logos, stamps,
                    # table strokes, or draw an opaque rectangle over artwork.
                    page.apply_redactions(images=0, graphics=0, text=0)
            if not self.fields:
                raise ValueError("PDF 中找不到 {字段名}。请导出带占位符、文字可选中的原始模板，不要上传已生成的通知书或扫描件")
            self.base_pdf = doc.tobytes(garbage=4, deflate=True)

    def _choose_font(self, field, text):
        for font in self.fonts.get(field.font, []):
            if all(font.has_glyph(ord(c)) for c in text):
                return font, False
        if self._fallback is None:
            self._fallback = fitz.Font("china-s")
        if not all(self._fallback.has_glyph(ord(c)) for c in text):
            raise ValueError(f"字段 {{{field.name}}} 包含当前字体不支持的字符，请去除特殊符号")
        return self._fallback, True

    def _plans(self, mapping: dict):
        plans, substitutions = [], set()
        for field in self.fields:
            if field.name not in mapping:
                raise ValueError(f"缺少字段 {{{field.name}}} 的数据")
            text = str(mapping[field.name] if mapping[field.name] is not None else "")
            if not text:
                continue
            if any(c in text for c in "\n\r\t"):
                raise ValueError(f"字段 {{{field.name}}} 含换行或制表符，请使用单行内容")
            font, substituted = self._choose_font(field, text)
            width = font.text_length(text, fontsize=field.size)
            available = field.right - field.left
            if available <= 0:
                raise ValueError(f"字段 {{{field.name}}} 的可用空间无效，请调整 PDF 模板")
            # Never silently alter font size or let text overlap another field.
            if width > available + 0.1:
                raise ValueError(f"字段 {{{field.name}}} 内容过长，无法按原字号放入模板。请缩短内容，或在 Word 中加宽字段区域后重新导出 PDF")
            if field.align == "center":
                x = field.left + (available - width) / 2
            elif field.align == "right":
                x = field.right - width
            else:
                x = field.left
            if substituted:
                substitutions.add(field.name)
            plans.append((field, font, text, x))
        return plans, substitutions

    def render(self, mapping: dict) -> tuple[bytes, list[str]]:
        plans, substitutions = self._plans(mapping)
        with fitz.open(stream=self.base_pdf, filetype="pdf") as doc:
            # TextWriter reuses native Font objects instead of re-embedding the
            # same large CJK font and building glyph-width tables for each field.
            for field, font, text, x in plans:
                page = doc[field.page]
                # Separate fields avoid PDF-device relative-position rounding
                # between adjacent table cells while still sharing native fonts.
                writer = fitz.TextWriter(page.rect)
                writer.append((x, field.baseline), text, font=font, fontsize=field.size)
                color = ((field.color >> 16 & 255) / 255,
                         (field.color >> 8 & 255) / 255,
                         (field.color & 255) / 255)
                writer.write_text(page, color=color, overlay=True)
            doc.subset_fonts()
            return doc.tobytes(garbage=4, deflate=True), sorted(substitutions)
