"""Word (.docx) to PDF conversion.

Backends, tried in this order when ``CONVERTER_BACKEND=auto`` (the default):

1. ``libreoffice``  - headless LibreOffice (``soffice`` on PATH). Highest fidelity.
2. ``docx2pdf``     - Microsoft Word via the optional ``docx2pdf`` package (Windows/macOS).
3. ``python``       - built-in pure-Python renderer (python-docx + ReportLab).
                      Works everywhere and has no external dependencies.

If a higher-fidelity backend is missing or fails, the next one is tried, so a
conversion always falls back to the pure-Python renderer.
"""
import io
import logging
import os
import shutil
import subprocess
import sys
import tempfile
from pathlib import Path
from xml.sax.saxutils import escape, quoteattr

import docx
from docx.document import Document as DocxDocument
from docx.table import Table, _Cell
from docx.text.hyperlink import Hyperlink
from docx.text.paragraph import Paragraph as DocxParagraph
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_JUSTIFY, TA_LEFT, TA_RIGHT
from reportlab.lib.styles import ParagraphStyle
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (Image, PageBreak, Paragraph, SimpleDocTemplate,
                                Spacer, Table as RLTable, TableStyle)

log = logging.getLogger(__name__)

BACKENDS = ("auto", "libreoffice", "docx2pdf", "python")
LIBREOFFICE_TIMEOUT = 120

# --------------------------------------------------------------------------
# Backend detection and dispatch
# --------------------------------------------------------------------------


def _find_libreoffice():
    return shutil.which("soffice") or shutil.which("libreoffice")


def _docx2pdf_usable():
    """True when the optional docx2pdf package and Microsoft Word are present."""
    try:
        import docx2pdf  # noqa: F401
    except Exception:
        return False
    if sys.platform == "win32":
        try:
            import winreg
            winreg.CloseKey(winreg.OpenKey(winreg.HKEY_CLASSES_ROOT, "Word.Application"))
            return True
        except OSError:
            return False
    return sys.platform == "darwin" and os.path.isdir("/Applications/Microsoft Word.app")


def available_backends():
    """Names of the backends usable on this machine (``python`` is always there)."""
    found = []
    if _find_libreoffice():
        found.append("libreoffice")
    if _docx2pdf_usable():
        found.append("docx2pdf")
    found.append("python")
    return found


def _convert_libreoffice(docx_path, pdf_path):
    exe = _find_libreoffice()
    if not exe:
        raise RuntimeError("LibreOffice (soffice) not found on PATH")
    with tempfile.TemporaryDirectory() as tmp:
        profile = Path(tmp, "profile").as_uri()  # private profile avoids lock clashes
        out_dir = Path(tmp, "out")
        out_dir.mkdir()
        subprocess.run(
            [exe, f"-env:UserInstallation={profile}", "--headless",
             "--convert-to", "pdf", "--outdir", str(out_dir), str(docx_path)],
            check=True, capture_output=True, timeout=LIBREOFFICE_TIMEOUT,
        )
        produced = out_dir / (Path(docx_path).stem + ".pdf")
        if not produced.exists():
            raise RuntimeError("LibreOffice did not produce a PDF")
        shutil.move(str(produced), pdf_path)


def _convert_docx2pdf(docx_path, pdf_path):
    from docx2pdf import convert
    try:  # Word is driven through COM, which needs initialising per thread
        import pythoncom
        pythoncom.CoInitialize()
    except ImportError:
        pythoncom = None
    try:
        convert(str(docx_path), str(pdf_path))
    finally:
        if pythoncom:
            pythoncom.CoUninitialize()
    if not os.path.exists(pdf_path):
        raise RuntimeError("Word did not produce a PDF")


def convert_docx_to_pdf(docx_path, pdf_path, backend=None):
    """Convert ``docx_path`` to ``pdf_path`` and return the backend actually used."""
    backend = (backend or os.environ.get("CONVERTER_BACKEND", "auto")).lower()
    if backend not in BACKENDS:
        raise ValueError(f"Unknown backend {backend!r}; choose from {', '.join(BACKENDS)}")

    if backend == "auto":
        candidates = available_backends()
    elif backend == "python":
        candidates = ["python"]
    else:
        candidates = [backend, "python"]

    for name in candidates:
        try:
            if name == "libreoffice":
                _convert_libreoffice(docx_path, pdf_path)
            elif name == "docx2pdf":
                _convert_docx2pdf(docx_path, pdf_path)
            else:
                _convert_python(docx_path, pdf_path)
            return name
        except Exception as exc:
            if name == "python":
                raise
            log.warning("Backend %s failed (%s); falling back", name, exc)
            if os.path.exists(pdf_path):
                os.remove(pdf_path)
    raise RuntimeError("No conversion backend succeeded")  # pragma: no cover


# --------------------------------------------------------------------------
# Pure-Python renderer
# --------------------------------------------------------------------------

_FONT_CANDIDATES = [  # (regular, bold, italic, bold-italic)
    ("C:/Windows/Fonts/arial.ttf", "C:/Windows/Fonts/arialbd.ttf",
     "C:/Windows/Fonts/ariali.ttf", "C:/Windows/Fonts/arialbi.ttf"),
    ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-Oblique.ttf",
     "/usr/share/fonts/truetype/dejavu/DejaVuSans-BoldOblique.ttf"),
    ("/Library/Fonts/Arial.ttf", "/Library/Fonts/Arial Bold.ttf",
     "/Library/Fonts/Arial Italic.ttf", "/Library/Fonts/Arial Bold Italic.ttf"),
]
_font_cache = {}


def _body_fonts():
    """Return (regular, bold) font names; a Unicode TrueType family if one is installed."""
    if "fonts" not in _font_cache:
        fonts = ("Helvetica", "Helvetica-Bold")
        for files in _FONT_CANDIDATES:
            if all(os.path.exists(f) for f in files):
                try:
                    names = ("W2P", "W2P-Bold", "W2P-Italic", "W2P-BoldItalic")
                    for name, path in zip(names, files):
                        pdfmetrics.registerFont(TTFont(name, path))
                    pdfmetrics.registerFontFamily(
                        "W2P", normal=names[0], bold=names[1], italic=names[2], boldItalic=names[3])
                    fonts = ("W2P", "W2P-Bold")
                    break
                except Exception as exc:  # unreadable font: try the next family
                    log.debug("Font registration failed: %s", exc)
        _font_cache["fonts"] = fonts
    return _font_cache["fonts"]


_ALIGN = {0: TA_LEFT, 1: TA_CENTER, 2: TA_RIGHT, 3: TA_JUSTIFY}
# kind -> (size, leading, space_before, space_after, bold, colour)
_KINDS = {
    "title": (24, 28, 0, 14, True, "#0d6efd"),
    "subtitle": (12, 16, 0, 12, False, "#6c757d"),
    "h1": (20, 24, 18, 12, True, "#0d6efd"),
    "h2": (16, 20, 14, 10, True, "#2b3035"),
    "h3": (14, 18, 12, 8, True, "#495057"),
    "h4": (12, 16, 10, 6, True, "#495057"),
    "body": (10, 14, 0, 6, False, "#212529"),
    "list": (10, 14, 0, 3, False, "#212529"),
}


def _kind_of(paragraph):
    name = (paragraph.style.name if paragraph.style is not None else "") or ""
    name = name.lower()
    if name == "title":
        return "title"
    if name == "subtitle":
        return "subtitle"
    if name.startswith("heading"):
        level = name.replace("heading", "").strip()
        return {"1": "h1", "2": "h2", "3": "h3"}.get(level, "h4")
    return "body"


class _Renderer:
    def __init__(self, document):
        self.doc = document
        self.regular, self.bold = _body_fonts()
        self._styles = {}
        self.counters = {}
        section = document.sections[0]
        page_w = section.page_width.pt if section.page_width else 612.0
        page_h = section.page_height.pt if section.page_height else 792.0
        margin = lambda v: min(max(v.pt, 36.0), 144.0) if v is not None else 54.0  # noqa: E731
        self.margins = (margin(section.left_margin), margin(section.right_margin),
                        margin(section.top_margin), margin(section.bottom_margin))
        self.page_size = (page_w, page_h)
        self.width = max(page_w - self.margins[0] - self.margins[1], 200.0)
        self.height = max(page_h - self.margins[2] - self.margins[3], 200.0)

    # -- styles --------------------------------------------------------
    def style(self, kind, align, indent=0, compact=False):
        key = (kind, align, indent, compact)
        if key not in self._styles:
            size, leading, before, after, bold, colour = _KINDS[kind]
            if compact:
                size, leading, before, after = size * 0.95, leading * 0.95, 0, 1
            italic = kind == "subtitle"
            font = self.bold if bold else self.regular
            if italic and font == "Helvetica":
                font = "Helvetica-Oblique"
            self._styles[key] = ParagraphStyle(
                f"{kind}-{align}-{indent}-{int(compact)}", fontName=font, fontSize=size,
                leading=leading, textColor=colors.HexColor(colour), alignment=_ALIGN.get(align, TA_LEFT),
                spaceBefore=before, spaceAfter=after, leftIndent=indent,
                bulletIndent=max(indent - 14, 0), bulletFontName=self.regular, bulletFontSize=size)
        return self._styles[key]

    # -- inline content ------------------------------------------------
    @staticmethod
    def _run_markup(run):
        text = escape(run.text).replace("\r", "").replace("\n", "<br/>").replace("\t", "&nbsp;" * 4)
        if not text:
            return ""
        open_tags, close_tags = "", ""
        for flag, tag in ((run.bold, "b"), (run.italic, "i"), (run.underline, "u")):
            if flag:
                open_tags += f"<{tag}>"
                close_tags = f"</{tag}>" + close_tags
        if run.font.superscript:
            open_tags, close_tags = open_tags + "<super>", "</super>" + close_tags
        elif run.font.subscript:
            open_tags, close_tags = open_tags + "<sub>", "</sub>" + close_tags
        attrs = ""
        try:
            if run.font.color is not None and run.font.color.rgb is not None:
                attrs += f' color="#{run.font.color.rgb}"'
        except Exception:
            pass
        if run.font.size:
            attrs += f' size="{run.font.size.pt:g}"'
        if attrs:
            open_tags += f"<font{attrs}>"
            close_tags = "</font>" + close_tags
        return f"{open_tags}{text}{close_tags}"

    def _inline(self, paragraph, max_width):
        """Return (markup, images, page_break_after) for a paragraph."""
        markup, images, page_break = [], [], False
        for item in paragraph.iter_inner_content():
            runs = item.runs if isinstance(item, Hyperlink) else [item]
            inner = ""
            for run in runs:
                inner += self._run_markup(run)
                images.extend(self._images(run, max_width))
                if run._r.xpath('.//w:br[@w:type="page"]'):
                    page_break = True
            if isinstance(item, Hyperlink) and inner:
                url = item.url or ""
                if url.lower().startswith(("http://", "https://", "mailto:")):
                    inner = f'<a href={quoteattr(url)} color="#0d6efd"><u>{inner}</u></a>'
            markup.append(inner)
        return "".join(markup), images, page_break

    def _images(self, run, max_width):
        flowables = []
        for rid in run._r.xpath(".//a:blip/@r:embed"):
            try:
                blob = self.doc.part.related_parts[rid].blob
                extent = run._r.xpath(".//wp:extent")
                if extent:
                    width, height = int(extent[0].get("cx")) / 12700, int(extent[0].get("cy")) / 12700
                else:
                    width = height = None
                img = Image(io.BytesIO(blob))
                width = width or img.imageWidth
                height = height or img.imageHeight
                scale = min(1.0, max_width / width, (self.height * 0.9) / height)
                img.drawWidth, img.drawHeight = width * scale, height * scale
                flowables.append(img)
            except Exception as exc:  # unsupported format (e.g. EMF) is skipped, not fatal
                log.warning("Skipping image: %s", exc)
        return flowables

    # -- lists -----------------------------------------------------------
    def _numpr(self, paragraph):
        pPr = paragraph._p.pPr
        if pPr is not None and pPr.numPr is not None:
            return pPr.numPr
        style = paragraph.style
        while style is not None:
            pPr = style.element.pPr
            if pPr is not None and pPr.numPr is not None:
                return pPr.numPr
            style = style.base_style
        return None

    def _list_info(self, paragraph):
        """Return (num_id, level, is_bullet) or None when the paragraph is not a list item."""
        numpr = self._numpr(paragraph)
        name = ((paragraph.style.name if paragraph.style is not None else "") or "").lower()
        if numpr is None or numpr.numId is None or numpr.numId.val == 0:
            if name.startswith("list bullet"):
                return ("style-bullet", 0, True)
            if name.startswith("list number"):
                return ("style-number", 0, False)
            return None
        num_id = numpr.numId.val
        level = numpr.ilvl.val if numpr.ilvl is not None else 0
        is_bullet = not name.startswith("list number")
        try:
            root = self.doc.part.numbering_part.element
            abstract = root.xpath(f'w:num[@w:numId="{num_id}"]/w:abstractNumId/@w:val')
            fmt = root.xpath(
                f'w:abstractNum[@w:abstractNumId="{abstract[0]}"]/w:lvl[@w:ilvl="{level}"]/w:numFmt/@w:val')
            is_bullet = bool(fmt) and fmt[0] in ("bullet", "none")
        except Exception:
            pass
        return (num_id, level, is_bullet)

    def _bullet_text(self, info):
        num_id, level, is_bullet = info
        if is_bullet:
            return "\u2022"
        counts = self.counters.setdefault(num_id, [])
        del counts[level + 1:]
        counts.extend([0] * (level + 1 - len(counts)))
        counts[level] += 1
        return f"{counts[level]}."

    # -- block level -----------------------------------------------------
    def paragraph(self, paragraph, width, compact=False):
        info = self._list_info(paragraph)
        if info is None:
            self.counters.clear()
        markup, images, page_break = self._inline(paragraph, width)
        fmt = paragraph.paragraph_format
        flowables = []
        if fmt.page_break_before:
            flowables.append(PageBreak())
        align = int(paragraph.alignment) if paragraph.alignment is not None else (
            int(paragraph.style.paragraph_format.alignment)
            if paragraph.style is not None and paragraph.style.paragraph_format.alignment is not None else 0)
        if info:
            indent = 18 * (info[1] + 1) + 6
            style = self.style("list", align, indent, compact)
            bullet = self._bullet_text(info)
        else:
            style = self.style(_kind_of(paragraph), align, 0, compact)
            bullet = None
        if markup.strip():
            flowables.append(Paragraph(markup, style, bulletText=bullet))
        elif not images:
            flowables.append(Spacer(1, style.leading * 0.6))
        flowables.extend(images)
        if page_break:
            flowables.append(PageBreak())
        return flowables

    def table(self, table, width):
        self.counters.clear()
        grid = [c.w.pt for c in table._tbl.tblGrid.gridCol_lst if c.w is not None]
        rows = table._tbl.tr_lst
        ncols = max([len(grid)] + [sum(tc.grid_span for tc in tr.tc_lst) for tr in rows])
        if len(grid) != ncols or sum(grid) <= 0:
            grid = [width / ncols] * ncols
        elif sum(grid) > width:
            grid = [w * width / sum(grid) for w in grid]
        data, spans = [], []
        for r, tr in enumerate(rows):
            row, col = [], 0
            for tc in tr.tc_lst:
                span = tc.grid_span
                cell_width = max(sum(grid[col:col + span]) - 10, 20)
                cell = _Cell(tc, table)
                content = []
                for block in _iter_blocks(cell):
                    if isinstance(block, DocxParagraph):
                        content.extend(self.paragraph(block, cell_width, compact=True))
                    else:
                        content.append(self.table(block, cell_width))
                row.append(content or "")
                row.extend([""] * (span - 1))
                if span > 1:
                    spans.append(("SPAN", (col, r), (col + span - 1, r)))
                col += span
            row.extend([""] * (ncols - len(row)))
            data.append(row)
        if not data:
            return Spacer(1, 1)
        rl_table = RLTable(data, colWidths=grid, repeatRows=1)
        commands = [
            ("GRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#ced4da")),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
            ("TOPPADDING", (0, 0), (-1, -1), 4), ("BOTTOMPADDING", (0, 0), (-1, -1), 4),
            ("LEFTPADDING", (0, 0), (-1, -1), 5), ("RIGHTPADDING", (0, 0), (-1, -1), 5),
            ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#f1f3f5")),
        ] + spans
        commands += [("BACKGROUND", (0, i), (-1, i), colors.HexColor("#f8f9fa"))
                     for i in range(2, len(data), 2)]
        rl_table.setStyle(TableStyle(commands))
        return rl_table


def _iter_blocks(parent):
    """Yield paragraphs and tables of a document or table cell in document order."""
    element = parent.element.body if isinstance(parent, DocxDocument) else parent._tc
    for child in element.iterchildren():
        if child.tag.endswith("}p"):
            yield DocxParagraph(child, parent)
        elif child.tag.endswith("}tbl"):
            yield Table(child, parent)


def _convert_python(docx_path, pdf_path):
    document = docx.Document(str(docx_path))
    renderer = _Renderer(document)
    story = []
    for block in _iter_blocks(document):
        if isinstance(block, DocxParagraph):
            story.extend(renderer.paragraph(block, renderer.width))
        else:
            story.append(renderer.table(block, renderer.width))
            story.append(Spacer(1, 10))
    if not story:
        story.append(Spacer(1, 1))  # an empty document still yields a valid one-page PDF
    left, right, top, bottom = renderer.margins
    SimpleDocTemplate(
        str(pdf_path), pagesize=renderer.page_size, leftMargin=left, rightMargin=right,
        topMargin=top, bottomMargin=bottom, title=Path(docx_path).stem, author="",
    ).build(story)
