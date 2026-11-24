import docx
import pytest
from pypdf import PdfReader

import converter


def pdf_text(path):
    return "\n".join(page.extract_text() for page in PdfReader(str(path)).pages)


def test_converts_sample_to_valid_pdf_with_expected_text(sample_docx, tmp_path):
    out = tmp_path / "out.pdf"
    assert converter.convert_docx_to_pdf(sample_docx, out, backend="python") == "python"

    assert out.read_bytes().startswith(b"%PDF-")
    reader = PdfReader(str(out))
    assert len(reader.pages) == 2  # explicit page break before "Final page text"

    text = pdf_text(out)
    for expected in ("Quarterly Report", "Overview", "bold words", "italic words",
                     "Centered line", "Apples", "Revenue", "1200", "Merged footer"):
        assert expected in text
    assert "Final page text" in reader.pages[1].extract_text()


def test_lists_are_rendered_with_markers(sample_docx, tmp_path):
    out = tmp_path / "out.pdf"
    converter.convert_docx_to_pdf(sample_docx, out, backend="python")
    text = pdf_text(out)
    assert "•" in text and "Apples" in text
    assert "1." in text and "2." in text


def test_image_is_embedded(sample_docx, tmp_path):
    out = tmp_path / "out.pdf"
    converter.convert_docx_to_pdf(sample_docx, out, backend="python")
    images = [img for page in PdfReader(str(out)).pages for img in page.images]
    assert len(images) == 1


def test_bold_and_italic_use_distinct_fonts(sample_docx, tmp_path):
    out = tmp_path / "out.pdf"
    converter.convert_docx_to_pdf(sample_docx, out, backend="python")
    fonts = set()
    for page in PdfReader(str(out)).pages:
        fonts.update(page["/Resources"]["/Font"].keys())
    assert len(fonts) >= 3  # regular + bold + italic at minimum


def test_empty_document_still_produces_pdf(tmp_path):
    src = tmp_path / "empty.docx"
    docx.Document().save(src)
    out = tmp_path / "empty.pdf"
    converter.convert_docx_to_pdf(src, out, backend="python")
    assert len(PdfReader(str(out)).pages) == 1


def test_special_characters_are_escaped(tmp_path):
    document = docx.Document()
    document.add_paragraph("5 < 6 & 7 > 3 <b>not bold</b>")
    src = tmp_path / "special.docx"
    document.save(src)
    out = tmp_path / "special.pdf"
    converter.convert_docx_to_pdf(src, out, backend="python")
    assert "5 < 6 & 7 > 3 <b>not bold</b>" in pdf_text(out)


def test_unknown_backend_rejected(sample_docx, tmp_path):
    with pytest.raises(ValueError):
        converter.convert_docx_to_pdf(sample_docx, tmp_path / "x.pdf", backend="nope")


def test_failing_backend_falls_back_to_python(sample_docx, tmp_path, monkeypatch):
    def boom(*_):
        raise RuntimeError("no office here")

    monkeypatch.setattr(converter, "_convert_libreoffice", boom)
    monkeypatch.setattr(converter, "_find_libreoffice", lambda: "soffice")
    out = tmp_path / "out.pdf"
    assert converter.convert_docx_to_pdf(sample_docx, out, backend="auto") == "python"
    assert out.exists()


def test_available_backends_always_includes_python():
    assert "python" in converter.available_backends()
