import io

import docx
import pytest
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.shared import Inches, Pt, RGBColor
from PIL import Image

from app import create_app


@pytest.fixture
def sample_docx(tmp_path):
    """A document exercising headings, formatting, lists, a table and an image."""
    document = docx.Document()
    document.add_heading("Quarterly Report", level=1)
    document.add_heading("Overview", level=2)

    para = document.add_paragraph("Plain intro, ")
    para.add_run("bold words").bold = True
    para.add_run(", ")
    para.add_run("italic words").italic = True
    para.add_run(", and ")
    red = para.add_run("red text")
    red.font.color.rgb = RGBColor(220, 53, 69)
    red.font.size = Pt(14)

    centered = document.add_paragraph("Centered line")
    centered.alignment = WD_ALIGN_PARAGRAPH.CENTER

    document.add_paragraph("Apples", style="List Bullet")
    document.add_paragraph("Oranges", style="List Bullet")
    document.add_paragraph("Step one", style="List Number")
    document.add_paragraph("Step two", style="List Number")

    table = document.add_table(rows=3, cols=2)
    table.style = "Table Grid"
    table.rows[0].cells[0].text = "Metric"
    table.rows[0].cells[1].text = "Value"
    table.rows[1].cells[0].text = "Revenue"
    table.rows[1].cells[1].text = "1200"
    table.rows[2].cells[0].merge(table.rows[2].cells[1]).text = "Merged footer"

    buffer = io.BytesIO()
    Image.new("RGB", (120, 60), (30, 120, 220)).save(buffer, "PNG")
    buffer.seek(0)
    document.add_picture(buffer, width=Inches(2))

    document.add_page_break()
    document.add_paragraph("Final page text")

    path = tmp_path / "sample.docx"
    document.save(path)
    return path


@pytest.fixture
def app(tmp_path):
    return create_app({
        "TESTING": True,
        "UPLOAD_FOLDER": str(tmp_path / "uploads"),
        "OUTPUT_FOLDER": str(tmp_path / "outputs"),
        "DATABASE": str(tmp_path / "test.db"),
        "CONVERTER_BACKEND": "python",  # deterministic, independent of installed Office tools
    })


@pytest.fixture
def client(app):
    return app.test_client()
