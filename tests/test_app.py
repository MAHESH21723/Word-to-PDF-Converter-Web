import io
import os
import sqlite3

from pypdf import PdfReader


def upload(client, data, name="report.docx"):
    return client.post("/convert", data={"word_file": (io.BytesIO(data), name)},
                       content_type="multipart/form-data")


def test_index_page_renders(client):
    response = client.get("/")
    assert response.status_code == 200
    assert b"Word2PDF" in response.data


def test_startup_creates_database_and_folders(app):
    assert os.path.isfile(app.config["DATABASE"])
    assert os.path.isdir(app.config["UPLOAD_FOLDER"])
    assert os.path.isdir(app.config["OUTPUT_FOLDER"])


def test_convert_then_download(client, app, sample_docx):
    response = upload(client, sample_docx.read_bytes(), "My Report.docx")
    body = response.get_json()
    assert response.status_code == 200 and body["status"] == "success"
    assert body["filename"] == "My_Report.docx"
    assert body["backend"] == "python"

    download = client.get(f"/download/{body['pdf_filename']}")
    assert download.status_code == 200
    assert download.mimetype == "application/pdf"
    assert "My_Report.pdf" in download.headers["Content-Disposition"]
    assert download.data.startswith(b"%PDF-")
    assert len(PdfReader(io.BytesIO(download.data)).pages) == 2

    # The uploaded Word file must not be kept on disk.
    assert os.listdir(app.config["UPLOAD_FOLDER"]) == []


def test_history_and_metrics_reflect_conversions(client, sample_docx):
    upload(client, sample_docx.read_bytes())
    upload(client, b"not a docx at all", "bad.docx")

    history = client.get("/history").get_json()
    assert [row["status"] for row in history] == ["failed", "success"]

    metrics = client.get("/metrics").get_json()
    assert (metrics["total"], metrics["success"], metrics["failed"]) == (2, 1, 1)
    assert metrics["success_rate"] == 50.0


def test_clear_history_removes_records_and_files(client, app, sample_docx):
    upload(client, sample_docx.read_bytes())
    assert os.listdir(app.config["OUTPUT_FOLDER"])
    assert client.post("/history/clear").get_json()["status"] == "success"
    assert client.get("/history").get_json() == []
    assert os.listdir(app.config["OUTPUT_FOLDER"]) == []


def test_rejects_missing_file(client):
    assert client.post("/convert", data={}).status_code == 400


def test_rejects_empty_filename(client):
    assert upload(client, b"x", "").status_code == 400


def test_rejects_wrong_extension(client):
    response = upload(client, b"hello", "notes.txt")
    assert response.status_code == 400
    assert "Only .docx" in response.get_json()["message"]


def test_rejects_fake_docx_by_content(client, app):
    response = upload(client, b"MZ\x90\x00 pretend executable", "evil.docx")
    assert response.status_code == 422
    assert os.listdir(app.config["UPLOAD_FOLDER"]) == []
    assert os.listdir(app.config["OUTPUT_FOLDER"]) == []


def test_rejects_zip_that_is_not_word(client):
    import zipfile
    buffer = io.BytesIO()
    with zipfile.ZipFile(buffer, "w") as archive:
        archive.writestr("hello.txt", "hi")
    assert upload(client, buffer.getvalue(), "zip.docx").status_code == 422


def test_rejects_oversized_upload(app, client, sample_docx):
    app.config["MAX_CONTENT_LENGTH"] = 1024
    response = upload(client, sample_docx.read_bytes())
    assert response.status_code == 413
    assert "too large" in response.get_json()["message"]


def test_path_traversal_in_upload_name_is_neutralised(client, app, sample_docx):
    response = upload(client, sample_docx.read_bytes(), "../../evil.docx")
    body = response.get_json()
    assert response.status_code == 200
    assert body["filename"] == "evil.docx"
    assert os.path.dirname(os.path.join(app.config["OUTPUT_FOLDER"], body["pdf_filename"])) \
        == app.config["OUTPUT_FOLDER"]


def test_download_rejects_traversal_and_unknown_names(client):
    assert client.get("/download/..%2Fdatabase.db").status_code == 404
    assert client.get("/download/" + "0" * 32 + ".pdf").status_code == 404
    assert client.get("/download/anything.txt").status_code == 404


def test_failed_conversion_is_logged_without_leaking_details(client, app):
    response = upload(client, b"PK\x03\x04garbage", "broken.docx")
    assert response.status_code == 422
    with sqlite3.connect(app.config["DATABASE"]) as conn:
        status, message = conn.execute("SELECT status, error_message FROM conversions").fetchone()
    assert status == "failed" and message
