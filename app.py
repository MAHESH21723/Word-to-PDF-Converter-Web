"""Flask web app that converts Word (.docx) documents to PDF.

Configuration comes from environment variables (all optional):

    SECRET_KEY         Flask secret key (random per start-up if unset)
    DATA_DIR           Where uploads/, outputs/ and database.db live (default: project folder)
    MAX_UPLOAD_MB      Upload size limit in megabytes (default: 15)
    CONVERTER_BACKEND  auto | libreoffice | docx2pdf | python (default: auto)
    HOST / PORT        Bind address (default: 127.0.0.1:5000)
    FLASK_DEBUG        Set to 1 to enable Flask's debug mode (default: off)
"""
import logging
import os
import re
import secrets
import time
import uuid
import zipfile

from flask import Flask, jsonify, render_template, request, send_from_directory
from werkzeug.exceptions import RequestEntityTooLarge
from werkzeug.utils import secure_filename

import converter
import database

log = logging.getLogger(__name__)

STORED_PDF_RE = re.compile(r"^[0-9a-f]{32}\.pdf$")
MAX_UNCOMPRESSED_BYTES = 200 * 1024 * 1024  # guard against zip bombs


class InvalidDocument(ValueError):
    """The uploaded file is not a usable .docx document."""


def validate_docx(path):
    """Check the file *content* is a real .docx (an OOXML zip with a Word body)."""
    if not zipfile.is_zipfile(path):
        raise InvalidDocument("The file is not a valid .docx document.")
    try:
        with zipfile.ZipFile(path) as archive:
            names = set(archive.namelist())
            if "[Content_Types].xml" not in names or "word/document.xml" not in names:
                raise InvalidDocument("The file is not a valid .docx document.")
            if sum(info.file_size for info in archive.infolist()) > MAX_UNCOMPRESSED_BYTES:
                raise InvalidDocument("The document expands to an unreasonable size.")
    except zipfile.BadZipFile:
        raise InvalidDocument("The file is corrupted.") from None


def create_app(test_config=None):
    app = Flask(__name__)
    data_dir = os.environ.get("DATA_DIR") or app.root_path
    app.config.update(
        SECRET_KEY=os.environ.get("SECRET_KEY") or secrets.token_hex(32),
        MAX_CONTENT_LENGTH=int(os.environ.get("MAX_UPLOAD_MB", "15")) * 1024 * 1024,
        UPLOAD_FOLDER=os.path.join(data_dir, "uploads"),
        OUTPUT_FOLDER=os.path.join(data_dir, "outputs"),
        DATABASE=os.path.join(data_dir, "database.db"),
        CONVERTER_BACKEND=None,  # None -> CONVERTER_BACKEND env var / auto
    )
    if test_config:
        app.config.update(test_config)

    os.makedirs(app.config["UPLOAD_FOLDER"], exist_ok=True)
    os.makedirs(app.config["OUTPUT_FOLDER"], exist_ok=True)
    database.init_db(app.config["DATABASE"])

    db = lambda: app.config["DATABASE"]  # noqa: E731

    def error(message, status):
        return jsonify({"status": "error", "message": message}), status

    @app.after_request
    def security_headers(response):
        response.headers.setdefault("X-Content-Type-Options", "nosniff")
        response.headers.setdefault("X-Frame-Options", "DENY")
        return response

    @app.errorhandler(RequestEntityTooLarge)
    def too_large(_):
        limit = app.config["MAX_CONTENT_LENGTH"] // (1024 * 1024)
        return error(f"File is too large. The maximum size is {limit} MB.", 413)

    @app.route("/")
    def index():
        return render_template("index.html")

    @app.route("/convert", methods=["POST"])
    def convert_document():
        upload = request.files.get("word_file")
        if upload is None:
            return error("No file part in request", 400)
        if not upload.filename:
            return error("No selected file", 400)
        if not upload.filename.lower().endswith(".docx"):
            return error("Invalid file type. Only .docx files are allowed", 400)

        safe_name = secure_filename(upload.filename)
        if not safe_name.lower().endswith(".docx"):
            safe_name = "document.docx"  # name was empty or lost its extension when sanitised
        token = uuid.uuid4().hex
        word_path = os.path.join(app.config["UPLOAD_FOLDER"], f"{token}.docx")
        stored_pdf = f"{token}.pdf"
        pdf_path = os.path.join(app.config["OUTPUT_FOLDER"], stored_pdf)
        started = time.time()
        original_size = 0
        succeeded = False

        def record(status, message=None, converted=None, size=None):
            database.add_conversion(
                db(), filename=safe_name, original_size=original_size,
                converted_filename=converted, converted_size=size, status=status,
                error_message=message, conversion_time=round(time.time() - started, 2))

        try:
            upload.save(word_path)
            original_size = os.path.getsize(word_path)
            validate_docx(word_path)
            backend = converter.convert_docx_to_pdf(
                word_path, pdf_path, app.config["CONVERTER_BACKEND"])
            converted_size = os.path.getsize(pdf_path)
            record("success", converted=stored_pdf, size=converted_size)
            succeeded = True
            return jsonify({
                "status": "success",
                "filename": safe_name,
                "pdf_filename": stored_pdf,
                "conversion_time": round(time.time() - started, 2),
                "size": converted_size,
                "backend": backend,
            })
        except InvalidDocument as exc:
            record("failed", str(exc))
            return error(str(exc), 422)
        except Exception:
            log.exception("Conversion failed for %s", safe_name)
            record("failed", "Conversion failed")
            return error("The document could not be converted. It may be corrupted "
                         "or use unsupported features.", 500)
        finally:
            # The uploaded Word file is never kept; drop partial output on failure.
            if os.path.exists(word_path):
                os.remove(word_path)
            if not succeeded and os.path.exists(pdf_path):
                os.remove(pdf_path)

    @app.route("/download/<filename>")
    def download_file(filename):
        if not STORED_PDF_RE.match(filename):
            return error("File not found", 404)
        record = database.get_conversion_by_stored_name(db(), filename)
        if record is None:
            return error("File not found", 404)
        download_name = os.path.splitext(record["filename"])[0] + ".pdf"
        return send_from_directory(
            app.config["OUTPUT_FOLDER"], filename, as_attachment=True,
            download_name=download_name, mimetype="application/pdf")

    @app.route("/history")
    def get_history():
        return jsonify(database.get_all_conversions(db()))

    @app.route("/metrics")
    def get_metrics():
        return jsonify(database.get_metrics(db()))

    @app.route("/history/clear", methods=["POST"])
    def clear_history():
        for folder in (app.config["UPLOAD_FOLDER"], app.config["OUTPUT_FOLDER"]):
            for entry in os.scandir(folder):
                if entry.is_file() and not entry.name.startswith("."):
                    try:
                        os.remove(entry.path)
                    except OSError:
                        log.warning("Could not delete %s", entry.path)
        database.clear_history(db())
        return jsonify({"status": "success", "message": "History and files cleared successfully"})

    return app


if __name__ == "__main__":
    create_app().run(
        debug=os.environ.get("FLASK_DEBUG") == "1",
        host=os.environ.get("HOST", "127.0.0.1"),
        port=int(os.environ.get("PORT", "5000")),
    )
