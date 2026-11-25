# Word to PDF Converter (Flask)

A small Flask web app that converts Word `.docx` documents to PDF. Drop a file in the browser, get a PDF back, and keep a history of past conversions. It runs entirely on your own machine.

## Features

- Drag-and-drop or click-to-browse upload, with progress and success/error states
- Pure-Python conversion that works on Windows, macOS and Linux with no Office install. It keeps headings, paragraphs, bold/italic/underline, colours and sizes, alignment, bullet and numbered lists, tables (including merged cells), images, hyperlinks and page breaks
- Optional higher-fidelity backends, detected automatically and used when present, with a fall back to the pure-Python renderer if they fail:
  - LibreOffice, when `soffice` is on `PATH`
  - Microsoft Word through the optional `docx2pdf` package (`pip install docx2pdf`)
- Conversion history (SQLite), usage metrics, re-download and clear-history
- Upload hardening: extension check, content check (the file must be a real OOXML Word package), size limit, `secure_filename`, random stored names, no path traversal on download, uploaded Word files deleted straight after conversion
- Configuration through environment variables; debug mode is off by default and the server binds to `127.0.0.1`

## Tech stack

Python 3.10+, Flask, python-docx, ReportLab, Pillow, SQLite, Bootstrap 5 and vanilla JavaScript on the front end. Tests use pytest and pypdf.

## Live demo

**Live demo: [https://word-to-pdf-converter-9jz1.onrender.com](https://word-to-pdf-converter-9jz1.onrender.com)** (free hosting: the first visit after a quiet spell can take up to a minute to wake up)

[![Deploy to Render](https://render.com/images/deploy-to-render-button.svg)](https://render.com/deploy?repo=https://github.com/MAHESH21723/Word-to-PDF-Converter-Web)

To host your own copy for free:

1. Sign in to [Render](https://render.com) with GitHub.
2. Click the **Deploy to Render** button above.
3. Review the blueprint (`render.yaml`) and click **Apply**.

Free-plan notes: the service sleeps after about 15 minutes idle, so the first request after a pause takes about a minute. The filesystem is ephemeral, so conversion history and files resets on every restart or redeploy. The demo sets `DEMO_DATA=1`, which shows a "public demo" banner. You need no server or domain of your own, and local use is unchanged (see Setup below).
## Setup

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1        # macOS/Linux: source .venv/bin/activate
pip install -r requirements.txt
```

## Run

```powershell
python app.py
```

Open <http://127.0.0.1:5000>. The `uploads/` and `outputs/` folders and `database.db` are created on first start.

### Configuration (all optional)

| Variable | Default | Meaning |
| --- | --- | --- |
| `SECRET_KEY` | random per start | Flask secret key |
| `DATA_DIR` | project folder | Where `uploads/`, `outputs/` and `database.db` live |
| `MAX_UPLOAD_MB` | `15` | Upload size limit |
| `CONVERTER_BACKEND` | `auto` | `auto`, `libreoffice`, `docx2pdf` or `python` |
| `HOST` / `PORT` | `127.0.0.1` / `5000` | Bind address |
| `FLASK_DEBUG` | off | Set to `1` for debug mode |

## Tests

```powershell
pip install -r requirements-dev.txt
python -m pytest -q
```

The tests build sample `.docx` files in code, convert them, and check the PDF is valid and contains the expected text, then exercise the routes (upload validation, download, history, metrics).

## Project structure

```
app.py                  Flask app factory and routes
converter.py            Backend detection + pure-Python DOCX -> PDF renderer
database.py             SQLite history and metrics
templates/index.html    UI
static/css, static/js   Styles and front-end logic
tests/                  pytest suite (converter + routes)
```

## Limitations

- The pure-Python renderer approximates Word's layout; complex documents (text boxes, charts, tracked changes, headers/footers, EMF/WMF images) are simplified. Use LibreOffice or Word for exact fidelity.
- Text outside the Latin range needs a Unicode TrueType font on the system (Arial or DejaVu Sans is picked up automatically).
- The page loads Bootstrap and Google Fonts from public CDNs, so the browser needs internet access for styling; documents themselves never leave your machine.

## Author

Maheshwaran

## License

MIT. See [LICENSE](LICENSE).
