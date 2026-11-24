"""Render / public-demo mode."""
import os
import time

from app import DEMO_NOTICE, create_app


def make(tmp_path, **overrides):
    config = {
        "TESTING": True,
        "UPLOAD_FOLDER": str(tmp_path / "uploads"),
        "OUTPUT_FOLDER": str(tmp_path / "outputs"),
        "DATABASE": str(tmp_path / "t.db"),
        "CONVERTER_BACKEND": "python",
    }
    config.update(overrides)
    return create_app(config)


def test_render_env_sets_secure_cookies(tmp_path, monkeypatch):
    monkeypatch.setenv("RENDER", "1")
    app = make(tmp_path)
    assert app.config["SESSION_COOKIE_SECURE"] is True
    assert app.test_client().get("/", headers={"X-Forwarded-Proto": "https"}).status_code == 200


def test_demo_banner_only_in_demo_mode(tmp_path, monkeypatch):
    monkeypatch.setenv("DEMO_DATA", "1")
    html = make(tmp_path).test_client().get("/").get_data(as_text=True)
    assert DEMO_NOTICE.replace("'", "&#39;") in html
    monkeypatch.delenv("DEMO_DATA")
    html = make(tmp_path).test_client().get("/").get_data(as_text=True)
    assert "Public demo" not in html


def test_demo_mode_prunes_old_outputs(tmp_path, monkeypatch, sample_docx):
    monkeypatch.setenv("DEMO_DATA", "1")
    app = make(tmp_path)
    old = tmp_path / "outputs" / ("0" * 32 + ".pdf")
    old.write_bytes(b"%PDF-old")
    stale = time.time() - 7200
    os.utime(old, (stale, stale))
    with open(sample_docx, "rb") as fh:
        resp = app.test_client().post("/convert", data={"word_file": (fh, "a.docx")},
                                      content_type="multipart/form-data")
    assert resp.status_code == 200
    assert not old.exists()
    assert os.listdir(tmp_path / "uploads") == []