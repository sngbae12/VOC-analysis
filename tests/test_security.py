# -*- coding: utf-8 -*-
from __future__ import annotations

import shutil
from pathlib import Path

import pytest
from fastapi.testclient import TestClient
from gradio.routes import App
from gradio.utils import get_cache_folder

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def _client() -> TestClient:
    demo = app.build_ui()
    fastapi_app = App.create_app(demo)
    return TestClient(fastapi_app)


def _file_url(path: Path) -> str:
    return "/gradio_api/file=" + str(Path(path).resolve())


def _get_file(client: TestClient, path: Path):
    url = _file_url(path)
    response = client.get(url)
    if response.status_code == 404:
        alt = "/gradio_api/file=" + str(path.resolve()).replace("\\", "/")
        response = client.get(alt)
    return response


@pytest.fixture(scope="module")
def http():
    app.install_download_guard()
    with _client() as client:
        yield client


def test_gradio_temp_csv_blocked_before_and_after_upload(http, tmp_path, logger):
    csv = write_voc_csv(tmp_path / "secret_voc.csv", dates=[20260101, 20260102])
    cache = Path(get_cache_folder())
    cache.mkdir(parents=True, exist_ok=True)
    cached = cache / "gradio_temp_upload.csv"
    shutil.copy(csv, cached)
    assert cached.exists()
    blocked = _get_file(http, cached)
    assert blocked.status_code == 403

    app.do_upload(csv, "file", csv.name, logger)
    blocked_after = _get_file(http, cached)
    assert blocked_after.status_code == 403
    upload_copy = app.UPLOAD_DIR / "uploaded_voc.csv"
    if upload_copy.exists():
        assert _get_file(http, upload_copy).status_code == 403


def test_secrets_source_and_uploads_blocked(http, tmp_path):
    env_path = tmp_path / "fake.env"
    env_path.write_text("OPENAI_API_KEY=sk-not-a-real-key\n", encoding="utf-8")
    assert _get_file(http, env_path).status_code == 403
    assert _get_file(http, app.BASE_DIR / "voc_analysis_app.py").status_code == 403
    assert _get_file(http, app.SAMPLE_CSV).status_code == 403
    git_head = app.BASE_DIR / ".git" / "HEAD"
    if git_head.exists():
        assert _get_file(http, git_head).status_code == 403


def test_registered_png_and_docx_download_ok(http, tmp_path, logger, monkeypatch):
    csv = write_voc_csv(tmp_path / "ok.csv", dates=["2026-01-06", "2026-01-07"])
    app.do_upload(csv, "file", csv.name, logger)
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    table, _fig, png = app.make_ratio_bar(app.require_df(), "산업군")
    assert png is not None
    png_res = _get_file(http, Path(png))
    assert png_res.status_code == 200
    result = app.do_report(logger)
    docx_res = _get_file(http, Path(result["path"]))
    assert docx_res.status_code == 200
    wc = app.STORE.get("wordcloud_path")
    if wc:
        assert _get_file(http, Path(wc)).status_code == 200


def test_gradio_cache_copy_of_export_allowed(http, tmp_path, logger):
    csv = write_voc_csv(tmp_path / "ok2.csv", dates=["2026-01-06", "2026-01-07"])
    app.do_upload(csv, "file", csv.name, logger)
    _table, _fig, png = app.make_ratio_bar(app.require_df(), "산업군")
    assert png is not None
    cache = Path(get_cache_folder())
    cache.mkdir(parents=True, exist_ok=True)
    cached = cache / "cached_export.png"
    shutil.copy(png, cached)
    assert _get_file(http, cached).status_code == 200
    csv_cache = cache / "not_an_export.csv"
    shutil.copy(csv, csv_cache)
    assert _get_file(http, csv_cache).status_code == 403


def test_extension_alone_is_not_enough(http, tmp_path):
    rogue = tmp_path / "random.png"
    rogue.write_bytes(b"\x89PNG\r\n\x1a\n" + b"not-registered")
    assert _get_file(http, rogue).status_code == 403
    rogue_docx = tmp_path / "random.docx"
    rogue_docx.write_bytes(b"PK\x03\x04not-registered")
    assert _get_file(http, rogue_docx).status_code == 403
