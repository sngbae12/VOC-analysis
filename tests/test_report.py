# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import pytest
from docx import Document

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def _load_sample(logger):
    df = app.load_voc_csv(app.SAMPLE_CSV)
    app.apply_loaded_df(df, "sample", app.SAMPLE_CSV.name, logger)
    return df


def _docx_text(path: Path) -> str:
    doc = Document(str(path))
    return "\n".join(p.text for p in doc.paragraphs)


def test_wordcloud_failure_still_saves_docx(logger, monkeypatch):
    _load_sample(logger)
    calls = {"n": 0}

    def boom(*_a, **_k):
        calls["n"] += 1
        raise RuntimeError("persistent-wordcloud-error")

    monkeypatch.setattr(app, "make_wordcloud", boom)
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    result = app.do_report(logger)
    assert calls["n"] == 1
    path = Path(result["path"])
    assert path.exists()
    assert path.suffix.lower() == ".docx"
    assert "워드클라우드" in result["missing_images"]
    text = _docx_text(path)
    assert "완전한 보고서가 아닙니다" in text
    assert "워드클라우드" in text
    assert "그림이 빠져" in text
    assert result["preview"] is not None
    assert result["path"] == str(app.STORE["report_path"])


def test_stats_png_failure_still_saves_docx(logger, monkeypatch):
    _load_sample(logger)

    def tables_only(df, column):
        table = app.ratio_table(df, column)
        fig = app.go.Figure()
        with app.STORE_LOCK:
            app.STORE["stats"][column] = table
            app.STORE["charts"][column] = None
        return table, fig, None

    monkeypatch.setattr(app, "make_ratio_bar", tables_only)
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    result = app.do_report(logger)
    assert Path(result["path"]).exists()
    assert "산업군 비율 그림" in result["missing_images"]
    assert "분야 비율 그림" in result["missing_images"]
    text = _docx_text(Path(result["path"]))
    assert "완전한 보고서가 아닙니다" in text


def test_all_images_fail_still_saves_incomplete_docx(logger, monkeypatch):
    _load_sample(logger)

    def tables_only(df, column):
        table = app.ratio_table(df, column)
        with app.STORE_LOCK:
            app.STORE["stats"][column] = table
            app.STORE["charts"][column] = None
        return table, app.go.Figure(), None

    monkeypatch.setattr(app, "make_ratio_bar", tables_only)
    monkeypatch.setattr(app, "make_wordcloud", lambda *_a, **_k: (_ for _ in ()).throw(RuntimeError("wc")))
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    result = app.do_report(logger)
    assert Path(result["path"]).exists()
    missing = result["missing_images"]
    assert "워드클라우드" in missing
    assert "산업군 비율 그림" in missing
    assert "분야 비율 그림" in missing
    text = _docx_text(Path(result["path"]))
    assert "불완전" in text or "완전한 보고서가 아닙니다" in text


def test_issue_timeout_does_not_reuse_previous(logger, monkeypatch, tmp_path):
    _load_sample(logger)
    old_path = tmp_path / "old.docx"
    old_path.write_bytes(b"old")
    with app.STORE_LOCK:
        app.STORE["issue_analysis"] = "이전 이슈 본문"
        app.STORE["report_path"] = old_path

    def timeout(*_a, **_k):
        raise TimeoutError("timed out")

    monkeypatch.setattr(app, "resolve_api_key", lambda: "sk-test-not-used")
    monkeypatch.setattr(app, "run_issue_analysis", timeout)
    result = app.do_report(logger)
    assert result["path"] != str(old_path)
    assert "이전 이슈 본문" not in result["preview"]
    assert "이슈 분석을 만들지 못했습니다" in result["preview"]
    text = _docx_text(Path(result["path"]))
    assert "이전 이슈 본문" not in text
    assert result["issue_ok"] is False


def test_docx_save_failure_does_not_return_old_report(logger, monkeypatch, tmp_path):
    _load_sample(logger)
    old = tmp_path / "prev.docx"
    old.write_bytes(b"prev")
    with app.STORE_LOCK:
        app.STORE["report_path"] = old
        app.STORE["issue_analysis"] = "이전 성공 이슈"

    def boom_save(self, *_a, **_k):
        raise OSError("cannot save")

    from docx.document import Document as DocxDocument

    monkeypatch.setattr(DocxDocument, "save", boom_save)
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    with pytest.raises(OSError):
        app.do_report(logger)
    assert app.STORE.get("report_path") in (None, old)
    assert app.STORE.get("issue_analysis") in ("", "이전 성공 이슈")
