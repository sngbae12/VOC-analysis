# -*- coding: utf-8 -*-
from __future__ import annotations

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def test_stats_and_wordcloud_do_not_call_llm_even_with_key(tmp_path, logger, monkeypatch):
    csv = write_voc_csv(tmp_path / "ok.csv", dates=["2026-01-01", "2026-01-02"])
    app.do_upload(csv, "file", "ok.csv", logger)
    monkeypatch.setenv("OPENAI_API_KEY", "sk-local-test-must-not-be-called")
    app.set_runtime_api_key("sk-local-test-must-not-be-called")

    def explode(*_a, **_k):
        raise AssertionError("LLM/API must not be created for local analysis")

    monkeypatch.setattr(app, "create_llm", explode)
    monkeypatch.setattr(app, "run_issue_analysis", explode)
    monkeypatch.setattr(app, "run_crew", explode)
    fig, table = app.do_stats("산업군", logger)
    assert table is not None and not table.empty
    path, = app.do_wordcloud(logger)
    assert PathLike(path)


def PathLike(path: str) -> bool:
    from pathlib import Path

    return Path(path).exists()


def test_python_range_rejects_3_14():
    class Info:
        major = 3
        minor = 14
        micro = 0

    ok, msg = app.python_support_message(Info())
    assert ok is False
    assert "3.14" in msg
    assert "3.11" in msg


def test_python_range_accepts_3_11():
    class Info:
        major = 3
        minor = 11
        micro = 9

    ok, msg = app.python_support_message(Info())
    assert ok is True
    assert msg == ""
