# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import pandas as pd

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def _period_from_docx(path: Path) -> str:
    from docx import Document

    doc = Document(str(path))
    texts = [p.text for p in doc.paragraphs if p.text.strip()]
    overview = next((t for t in texts if t.startswith("총 VOC")), "")
    return overview


def test_yyyymmdd_int_and_string_not_epoch():
    series = pd.Series([20260101, "20260102", 20260103])
    parsed, invalid, empty = app.parse_voc_dates(series)
    assert invalid == 0
    assert empty == 0
    assert list(parsed.dt.strftime("%Y-%m-%d")) == ["2026-01-01", "2026-01-02", "2026-01-03"]
    assert not any(ts.year == 1970 for ts in parsed)


def test_iso_dates_still_work():
    series = pd.Series(["2026-01-06", "2026-03-14"])
    parsed, invalid, empty = app.parse_voc_dates(series)
    assert invalid == 0 and empty == 0
    assert parsed.min().strftime("%Y-%m-%d") == "2026-01-06"
    assert parsed.max().strftime("%Y-%m-%d") == "2026-03-14"


def test_mixed_invalid_empty_do_not_crash(tmp_path):
    csv = write_voc_csv(
        tmp_path / "mixed.csv",
        dates=[20260101, "2026-01-02", "2026/01/03", "", "not-a-date", "20261399", "2026.01.04"],
        extra={5: {"불만": "잘못된 날짜 행도 불만은 있습니다"}},
    )
    df = app.load_voc_csv(csv)
    assert len(df) == 7
    assert int(df.attrs["date_invalid"]) == 2
    assert int(df.attrs["date_empty"]) == 1
    period = app.voc_period_text(df)
    assert period == "2026-01-01 ~ 2026-01-04"
    overview = app.voc_overview(df)
    assert "2026-01-01 ~ 2026-01-04" in overview
    assert "해석 실패 2건" in overview


def test_period_matches_digest_and_docx(tmp_path, logger, monkeypatch):
    csv = write_voc_csv(tmp_path / "ymd.csv", dates=[20260101, 20260102, "2026-01-03"])
    df = app.load_voc_csv(csv)
    app.apply_loaded_df(df, "file", "ymd.csv", logger)
    period = app.voc_period_text(df)
    assert period == "2026-01-01 ~ 2026-01-03"
    assert period in app.voc_overview(df)
    assert period in app.build_voc_digest(df)
    monkeypatch.setattr(app, "run_issue_analysis", lambda *a, **k: (_ for _ in ()).throw(RuntimeError("skip")))
    monkeypatch.setattr(app, "resolve_api_key", lambda: "")
    result = app.do_report(logger)
    overview = _period_from_docx(Path(result["path"]))
    assert "2026-01-01 ~ 2026-01-03" in overview
    assert app.snapshot_from_df(df, "file", "ymd.csv")["period"] == period


def test_all_invalid_dates_keep_rows(tmp_path):
    csv = write_voc_csv(tmp_path / "bad.csv", dates=["어제", "내일"])
    df = app.load_voc_csv(csv)
    assert len(df) == 2
    assert app.voc_period_text(df) == "-"
    assert "해석 실패" in app.voc_overview(df)
