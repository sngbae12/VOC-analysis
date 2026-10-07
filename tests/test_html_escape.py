# -*- coding: utf-8 -*-
from __future__ import annotations

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def test_filename_and_csv_values_are_escaped(tmp_path, logger):
    nasty = write_voc_csv(
        tmp_path / "ok.csv",
        dates=["2026-01-01"],
        extra={1: {"산업군": "<img src=x onerror=alert(1)>", "고객명": "<b>x</b>"}},
    )
    df = app.load_voc_csv(nasty)
    state = app.apply_loaded_df(df, "file", "<script>alert(1)</script>.csv", logger)
    header = app.render_header(state)
    cards = app.render_upload_cards(state)
    notice = app.render_notice("fail", "<script>", "<img src=x>")
    table = app.ratio_table(df, "산업군")
    stats = app.render_stats_cards(table, "산업군", state)
    assert "<script>" not in header
    assert "<script>" not in cards
    assert "<img src=x onerror=alert(1)>" not in stats
    assert "&lt;script&gt;" in header
    assert "&lt;img" in stats
    assert "<script>" not in notice
    assert "&lt;script&gt;" in notice


def test_stats_cards_escape_category(tmp_path, logger):
    csv = write_voc_csv(
        tmp_path / "ok.csv",
        dates=["2026-01-01", "2026-01-02"],
        extra={1: {"산업군": "<svg/onload=alert(1)>"}},
    )
    df = app.load_voc_csv(csv)
    state = app.apply_loaded_df(df, "file", "ok.csv", logger)
    table = app.ratio_table(df, "산업군")
    html = app.render_stats_cards(table, "산업군", state)
    assert "<svg" not in html
    assert "&lt;svg" in html
