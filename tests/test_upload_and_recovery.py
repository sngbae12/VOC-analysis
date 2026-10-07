# -*- coding: utf-8 -*-
from __future__ import annotations

from pathlib import Path

import voc_analysis_app as app
from tests.conftest import write_voc_csv


def test_sample_fifty_rows(logger):
    out = app.do_upload(app.SAMPLE_CSV, "sample", app.SAMPLE_CSV.name, logger)
    assert out["state"]["row_count"] == 50
    assert out["state"]["source_kind"] == "sample"
    assert "sample_voc.csv" in out["state"]["source_label"]
    chip = app.sample_button_update(out["state"])
    assert "선택됨" in str(chip.get("value", chip))


def test_new_csv_replaces_and_clears_previous(tmp_path, logger):
    first = write_voc_csv(tmp_path / "a.csv", dates=["2026-01-01", "2026-01-02"])
    second = write_voc_csv(tmp_path / "b.csv", dates=["2026-02-01", "2026-02-02", "2026-02-03"])
    app.do_upload(first, "file", "a.csv", logger)
    app.do_stats("산업군", logger)
    app.do_wordcloud(logger)
    assert app.STORE["stats"]
    assert app.STORE["wordcloud_path"]
    out = app.do_upload(second, "file", "b.csv", logger)
    assert out["state"]["row_count"] == 3
    assert out["state"]["source_label"] == "b.csv"
    assert app.STORE["stats"] == {}
    assert app.STORE["wordcloud_path"] is None
    assert app.STORE["report_path"] is None


def test_repeat_same_file_reapply(tmp_path, logger):
    csv = write_voc_csv(tmp_path / "same.csv", dates=["2026-01-10", "2026-01-11"])
    first = app.do_upload(csv, "file", "same.csv", logger)
    second = app.do_upload(csv, "file", "same.csv", logger)
    assert first["state"]["row_count"] == second["state"]["row_count"] == 2


def test_bad_file_keeps_previous_data(tmp_path, logger):
    good = write_voc_csv(tmp_path / "good.csv", dates=["2026-01-01"])
    app.do_upload(good, "file", "good.csv", logger)
    previous = app.STORE["df"].copy()
    bad = tmp_path / "bad.csv"
    bad.write_text("not,a,voc\n1,2,3\n", encoding="utf-8")
    try:
        app.do_upload(bad, "file", "bad.csv", logger)
        raise AssertionError("bad csv should fail")
    except ValueError:
        pass
    assert len(app.STORE["df"]) == len(previous)
    assert list(app.STORE["df"]["일자"]) == list(previous["일자"])


def test_control_updates_restore_buttons():
    state = app.empty_state()
    busy = app.control_updates(True, state, False)
    idle = app.control_updates(False, state, False)
    assert busy[1]["interactive"] is False
    assert idle[1]["interactive"] is True
    assert idle[3]["interactive"] is True
    assert idle[4]["interactive"] is True
    assert idle[5]["interactive"] is True
