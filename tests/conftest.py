# -*- coding: utf-8 -*-
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
if str(ROOT) not in sys.path:
    sys.path.insert(0, str(ROOT))

import voc_analysis_app as app  # noqa: E402


REQUIRED = ["순번", "일자", "고객명", "산업군", "지역", "제품명", "분야", "불만"]


def fake_rows(dates, extra=None):
    rows = []
    for i, day in enumerate(dates, start=1):
        row = {
            "순번": i,
            "일자": day,
            "고객명": f"가상고객{i}",
            "산업군": "제조" if i % 2 else "유통",
            "지역": "경기",
            "제품명": "MES",
            "분야": "품질",
            "불만": f"가상 불만 문구 {i}번입니다",
        }
        if extra and i in extra:
            row.update(extra[i])
        rows.append(row)
    return rows


def write_voc_csv(path: Path, dates, extra=None, encoding="utf-8-sig") -> Path:
    df = pd.DataFrame(fake_rows(dates, extra))
    path.write_text(df.to_csv(index=False), encoding=encoding)
    return path


def reset_store() -> None:
    with app.STORE_LOCK:
        app.STORE.update(
            {
                "df": None,
                "csv_path": None,
                "source_kind": None,
                "source_label": None,
                "charts": {},
                "stats": {},
                "wordcloud_path": None,
                "keywords": {},
                "issue_analysis": "",
                "report_path": None,
                "chart_errors": {},
                "report_missing_images": [],
            }
        )
    with app._EXPORT_LOCK:
        app._EXPORT_DIGESTS.clear()


@pytest.fixture(autouse=True)
def _clean_store():
    reset_store()
    yield
    reset_store()


@pytest.fixture
def logger():
    return lambda _msg: None
