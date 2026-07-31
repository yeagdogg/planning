"""Harvester contracts (D66): what it reads, and what it refuses.

The harvester is the book's only source of truth, so its failure modes matter
more than its happy path — a silently under-reported harvest would roll up
into a plausible, wrong book.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import datetime as dt
from pathlib import Path

import openpyxl
import pytest

from src.build_workbook import Layout as L, load_config
from src.sheets_calc import BOOK_PUB, BOOK_SLOTS
from tools.harvest import (LAST_COL, REQUIRED, SCALARS, BookData, HarvestError,
                           harvest, read_lob)

CFG = load_config("config/config.yaml")


def _stub_workbook(path: Path, keys, *, calculated=True, version="9.9.9",
                   n_cols=LAST_COL):
    """A minimal stand-in for a generated workbook's published surface."""
    wb = openpyxl.Workbook()
    wb.remove(wb.active)
    rm = wb.create_sheet("Read Me")
    rm["B3"] = f"Plan year 2027  |  x  |  version {version}  |  built today"
    ws = wb.create_sheet("_calc")
    for i, key in enumerate(keys):
        r = L.CALC_RES_FIRST + i
        for c in range(1, n_cols + 1):
            ws.cell(row=r, column=c, value=(1.0 if calculated else None))
        ws.cell(row=r, column=1, value=key)
        ws.cell(row=r, column=SCALARS["state"], value=key.split("|")[1])
        ws.cell(row=r, column=SCALARS["bu"], value=key.split("|")[0])
    wb.save(path)
    return path


def test_reads_rows_and_provenance(tmp_path):
    p = _stub_workbook(tmp_path / "wb.xlsx", ["A|AZ", "A|CO", "B|AZ"])
    rows, src = read_lob(p, "Property")
    assert [r["key"] for r in rows] == ["A|AZ", "A|CO", "B|AZ"]
    assert all(r["lob"] == "Property" for r in rows)
    assert src.lob == "Property" and src.combos == 3 and src.version == "9.9.9"
    assert len(src.modified) == 16          # "YYYY-MM-DD HH:MM"
    # shapes: the monthly families and the rate-change slots
    assert len(rows[0]["delivered"]) == 12 and len(rows[0]["epw"]) == 12
    assert len(rows[0]["slots"]) == BOOK_SLOTS
    assert all(len(s) == 3 for s in rows[0]["slots"])


def test_skips_spare_rows(tmp_path):
    """Blank keys are the tbl_LR spares — skipped, not harvested as zeros."""
    p = tmp_path / "wb.xlsx"
    _stub_workbook(p, ["A|AZ"])
    wb = openpyxl.load_workbook(p)
    ws = wb["_calc"]
    for c in range(1, LAST_COL + 1):        # a fully blank spare row
        ws.cell(row=L.CALC_RES_FIRST + 1, column=c, value=None)
    wb.save(p)
    rows, src = read_lob(p, "Property")
    assert len(rows) == 1 and src.combos == 1


def test_refuses_uncalculated_workbook(tmp_path):
    p = _stub_workbook(tmp_path / "stale.xlsx", ["A|AZ"], calculated=False)
    with pytest.raises(HarvestError) as e:
        read_lob(p, "Property")
    assert "never" in str(e.value) and "A|AZ" in str(e.value)


def test_refuses_workbook_from_an_older_generator(tmp_path):
    """No BOOK_PUB columns at all — the schema probe must catch it."""
    p = _stub_workbook(tmp_path / "old.xlsx", ["A|AZ"],
                       n_cols=BOOK_PUB["ntaken"] - 1)
    with pytest.raises(HarvestError) as e:
        read_lob(p, "Property")
    assert "ntaken" in str(e.value)


def test_refuses_missing_file(tmp_path):
    with pytest.raises(HarvestError) as e:
        read_lob(tmp_path / "nope.xlsx", "Property")
    assert "not found" in str(e.value)


def test_refuses_duplicate_keys(tmp_path):
    p = _stub_workbook(tmp_path / "dupe.xlsx", ["A|AZ", "A|AZ"])
    with pytest.raises(HarvestError) as e:
        read_lob(p, "Property")
    assert "duplicate" in str(e.value) and "A|AZ" in str(e.value)


def test_refuses_empty_publication(tmp_path):
    p = _stub_workbook(tmp_path / "empty.xlsx", [])
    with pytest.raises(HarvestError):
        read_lob(p, "Property")


def test_book_rosters_preserve_first_appearance_order():
    book = BookData(plan_year=2027, rows=[
        {"lob": "P", "bu": "MM", "state": "AZ"},
        {"lob": "P", "bu": "ABD", "state": "CO"},
        {"lob": "GL", "bu": "MM", "state": "AZ"},
    ])
    assert book.states == ["AZ", "CO"]          # de-duplicated, order kept
    assert book.business_units == ["ABD", "MM"]  # sorted


def test_harvest_of_the_shipped_book():
    """End to end against the real output/ workbooks (skips if not built)."""
    try:
        book = harvest(CFG, now=dt.datetime(2027, 1, 2, 3, 4))
    except HarvestError as e:
        pytest.skip(f"LOB workbooks not harvestable: {e}")
    assert book.as_of == "2027-01-02 03:04"
    assert len(book.sources) == len(CFG.output_lobs)
    assert len(book.rows) == sum(s.combos for s in book.sources)
    assert set(book.business_units) == set(CFG.business_units)
    assert book.states == list(CFG.states)
    for r in book.rows:
        assert all(r[f] is not None for f in REQUIRED)
        assert r["key"] == f"{r['bu']}|{r['state']}"
