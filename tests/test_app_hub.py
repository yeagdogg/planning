"""P5: the hub — master paste, scenario files with a changelog, export.

The property under test everywhere is CLOSURE: master rows == per-line
rows (one vocabulary), a saved book loads back identical, and an exported
workbook re-imports as the scenario that built it (workbook → app →
workbook → app is one loop). UI-free layers run on both interpreters;
the page flows at the bottom need the app venv.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"

MASTER_RL = (
    "Line\tBU\tState\tEffective\tFiled %\tStatus\tConsidered"
    "\tAchievement\tComment\n"
    "Property\tABD\tAZ\t3/31/2026\t10.3%\ttaken\tY\t\t\n"
    "General Liability\tABD\tAZ\t5/1/2027\t8.0%\tplanned\t\t75%\tspring\n"
    "property\tABD\tCA\t6/1/2027\t4.0%\tplanned\t\t\t\n"
    "\t\t\t\t\t\t\t\t")                       # spare row: skipped


class _Sess:
    """The attribute surface masters/compute need, on both interpreters."""

    def __init__(self, book=None):
        self.page = SimpleNamespace(config=None, book=dict(book or {}),
                                    data_version=0, results={}, caches={})
        self.bus = SimpleNamespace(rev=0)
        self.ctx = SimpleNamespace(results_rev=0)
        self.activated = []

    def activate(self, lob):
        self.activated.append(lob)
        self.page.config = self.page.book.get(lob)
        return True


# ---------------------------------------------------------------- masters

def test_split_master_by_line_case_insensitive_and_spares_drop():
    from app.masters import split_master
    notes: list = []
    by = split_master("Rate Log", MASTER_RL,
                      ["Property", "General Liability"], notes=notes)
    assert notes == ["header row skipped"]
    assert set(by) == {"Property", "General Liability"}
    assert len(by["Property"]) == 2               # 'property' folded in
    assert by["Property"][0]["filed"] == pytest.approx(0.103)
    assert "lob" not in by["Property"][0]         # per-line vocabulary
    # ...and the rows equal what the per-line paste block produces
    from app.grids import RL_SCHEMA
    from app.paste import apply_block
    solo = apply_block(RL_SCHEMA, "ABD\tAZ\t3/31/2026\t10.3%\ttaken\tY\t\t")
    assert by["Property"][0] == solo[0]


def test_split_master_is_loud_about_bad_lines():
    from app.masters import split_master
    with pytest.raises(ValueError, match="unknown line 'Umbrela'"):
        split_master("Rate Log",
                     "Umbrela\tABD\tAZ\t3/31/2026\t10%\ttaken",
                     ["Property", "Umbrella"])
    with pytest.raises(ValueError, match="Line column is blank"):
        split_master("Rate Log", "\tABD\tAZ\t3/31/2026\t10%\ttaken",
                     ["Property"])


def test_apply_master_creates_lines_with_config_defaults():
    from app import importers
    from app.masters import apply_master
    sess = _Sess()
    by, created = apply_master(sess, "Rate Log", MASTER_RL)
    assert set(created) == {"Property", "General Liability"}
    cfg = importers.app_config()
    prop = sess.page.book["Property"]
    assert prop.term_months == cfg.lob("Property").term_months
    assert prop.plan_year == cfg.plan_year
    assert prop.rate_rows == by["Property"]
    assert prop.lr_rows == []                     # only the pasted table
    assert sess.activated                          # one wake-up swap

    # partial re-paste: only the pasted line's table moves
    apply_master(sess, "Rate Log",
                 "Property\tABD\tAZ\t7/1/2027\t2.0%\ttaken\t\t\t")
    assert len(sess.page.book["Property"].rate_rows) == 1
    assert len(sess.page.book["General Liability"].rate_rows) == 1


# ----------------------------------------------------------- scenario files

@pytest.fixture(scope="module")
def two_line_book():
    from app import importers
    prop = importers.from_workbook(WB_PROP)
    gl = importers.from_workbook(WB_GL)
    return {prop.lob: prop, gl.lob: gl}


def _tables_equal(a, b):
    return (a.lr_rows == b.lr_rows and a.rate_rows == b.rate_rows
            and a.mod_rows == b.mod_rows and a.season_rows == b.season_rows
            and a.term_months == b.term_months
            and a.season_on == b.season_on and a.mod_master == b.mod_master
            and a.trend_default == b.trend_default)


def test_scenario_file_round_trips_and_logs_changes(two_line_book, tmp_path):
    from app import scenarios_io as sio
    p = tmp_path / "trip.yaml"
    changes = sio.save_book(p, two_line_book, "trip")
    assert changes and "initial save" in changes[0]

    loaded, doc = sio.load_book(p)
    assert set(loaded) == set(two_line_book)
    for lob in loaded:
        assert _tables_equal(loaded[lob], two_line_book[lob]), lob

    # a field edit -> the changelog names it, formatted per its kind
    import copy
    book2 = copy.deepcopy(two_line_book)
    prop = book2["Property"]
    key = prop.combo_keys()[0]
    prop.row(key)["netp"] = 0.025
    changes = sio.save_book(p, book2, "trip", note="net bump")
    assert any("Net rate P" in c and "2.5%" in c and key in c
               for c in changes)
    # a no-op save says so instead of logging nothing
    assert sio.save_book(p, book2, "trip") == ["no input changes"]
    log = sio.read_doc(p).get("changelog")
    assert len(log) == 3 and log[1]["note"] == "net bump"


def test_diff_books_sees_lines_and_log_rows(two_line_book):
    from app import scenarios_io as sio
    full = sio.book_to_doc(two_line_book, "x")["lines"]
    one = {"Property": full["Property"]}
    d = sio.diff_books(one, full)
    assert any(c.startswith("+ General Liability: line added") for c in d)
    import copy
    trimmed = copy.deepcopy(full)
    dropped = trimmed["Property"]["rate_rows"].pop(0)
    d2 = sio.diff_books(full, trimmed)
    assert any("Rate Log: −" in c for c in d2)
    assert any(str(dropped["eff"]) in c for c in d2 if "Rate Log" in c)


# ----------------------------------------------------------------- export

def test_export_reimports_as_the_same_scenario(tmp_path):
    """The loop: workbook → app → workbook → app. The exported file is a
    real generator workbook carrying the scenario's inputs; re-importing
    it must reproduce every table and toggle."""
    from app import exporters, importers
    scn = importers.from_workbook(WB_PROP)
    out = exporters.export_line(scn, out_dir=tmp_path)
    assert out.name.endswith("_APP.xlsx")
    back = importers.from_workbook(out)
    assert back.lob == scn.lob
    assert _tables_equal(back, scn)


# ------------------------------------------------------- page flows (venv)

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_book_page_master_paste_flow():
    from app.glue.session import PlanSession
    from app.pages import book as book_page

    session = PlanSession()
    page = book_page.build(session)
    ta, btn, flash = page["masters"]["Rate Log"]

    btn.clicks += 1                                # empty box -> loud
    assert "Nothing to paste" in flash.object

    ta.value_input = MASTER_RL
    btn.clicks += 1
    assert "Applied" in flash.object and "2 line(s)" in flash.object
    assert set(session.page.book) == {"Property", "General Liability"}
    assert ta.value_input == "" and ta.value == ""
    assert len(page["table"].value) == 0           # no combos yet (no tbl_LR)

    ta.value_input = "Umbrela\tABD\tAZ\t3/31/2026\t10%\ttaken"
    btn.clicks += 1
    assert "unknown line" in flash.object and "nothing was applied" in flash.object


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_scenarios_page_save_load_diff_generate(tmp_path, monkeypatch):
    from app import importers, scenarios_io
    from app.glue.session import PlanSession
    from app.pages import scenarios as scen_page

    monkeypatch.setattr(scenarios_io, "SCEN_DIR", tmp_path)
    session = PlanSession()
    session.replace_config(importers.from_workbook(WB_PROP))
    page = scen_page.build(session)
    w, f = page["widgets"], page["flashes"]

    # save
    w["name"].value = "pytest-book"
    w["save"].clicks += 1
    assert "Saved" in f["save"].object
    saved = tmp_path / "pytest-book.yaml"
    assert saved.exists()
    assert w["file"].value == str(saved)

    # edit a row, diff file -> current names it
    scn = session.page.config
    key = scn.combo_keys()[0]
    scn.row(key)["netp1"] = 0.031
    session.bus.bump()
    w["diff"].clicks += 1
    assert "Net rate P+1" in f["log"].object and "3.1%" in f["log"].object

    # two-step load restores the file's value
    w["load"].clicks += 1
    assert "Confirm" in w["load"].label            # armed, nothing loaded yet
    assert scn.row(key)["netp1"] == 0.031
    w["load"].clicks += 1
    assert "Loaded" in f["load"].object
    fresh = session.page.config
    assert fresh.row(key)["netp1"] != 0.031        # the edit was replaced

    # generate one line, no recalc (headless run_async is synchronous)
    w["lines"].value = ["Property"]
    w["generate"].clicks += 1
    assert "Built 1 workbook" in f["gen"].object
    assert "first open in Excel" in f["gen"].object
