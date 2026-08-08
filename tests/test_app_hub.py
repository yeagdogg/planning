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


def test_split_master_fills_a_blank_line_down():
    """W3h: a real Excel master names the line once per block (merged or
    typed once) — continuation rows paste with a BLANK Line and must
    carry the line above down, exactly like Excel reads them. A blank
    with nothing above still raises (the test above)."""
    from app.masters import split_master
    text = ("Property\tABD\tAZ\t3/31/2026\t10%\ttaken\t\t\t\n"
            "\tABD\tCO\t6/30/2026\t5%\ttaken\t\t\t\n"
            "General Liability\tABD\tAZ\t9/30/2026\t4%\tplanned\t\t\t\n"
            "\tSBA\tKY\t12/31/2026\t3%\tplanned\t\t\t")
    by = split_master("Rate Log", text, ["Property", "General Liability"])
    assert [r["state"] for r in by["Property"]] == ["AZ", "CO"]
    assert [r["bu"] for r in by["General Liability"]] == ["ABD", "SBA"]
    # a fully blank spare row between blocks neither breaks nor carries
    text2 = ("Property\tABD\tAZ\t3/31/2026\t10%\ttaken\t\t\t\n"
             "\t\t\t\t\t\t\t\t\n"
             "\tABD\tCO\t6/30/2026\t5%\ttaken\t\t\t")
    by2 = split_master("Rate Log", text2, ["Property"])
    assert [r["state"] for r in by2["Property"]] == ["AZ", "CO"]


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


# ------------------------------------------------------ masters copy-out

def test_master_text_round_trips(two_line_book):
    """The law: split_master(master_text(book)) == the book's own rows,
    for every master table — a generated master IS the book."""
    from app.masters import MASTERS, master_text, split_master
    known = list(two_line_book)
    for table, (_schema, attr) in MASTERS.items():
        text = master_text(two_line_book, table)
        assert text.splitlines()[0].startswith("Line\tBU\tState")
        notes: list = []
        by = split_master(table, text, known, notes=notes)
        assert notes == ["header row skipped"]
        for lob, scn in two_line_book.items():
            rows = getattr(scn, attr)
            if rows:
                assert by[lob] == rows, f"{table} / {lob}"
            else:
                assert lob not in by


def test_make_masters_writes_paste_ready_files(tmp_path):
    from app import importers
    from app.make_masters import make_masters
    from app.masters import split_master
    written = make_masters(tmp_path)
    assert set(written) == {"tbl_LR", "Rate Log", "Mod Log"}
    fleet = importers.fleet_choices()
    text = written["tbl_LR"].read_text(encoding="utf-8")
    by = split_master("tbl_LR", text, list(fleet))
    assert set(by) == set(fleet)                  # every line present
    prop = importers.from_workbook(fleet["Property"])
    assert by["Property"] == prop.lr_rows


def test_master_files_rebuild_the_fleet_book(tmp_path):
    """The closing closure: paste the GENERATED masters into an empty
    session, carry the toggles masters deliberately do not cover, and the
    engine's book is the fleet's book — every combo's plan LR equal."""
    import dataclasses

    from app import compute, importers
    from app.make_masters import make_masters
    from app.masters import MASTERS, apply_master

    fleet = {lob: importers.from_workbook(p)
             for lob, p in importers.fleet_choices().items()}
    written = make_masters(tmp_path)

    sess = _Sess()
    for table in MASTERS:
        apply_master(sess, table,
                     written[table].read_text(encoding="utf-8"))
    assert set(sess.page.book) == set(fleet)
    # masters carry the three tables; toggles + seasonality ride scenario
    # files / Inputs — copy them over so the comparison is like-for-like
    for lob, src in fleet.items():
        scn = sess.page.book[lob]
        sess.page.book[lob] = dataclasses.replace(
            scn, term_months=src.term_months, season_on=src.season_on,
            mod_master=src.mod_master, trend_default=src.trend_default,
            season_rows=[dict(r) for r in src.season_rows])
    sess.page.config = sess.page.book[next(iter(fleet))]

    mine = compute.book_frame(sess).set_index("key")
    theirs = compute.book_frame(_Sess(fleet)).set_index("key")
    assert set(mine.index) == set(theirs.index)
    assert len(mine) == len(theirs)
    for key in theirs.index:
        assert mine.loc[key, "lr_p"] == pytest.approx(
            theirs.loc[key, "lr_p"], abs=1e-12), key


# --------------------------------------------------- fleet + book (mocked)

def test_export_fleet_and_book_orchestration(two_line_book, tmp_path,
                                             monkeypatch):
    """Sequence + guards, with Excel and the subprocess mocked out (the
    real COM chain is smoked outside pytest, release-harness style)."""
    from app import exporters, importers
    cfg = importers.app_config()

    with pytest.raises(ValueError, match="missing"):
        exporters.export_fleet_and_book(dict(two_line_book))

    fleet = {lob: importers.from_workbook(p)
             for lob, p in importers.fleet_choices().items()}
    calls = {"recalc": [], "run": []}
    monkeypatch.setattr(exporters, "recalc_files",
                        lambda paths, **k: calls["recalc"].append(
                            [Path(p).name for p in paths]))

    def _fake_run(cmd, **kw):
        calls["run"].append(cmd)
        return SimpleNamespace(returncode=0, stdout="ok", stderr="")

    monkeypatch.setattr(exporters.subprocess, "run", _fake_run)
    paths, book_file = exporters.export_fleet_and_book(fleet,
                                                       out_dir=tmp_path)
    assert len(paths) == len(cfg.output_lobs)
    assert all(p.exists() and not p.stem.endswith("_APP") for p in paths)
    assert calls["recalc"][0] == [p.name for p in paths]   # lines first
    assert calls["recalc"][1] == [book_file.name]          # then the book
    assert any("tools/build_book.py" in str(a) for a in calls["run"][0])
    assert str(tmp_path) in calls["run"][0]

    bad = dict(fleet)
    bad["Property"] = dataclasses_replace_year(bad["Property"], 1999)
    with pytest.raises(ValueError, match="plan year"):
        exporters.export_fleet_and_book(bad, out_dir=tmp_path)


def dataclasses_replace_year(scn, year):
    import dataclasses
    return dataclasses.replace(scn, plan_year=year)


# ------------------------------------------------------- page flows (venv)

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_every_paste_box_takes_a_fleet_sized_block():
    """W2a: Panel's default max_length (5000) silently truncated pasted
    masters in the browser — every paste box must be uncapped far beyond
    any real block (tbl_LR master ≈ 60KB). W3g moved the masters to the
    Inputs page, so every box now lives there."""
    from app.glue.session import PlanSession
    from app.pages import inputs as inputs_page

    session = PlanSession()
    inputs = inputs_page.build(session)
    for label, (ta, _btn, _fl) in inputs["paste"].items():
        assert ta.max_length == 2_000_000, label
    for label, (ta, _btn, _fl) in inputs["masters"].items():
        assert ta.max_length == 2_000_000, label


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_guide_page_renders_from_the_schemas():
    from app.glue.session import PlanSession
    from app.pages import guide
    page = guide.build(PlanSession())
    body = page["main"][0].object
    assert "Plan EP (000s)" in body          # column lists come FROM schemas
    assert "`Line`" in body and "master" in body.lower()
    assert "mix" in body                      # the honesty rules are stated


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_inputs_page_master_paste_flow():
    """W3g: the masters live on Inputs and work on an EMPTY book — the
    true cold start. The Book page no longer carries them."""
    from app.glue.session import PlanSession
    from app.pages import book as book_page
    from app.pages import inputs as inputs_page

    session = PlanSession()
    page = inputs_page.build(session)
    assert page["master_card"].collapsed is False  # cold start: open
    ta, btn, flash = page["masters"]["Rate Log"]

    btn.clicks += 1                                # empty box -> loud
    assert "Nothing to paste" in flash.object

    ta.value_input = MASTER_RL
    btn.clicks += 1
    assert "Applied" in flash.object and "2 line(s)" in flash.object
    assert set(session.page.book) == {"Property", "General Liability"}
    assert ta.value_input == "" and ta.value == ""

    ta.value_input = "Umbrela\tABD\tAZ\t3/31/2026\t10%\ttaken"
    btn.clicks += 1
    assert "unknown line" in flash.object and "nothing was applied" in flash.object

    book = book_page.build(PlanSession())
    assert "masters" not in book                   # the Book slimmed (W3g)


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
