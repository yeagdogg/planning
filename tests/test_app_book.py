"""P4: the multi-line book — engine ties, the wtd() rule, cache discipline.

The frame and the weighted math are pure (pandas + the engine) and run on
both interpreters through a stub session; the aggregate tie is the pivot-
check pattern — recompute every combo INDEPENDENTLY through the engine and
land on the same EP-weighted number. The page test at the bottom drives
the real PlanSession + Book + Combo pages in the app venv, including the
click-through.
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


def _stub(scns):
    page = SimpleNamespace(config=scns[0], data_version=0, results={},
                           caches={}, book={s.lob: s for s in scns})
    return SimpleNamespace(page=page, bus=SimpleNamespace(rev=0),
                           ctx=SimpleNamespace(results_rev=0))


@pytest.fixture(scope="module")
def sess():
    from app import importers
    return _stub([importers.from_workbook(WB_PROP),
                  importers.from_workbook(WB_GL)])


def test_fleet_choices_prefers_the_untagged_file():
    """output/ holds Property AND Property_MODERN — the fleet loader must
    pick the plain fleet file, one entry per configured line."""
    from app import importers
    choices = importers.fleet_choices()
    assert len(choices) == len(importers.app_config().lobs)
    assert choices["Property"].endswith("Plan_LR_Workbook_2027_Property.xlsx")


def test_book_frame_ties_the_engine(sess):
    from src import engine

    from app import compute
    frame = compute.book_frame(sess)
    assert len(frame) == sum(len(s.combo_keys())
                             for s in sess.page.book.values())
    assert frame["error"].eq("").all()
    # spot rows, one per line, against a FRESH engine run
    for lob, scn in sess.page.book.items():
        key = scn.combo_keys()[0]
        row = frame[frame["lob"] == lob].iloc[0]
        assert f"{row['bu']}|{row['state']}" == key
        res = engine.run_bridge(scn.plan_year,
                                compute.combo_inputs(scn, scn.row(key)),
                                "monthly")
        assert row["lr_p"] == pytest.approx(res.cy_lr_p, abs=1e-12)
        assert row["lr_p1"] == pytest.approx(res.cy_lr_p1, abs=1e-12)
        assert row["a_rate"] == pytest.approx(res.a_rate_p, abs=1e-12)


def test_weighted_follows_the_workbook_rule(sess):
    """The pivot-check pattern: rebuild the book LR by hand — every combo
    through the engine independently, EP-weighted — and match
    book_weighted() to 1e-12. Then the per-line roll-up must equal
    filtering to that line."""
    from src import engine

    from app import compute
    frame = compute.book_frame(sess)
    w = compute.book_weighted(frame)

    num = den = 0.0
    for scn in sess.page.book.values():
        for key in scn.combo_keys():
            res = engine.run_bridge(scn.plan_year,
                                    compute.combo_inputs(scn, scn.row(key)),
                                    "monthly")
            ep = scn.ep(key)
            num += ep * res.cy_lr_p
            den += ep
    assert den > 0
    assert w["lr_p"] == pytest.approx(num / den, abs=1e-12)
    assert w["ep"] == pytest.approx(den)

    roll = compute.line_rollup(frame)
    for lob in sess.page.book:
        sub_w = compute.book_weighted(frame[frame["lob"] == lob])
        r = roll[roll["lob"] == lob].iloc[0]
        assert r["lr_p"] == pytest.approx(sub_w["lr_p"], abs=1e-15)
        assert int(r["combos"]) == sub_w["combos"]

    # no EP in view -> means are None, not a divide-by-zero
    z = frame.copy()
    z["ep"] = None
    assert compute.book_weighted(z)["lr_p"] is None


def test_only_the_active_line_recomputes(monkeypatch):
    """The cache contract the live book depends on: a bus bump (a grid
    edit) recomputes the ACTIVE line only; untouched lines serve cached."""
    from app import compute, importers
    sess = _stub([importers.from_workbook(WB_PROP),
                  importers.from_workbook(WB_GL)])
    compute.book_results(sess)                     # cold fill

    calls = []
    real = compute.engine.run_bridge

    def counting(plan_year, ci, *a, **k):
        calls.append(ci)
        return real(plan_year, ci, *a, **k)

    monkeypatch.setattr(compute.engine, "run_bridge", counting)
    compute.book_results(sess)                     # warm: nothing runs
    assert calls == []

    sess.bus.rev += 1                              # an edit on the active line
    compute.book_results(sess)
    assert len(calls) == len(sess.page.config.combo_keys())


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_book_page_filters_and_clicks_through():
    """The real thing: two lines loaded, the table shows them all, a line
    filter narrows it, and clicking a combo row activates that line,
    focuses that combo on the Combo page, and navigates."""
    from app import compute, importers
    from app.glue.session import PlanSession
    from app.pages import book as book_page
    from app.pages import combo as combo_page

    session = PlanSession()
    gone = []
    session.goto = gone.append
    combo = combo_page.build(session)          # built first: adopts focus
    page = book_page.build(session)

    gl = importers.from_workbook(WB_GL)
    prop = importers.from_workbook(WB_PROP)
    session.replace_config(gl)
    session.replace_config(prop)
    assert set(session.page.book) == {gl.lob, prop.lob}
    assert session.page.config is prop            # last loaded is active

    table = page["table"]
    frame = compute.book_frame(session)
    assert len(table.value) == len(frame)
    assert page["cards"].object and "plan LR" in page["cards"].object

    page["filters"]["line"].value = [gl.lob]
    assert set(table.value["lob"]) == {gl.lob}
    target = table.value.iloc[0]
    page["filters"]["line"].value = []
    assert len(table.value) == len(frame)

    # click-through on the first GL row (unfiltered frame starts with GL —
    # it was loaded first, and the book keeps load order)
    first = table.value.iloc[0]
    assert first["lob"] == gl.lob == target["lob"]
    page["open_row"](SimpleNamespace(row=0))
    assert session.page.config is session.page.book[gl.lob]
    assert combo["combo_sel"].value == f"{first['bu']}|{first['state']}"
    assert session.ctx.focus is None              # consumed, one-shot
    assert gone == ["combo"]

    # the on_show hooks re-render without error (the visibility-flip fix)
    assert callable(page["on_show"])
    page["on_show"]()
    assert len(page["table"].value) == len(frame)


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_click_through_before_the_combo_page_exists():
    """The LIVE bug this pins: pages build lazily, so a user's FIRST-ever
    click-through sets focus and bumps data_rev BEFORE the Combo page is
    built — the watcher alone misses it. Building the page must adopt the
    pending focus."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import book as book_page
    from app.pages import combo as combo_page

    session = PlanSession()
    gone = []
    session.goto = gone.append
    page = book_page.build(session)               # book only — no combo yet
    session.replace_config(importers.from_workbook(WB_PROP))

    target = page["table"].value.iloc[3]          # not the first combo
    page["open_row"](SimpleNamespace(row=3))
    assert gone == ["combo"]

    combo = combo_page.build(session)             # what _switch would do now
    assert combo["combo_sel"].value == f"{target['bu']}|{target['state']}"
    assert session.ctx.focus is None
