"""P3: the deep-dive visuals' NUMBERS, tied to the engine.

There is no pixel testing — the tie is the point. Every visual is an
ARRANGEMENT of engine output (the app has no second implementation), so
these tests pin the arrangement to the engine the same way the workbook
harness pins cells to it: rebuild each frame's aggregate the engine's way
and land on the engine's own number. Frames are pure pandas and run on
both interpreters; the final page-render test needs the app venv.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"


@pytest.fixture(scope="module")
def scn():
    from app import importers
    return importers.from_workbook(WB)


@pytest.fixture(scope="module")
def combo(scn):
    from app import compute
    key = scn.combo_keys()[0]
    row = scn.row(key)
    return key, row, compute.combo_inputs(scn, row)


@pytest.fixture(scope="module")
def res(scn, combo):
    from src import engine
    return engine.run_bridge(scn.plan_year, combo[2], "monthly")


@pytest.fixture(scope="module")
def eng(scn, combo):
    from src import engine
    return engine.MonthlyEngine(scn.plan_year, combo[2])


# ----------------------------------------------------------------- bridge

def test_chain_multiplies_to_the_headlines(scn, combo, res):
    """The card/waterfall chain IS the bridge: start × factors lands on the
    engine's own CY numbers for both years (the mapping tie)."""
    from app.views.bridge import chain
    start, steps_p, steps_p1, _net = chain(scn, combo[0], res)
    prod = start
    for _lbl, f in steps_p:
        prod *= f
    assert prod == pytest.approx(res.cy_lr_p, abs=1e-12)
    prod = start
    for _lbl, f in steps_p1:
        prod *= f
    assert prod == pytest.approx(res.cy_lr_p1, abs=1e-12)


def test_waterfall_bars_are_contiguous_and_land_on_the_total(scn, combo, res):
    from app.views.bridge import chain, waterfall_bars
    start, steps_p, _s1, _net = chain(scn, combo[0], res)
    bars = waterfall_bars("Projected LR", start, steps_p, "CY plan LR")
    assert bars[0]["lo"] == 0.0 and bars[0]["hi"] == pytest.approx(start)
    run = start
    for b, (_lbl, f) in zip(bars[1:-1], steps_p):
        nxt = run * f
        assert b["lo"] == pytest.approx(min(run, nxt), abs=1e-15)
        assert b["hi"] == pytest.approx(max(run, nxt), abs=1e-15)
        assert b["role"] == ("up" if nxt >= run else "down")
        run = nxt
    total = bars[-1]
    assert total["lo"] == 0.0
    assert total["hi"] == pytest.approx(res.cy_lr_p, abs=1e-12)


# ---------------------------------------------------------------- earning

def test_parallelogram_cells_rebuild_the_earned_series(scn, eng, res):
    """Aggregate the heatmap's cells the engine's way for one month and
    land on the engine's own earned index for that month — the frame holds
    every cohort the month earns from, at the right shares."""
    from src.engine import month_index

    from app.views.earning import parallelogram_frame
    df = parallelogram_frame(eng, scn.plan_year)
    lbl = f"{scn.plan_year}-07"
    sub = df[df["month"] == lbl]
    num = (sub["share"] * sub["weight"] * sub["index"]).sum()
    den = (sub["share"] * sub["weight"]).sum()
    mi = month_index(scn.plan_year, 7)
    got = res.earned_index_m[res.earned_month_mi.index(mi)]
    assert num / den == pytest.approx(got, abs=1e-12)

    # a cohort fully inside the window earns to exactly 1.0 across it
    jan = df[df["cohort"] == f"{scn.plan_year}-01"]
    assert jan["share"].sum() == pytest.approx(1.0, abs=1e-12)


def test_earn_lag_frame_reads_the_result_series(res):
    from app.views.earning import earn_lag_frame
    df = earn_lag_frame(res)
    assert len(df) == 24
    jan_p = res.earned_month_mi[12]              # series start Jan P-1
    assert df.loc[0, "written"] == pytest.approx(
        res.inforce_index[res.cohort_mi.index(jan_p)])
    assert df.loc[0, "earned"] == pytest.approx(res.earned_index_m[12])


# --------------------------------------------------------------- delivery

def test_band_shares_stack_to_one_and_fall_through_the_year(scn, eng):
    """Every plan-year cohort's premium lands in P or P+1 (term ≤ 12), and
    the share the plan year keeps falls monotonically through the year —
    January's cohort is nearly all P's, December's nearly all P+1's."""
    from app.views.delivery import band_frame
    df = band_frame(eng, scn.plan_year)
    by_month = df.pivot(index="month", columns="dest", values="share")
    for _m, row in by_month.iterrows():
        assert row.sum() == pytest.approx(1.0, abs=1e-12)
    keep = by_month.sort_index()[f"earned in CY {scn.plan_year}"].tolist()
    assert all(a > b for a, b in zip(keep, keep[1:]))


def test_walk_frame_carries_the_flow_and_the_race_reads_it(scn, combo):
    from src import engine

    from app.views.delivery import race_sentence, walk_frame
    flow = engine.lr_flow_by_month(scn.plan_year, combo[2])
    df = walk_frame(flow)
    assert len(df) == 24
    for i, r in enumerate(flow.rows):
        if r["lr"] is None:
            assert pd.isna(df.loc[i, "lr"])
        else:
            assert df.loc[i, "lr"] == pytest.approx(r["lr"])
    s = race_sentence(flow, combo[2].net_trend)
    assert ("break even" in s) or ("trend alone" in s)


# ------------------------------------------------------- page render (venv)

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_combo_page_renders_every_visual():
    """Open the shipped workbook through the real page: the selector adopts
    the first combo, the headless debounce renders synchronously, and every
    pane holds an object — with the error note EMPTY (a note here means a
    visual died and the page swallowed it)."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page

    session = PlanSession()
    page = combo_page.build(session)
    session.replace_config(importers.from_workbook(WB))

    assert page["combo_sel"].value                     # auto-adopted
    panes = page["panes"]
    assert panes["note"].object == ""
    for name in ("waterfall_p", "waterfall_p1", "earn_lag", "parallelogram",
                 "walk", "band"):
        assert panes[name].object is not None, f"{name} did not render"
    assert panes["race"].object                        # the verdict sentence
