"""W2d: the on-leveling rectangle — the picture cannot drift from the math.

The tie: rebuild the engine's own closed-form parallelogram number
(``continuous_earned_index``) from the diagram's region areas, to 1e-9,
for both calendar years, on an annual-term AND a 6-month-term combo. The
geometry lives in pure clipped polygons; if it ever disagrees with the
engine, the diagram is lying and these tests say so.
"""
from __future__ import annotations

import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_IM = ROOT / "output" / "Plan_LR_Workbook_2027_Inland_Marine.xlsx"


def _combo(wb_path, want_changes=2):
    from src import engine

    from app import compute, importers
    scn = importers.from_workbook(wb_path)
    for key in scn.combo_keys():
        ci = compute.combo_inputs(scn, scn.row(key))
        if len(ci.rate_changes) >= want_changes and not ci.net_mode:
            return scn, key, ci, engine.run_bridge(scn.plan_year, ci,
                                                   "monthly")
    raise AssertionError("no suitable combo")


@pytest.mark.parametrize("wb", [WB_PROP, WB_IM],
                         ids=["annual-term", "six-month-term"])
def test_region_areas_rebuild_the_continuous_index(wb):
    from src import engine

    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(wb)
    fr = onlevel_frame(ci, res)
    for j, yr in enumerate((scn.plan_year, scn.plan_year + 1)):
        shares = [(r["share_p"], r["index"]) if j == 0
                  else (r["share_p1"], r["index"]) for r in fr.regions]
        tot = sum(s for s, _i in shares)
        assert tot == pytest.approx(1.0, abs=1e-9)      # partition
        rebuilt = sum(s * i for s, i in shares) / tot
        want = engine.continuous_earned_index(ci.rate_changes, yr,
                                              ci.term_months)
        assert rebuilt == pytest.approx(want, abs=1e-9), yr


def test_regions_follow_the_sorted_cumulative_product():
    from src.engine import month_coord, month_index

    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(WB_PROP)
    fr = onlevel_frame(ci, res)
    idx = 1.0
    seen = {r["index"] for r in fr.regions}
    assert 1.0 in seen or fr.regions[0]["index"] != 1.0  # base may clip out
    for _c, pct, _d in sorted(
            ((month_coord(rc.effective), rc.effective_pct, rc.effective)
             for rc in ci.rate_changes), key=lambda t: t[0]):
        idx *= 1.0 + pct
    assert max(seen) == pytest.approx(idx, rel=1e-12)   # last region = CRL-ish
    # diagonals sit exactly at the change coordinates
    base = float(month_index(scn.plan_year, 1))
    coords = sorted(month_coord(rc.effective) - base
                    for rc in ci.rate_changes)
    for d in fr.diagonals:
        c = d["x0"] - fr.term * d["y0"]
        assert any(abs(c - cc) < 1e-9 for cc in coords)


def test_net_mode_draws_the_program_with_the_note():
    from src import engine

    from app import compute, importers
    from app.views.onlevel import onlevel_frame
    scn = importers.from_workbook(WB_PROP)
    key = next(k for k in scn.combo_keys()
               if scn.row(k).get("netp") is not None)
    ci = compute.combo_inputs(scn, scn.row(key))
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    fr = onlevel_frame(ci, res)
    assert fr.net_mode
    assert any("NET SELECTION" in n for n in fr.notes)
    assert fr.years[0]["e_cy"] == pytest.approx(res.e_cy[scn.plan_year])


def test_year_annotations_are_production_numbers():
    from app.views.onlevel import onlevel_frame
    _scn, _key, ci, res = _combo(WB_PROP)
    fr = onlevel_frame(ci, res)
    assert fr.years[0]["e_cy"] == res.e_cy[res.plan_year]
    assert fr.years[0]["a_rate"] == res.a_rate_p
    assert fr.years[1]["a_rate"] == res.a_rate_p1
    assert not math.isnan(fr.years[1]["e_cont"])


# ------------------------------------------------------------- venv render

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_combo_page_renders_onlevel_first_and_chips():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page

    session = PlanSession()
    page = combo_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))

    assert page["panes"]["onlevel"].object is not None
    chips = page["panes"]["chips"].object
    assert "plan LR" in chips and "vs target" in chips
    assert page["panes"]["note"].object == ""


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_book_table_carries_the_chrome_styles():
    from app import importers, summary
    from app.glue.session import PlanSession
    from app.pages import book as book_page

    session = PlanSession()
    page = book_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    table = page["table"]
    assert table.style is not None
    css = summary.three_color_scale(table.value["lr_p"])
    assert any(css)                                     # scale has colors
