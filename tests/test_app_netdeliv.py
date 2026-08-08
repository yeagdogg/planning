"""W3f: the Net Delivery microscope — the engine's closed forms, arranged.

Microscope rows are ``net_delivery_by_month`` verbatim; at the SUGGESTED
change the w-weighted delivered-at-projected-mods equals the target at
1e-12 (the closed form's own identity, re-aggregated independently); the
log-scaled shares stack to the target exactly per month; defaults adopt
the planned filing (MULTI-PLANNED flagged) and an out-of-year date
surfaces the engine's words rather than a blank tab.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"

_needs_panel = pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")


def _net(wb=WB_PROP):
    from src import engine

    from app import compute, importers
    scn = importers.from_workbook(wb)
    key = next(k for k in scn.combo_keys()
               if scn.row(k).get("netp") is not None)
    ci = compute.combo_inputs(scn, scn.row(key))
    return scn, key, ci


def test_microscope_rows_are_engine_verbatim_and_defaults_adopt():
    from src import engine

    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    p = scn.plan_year
    mic = microscope(p, ci)
    planned = engine.planned_change_in_plan_year(p, ci)
    if planned:
        assert mic["eff_date"] == planned[0]
        assert mic["r"] == planned[1]
        assert mic["source"].startswith("planned filing")
        assert mic["multi_planned"] == planned[3]
    else:
        assert mic["eff_date"] == dt.date(p, 4, 1)
        assert mic["source"] == "suggested"
    rows = engine.net_delivery_by_month(p, ci, mic["eff_date"], mic["r"])
    df = mic["rows"]
    assert len(df) == 12
    for i, rr in enumerate(rows):
        assert df.iloc[i]["rate_leg"] == rr["rate_leg"]
        assert df.iloc[i]["price_req"] == rr["price_leg_required"]
        assert df.iloc[i]["m_proj"] == rr["m_proj"]
        assert df.iloc[i]["delivered"] == rr["delivered_at_proj"]


def test_suggested_change_delivers_the_target():
    """Independent re-aggregation of the closed form's own promise: at
    r*, Σ w·(1+delivered_at_proj) / Σ w − 1 == the target."""
    from src import engine

    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    p = scn.plan_year
    mic = microscope(p, ci)
    rows = engine.net_delivery_by_month(p, ci, mic["eff_date"],
                                        mic["suggested"])
    num = sum(r["w"] * (1.0 + r["delivered_at_proj"]) for r in rows)
    den = sum(r["w"] for r in rows)
    assert num / den - 1.0 == pytest.approx(ci.net_sel_p, abs=1e-12)


def test_shares_stack_exactly_to_the_target():
    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    mic = microscope(scn.plan_year, ci)
    df = mic["rows"]
    if df["share_rate"].iloc[0] is None:
        pytest.skip("mod-off combo — shares undefined")
    for _i, r in df.iterrows():
        assert r["share_rate"] + r["share_price"] == pytest.approx(
            ci.net_sel_p, abs=1e-12)


def test_p1_twin_and_target_source():
    from src import engine

    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    p = scn.plan_year
    mic = microscope(p, ci)
    p1 = engine.net_delivery_by_month_p1(p, ci, mic["eff_date"], mic["r"])
    assert len(mic["p1"]) == 12
    for i, rr in enumerate(p1):
        assert mic["p1"].iloc[i]["rate_leg"] == rr["rate_leg"]
        assert mic["p1"].iloc[i]["m_base"] == rr["m_base"]
    assert mic["x1"] == p1[0]["x"]
    assert mic["p1_src"] == ("entered" if ci.net_sel_p1 is not None
                             else "carried from P")


def test_mod_off_has_no_mod_ask_and_override_wins():
    import dataclasses

    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    p = scn.plan_year
    off = dataclasses.replace(ci, mod_adjustment_enabled=False)
    mic = microscope(p, off)
    assert mic["step"] is None
    assert "rate-only" in (mic["step_reason"] or "")
    assert mic["m1_prime"] is None and "n/a" in mic["m1_reason"]
    assert mic["rows"]["price_req"].iloc[0] is None

    mic2 = microscope(p, ci, filed_override=0.07)
    assert mic2["r"] == 0.07 and mic2["source"] == "override"


def test_out_of_year_date_surfaces_the_engines_words():
    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    with pytest.raises(ValueError):
        microscope(scn.plan_year, ci, eff_date=dt.date(scn.plan_year - 1,
                                                       6, 1))


def test_required_mod_step_surfaces_feasibility():
    from app.views.netdeliv import microscope
    scn, _key, ci = _net()
    mic = microscope(scn.plan_year, ci)
    if mic["step"] is None:
        pytest.skip(mic["step_reason"] or "no step available")
    s = mic["step"]
    assert isinstance(s.feasible, bool)
    assert s.directed_equivalent == pytest.approx(
        s.required_step / mic["ach"], abs=1e-12)


# ------------------------------------------------------------- venv render

@_needs_panel
def test_delivery_tab_solves_for_net_combos_only():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page

    session = PlanSession()
    page = combo_page.build(session)
    scn = importers.from_workbook(WB_PROP)
    session.replace_config(scn)

    nd = page["nd"]
    assert nd["col"].visible is False          # first combo is non-net
    assert "explicit program" in nd["hint"].object

    net_key = next(k for k in scn.combo_keys()
                   if scn.row(k).get("netp") is not None)
    page["combo_sel"].value = net_key
    assert nd["col"].visible is True
    assert "solving for" in nd["chips"].object
    assert len(nd["table"].value) == 12
    assert len(nd["p1"].value) == 12
    assert "carryover" in nd["p1_cap"].object

    nd["override"].value = "7%"                # override re-solves
    assert "override" in nd["chips"].object
    nd["override"].value = ""

    nd["eff"].value = None                     # reset to default adoption


@_needs_panel
def test_flow_page_suggest_column_gates():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import flow as flow_page

    session = PlanSession()
    page = flow_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    page["filters"]["bu"].value = ["SBA"]      # KY resolves to one net combo
    page["on_show"]()
    ndf = page["nd_tbl"].value
    ky = ndf.loc[ndf["state"] == "KY"].iloc[0]
    assert ky["suggest"] is not None
    others = ndf.loc[(ndf["netcnt"] == 0)
                     & (ndf["state"] != ndf.iloc[-1]["state"])]
    assert all(v is None for v in others["suggest"])
    assert ndf.iloc[-1]["suggest"] is None     # never aggregated
