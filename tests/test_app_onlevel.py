"""W2d/W3b: the on-leveling rectangle — the picture cannot drift from the math.

The ties: rebuild the engine's own closed-form numbers from the diagram's
per-year cells, to 1e-9 — ``continuous_earned_index`` for the program
geometry, and (W3b) ``continuous_net_index`` for the NET-path geometry,
for all three calendar years, on an annual-term AND a 6-month-term book.
The net picture is the user's own construction, verified: a restart
diagonal at each 1/1 (a sawtooth), anniversary echoes of the trailing
year's filings, and NO echo for anything older — pinned on a hand-built
case. The production anchor for net combos rides a 1e-3 tolerance: the
monthly engine interpolates mod anchors in day-ordinal space while the
continuous convention interpolates in month coordinates (the same
long-standing pair every continuous cross-check lives with).
"""
from __future__ import annotations

import datetime as dt
import math
import sys
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_IM = ROOT / "output" / "Plan_LR_Workbook_2027_Inland_Marine.xlsx"

_needs_panel = pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")


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


def _net_combo(wb_path):
    from src import engine

    from app import compute, importers
    scn = importers.from_workbook(wb_path)
    key = next(k for k in scn.combo_keys()
               if scn.row(k).get("netp") is not None)
    ci = compute.combo_inputs(scn, scn.row(key))
    return scn, key, ci, engine.run_bridge(scn.plan_year, ci, "monthly")


def _rebuild(fr, yr):
    """Independent re-aggregation: the year's exposure partition and its
    share-weighted level, straight off the frame's cells."""
    cells = [c for r in fr.regions for c in r["cells"] if c["year"] == yr]
    tot = sum(c["share"] for c in cells)
    lvl = sum(c["share"] * c["level"] for c in cells
              if c["level"] is not None)
    return tot, lvl


def _hand_net(netp=0.08, mods=None, mod_changes=(), old_change=True):
    """+5% taken 6/1/(P-2) [older than the echo horizon], +10% taken
    10/1/(P-1), net selection from 1/1/P. Flat mods unless given."""
    from src.engine import ComboInputs, ModInputs, RateChange, run_bridge
    mods = mods or ModInputs(m_ind=0.85, m0=0.85,
                             m0_asof=dt.date(2026, 9, 30), m1=0.85)
    changes = ((RateChange(dt.date(2025, 6, 1), 0.05, "taken",
                           considered=True),) if old_change else ()) + (
        RateChange(dt.date(2026, 10, 1), 0.10, "taken", considered=True),)
    ci = ComboInputs(mods=mods, rate_changes=changes, term_months=12,
                     lr_proj=0.65, net_sel_p=netp, mod_changes=mod_changes)
    return ci, run_bridge(2027, ci, "monthly")


# ------------------------------------------------------ program geometry

@pytest.mark.parametrize("wb", [WB_PROP, WB_IM],
                         ids=["annual-term", "six-month-term"])
def test_region_areas_rebuild_the_continuous_index(wb):
    from src import engine

    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(wb)
    fr = onlevel_frame(ci, res)
    for yr in (scn.plan_year - 1, scn.plan_year, scn.plan_year + 1):
        tot, lvl = _rebuild(fr, yr)
        assert tot == pytest.approx(1.0, abs=1e-9), yr       # partition
        want = engine.continuous_earned_index(ci.rate_changes, yr,
                                              ci.term_months)
        assert lvl / tot == pytest.approx(want, abs=1e-9), yr


def test_regions_follow_the_sorted_cumulative_product():
    from src.engine import month_coord, month_index

    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(WB_PROP)
    fr = onlevel_frame(ci, res)
    assert fr.window == (-12.0, 24.0)
    idx = 1.0
    seen = {r["index"] for r in fr.regions}
    assert 1.0 in seen or fr.regions[0]["index"] != 1.0  # base may clip out
    for _c, pct, _d in sorted(
            ((month_coord(rc.effective), rc.effective_pct, rc.effective)
             for rc in ci.rate_changes), key=lambda t: t[0]):
        idx *= 1.0 + pct
    assert max(seen) == pytest.approx(idx, rel=1e-12)   # last region = CRL-ish
    # diagonals sit exactly at the change coordinates, all kind='filing'
    base = float(month_index(scn.plan_year, 1))
    coords = sorted(month_coord(rc.effective) - base
                    for rc in ci.rate_changes)
    assert all(d["kind"] == "filing" for d in fr.diagonals)
    for d in fr.diagonals:
        c = d["x0"] - fr.term * d["y0"]
        assert any(abs(c - cc) < 1e-9 for cc in coords)


def test_window_24_drops_the_pm1_strip():
    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(WB_PROP)
    fr = onlevel_frame(ci, res, window=24)
    assert fr.window == (0.0, 24.0)
    assert [y["year"] for y in fr.years] == [scn.plan_year,
                                             scn.plan_year + 1]
    assert all(r["share_pm1"] == 0.0 for r in fr.regions)


# ---------------------------------------------------------- net geometry

@pytest.mark.parametrize("wb", [WB_PROP, WB_IM],
                         ids=["annual-term", "six-month-term"])
def test_net_regions_rebuild_the_net_index(wb):
    """The W3b headline: the NET picture rebuilds the engine's (fixed)
    continuous net index at 1e-9 for all three years — and anchors to the
    monthly production number at the honest cross-convention tolerance."""
    from src import engine

    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _net_combo(wb)
    p = scn.plan_year
    fr = onlevel_frame(ci, res)
    assert fr.net_mode
    for yr in (p - 1, p, p + 1):
        tot, lvl = _rebuild(fr, yr)
        assert tot == pytest.approx(1.0, abs=1e-9), yr
        want = engine.continuous_net_index(p, ci, yr)
        assert lvl / tot == pytest.approx(want, abs=1e-9), yr
        assert lvl / tot == pytest.approx(res.e_cy[yr], abs=1e-3), yr


def test_net_jump_set_and_restart_diagonals():
    """The user's construction, pinned: restarts at 1/1 (DOWN-steps when
    the book out-earned the selection), echoes at +12/+24 of the trailing
    year's filings only — a P-2 filing echoes NOTHING."""
    from app.views.onlevel import onlevel_frame
    ci, res = _hand_net()
    fr = onlevel_frame(ci, res)

    def solved(d):
        return d["x0"] - fr.term * d["y0"]

    by_kind: dict = {}
    for d in fr.diagonals:
        by_kind.setdefault(d["kind"], []).append(round(solved(d), 6))
    assert sorted(by_kind["restart"]) == [0.0, 12.0]
    assert sorted(by_kind["echo"]) == [9.0, 21.0]          # -3 + 12, + 24
    assert -3.0 in by_kind["filing"]
    # the 6/1/(P-2) filing: its diagonal may show, its echoes must NOT
    assert all(abs(c - (-7.0)) > 1e-6 and abs(c - 5.0) > 1e-6
               for c in by_kind["echo"])

    # the sawtooth: the restart band steps DOWN by (1+x)/(1+last filing)
    order = sorted(range(len(fr.regions)),
                   key=lambda i: min(c for c, _p in
                                     [(v[0] - fr.term * v[1], 0)
                                      for v in fr.regions[i]["poly"]]))
    seq = [fr.regions[i] for i in order]
    at0 = next(i for i, r in enumerate(seq)
               if r["kind"] == "restart" and "1/1/2027" in r["opened"])
    left, right = seq[at0 - 1], seq[at0]
    assert right["index"] < left["index"]
    assert right["index"] / left["index"] == pytest.approx(
        1.08 / 1.10, rel=1e-9)


def test_net_mod_reanchor_moves_the_picture():
    """D70a through the diagram: a mod log re-anchors HISTORY, which moves
    every band's q — and the rebuild still ties the (fixed) engine."""
    from src.engine import ModInputs, RateChange, continuous_net_index

    from app.views.onlevel import onlevel_frame
    mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30),
                     m1=0.89, m_end_prior=0.875)
    steps = (RateChange(dt.date(2027, 3, 1), 0.03, "taken",
                        considered=False),)
    ci_a, res_a = _hand_net(netp=0.10, mods=mods, mod_changes=steps,
                            old_change=False)
    ci_b, res_b = _hand_net(netp=0.10, mods=mods, old_change=False)
    fr_a = onlevel_frame(ci_a, res_a)
    fr_b = onlevel_frame(ci_b, res_b)
    for yr in (2026, 2027, 2028):
        for ci, fr in ((ci_a, fr_a), (ci_b, fr_b)):
            tot, lvl = _rebuild(fr, yr)
            assert lvl / tot == pytest.approx(
                continuous_net_index(2027, ci, yr), abs=1e-9), yr
    ra, rb = (_rebuild(fr_a, 2026)[1], _rebuild(fr_b, 2026)[1])
    assert abs(ra - rb) > 1e-4          # the re-anchor genuinely moves it


def test_net_mode_draws_the_net_path_and_program_flip():
    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _net_combo(WB_PROP)
    p = scn.plan_year
    fr = onlevel_frame(ci, res)
    assert fr.net_mode
    assert not any("would be a lie" in n or "NET SELECTION IN FORCE" in n
                   for n in fr.notes)
    assert any("NET PATH" in n for n in fr.notes)
    assert any(d["kind"] == "restart" for d in fr.diagonals)
    assert fr.years[1]["year"] == p
    assert fr.years[1]["e_cy"] == pytest.approx(res.e_cy[p])

    fr2 = onlevel_frame(ci, res, basis="program")
    assert not fr2.net_mode
    assert any("PROGRAM BASIS" in n for n in fr2.notes)
    assert all(d["kind"] == "filing" for d in fr2.diagonals)
    # annotations stay the PRODUCTION net numbers in both views
    assert fr2.years[1]["e_cy"] == pytest.approx(res.e_cy[p])


def test_year_annotations_are_production_numbers():
    from app.views.onlevel import onlevel_frame
    scn, _key, ci, res = _combo(WB_PROP)
    p = scn.plan_year
    fr = onlevel_frame(ci, res)
    assert [y["year"] for y in fr.years] == [p - 1, p, p + 1]
    assert fr.years[0]["a_rate"] is None          # no published A_rate P-1
    assert fr.years[0]["e_cy"] == res.e_cy[p - 1]
    assert fr.years[1]["a_rate"] == res.a_rate_p
    assert fr.years[2]["a_rate"] == res.a_rate_p1
    assert all(not math.isnan(y["e_cont"]) for y in fr.years)


# ------------------------------------------------------------- venv render

@_needs_panel
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


@_needs_panel
def test_onlevel_controls_window_and_basis():
    """The W3b interaction: the 24/36 window checkbox always renders; the
    basis flip appears only on net combos and swaps the geometry."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page

    session = PlanSession()
    page = combo_page.build(session)
    scn = importers.from_workbook(WB_PROP)
    session.replace_config(scn)

    assert page["ol_basis"].visible is False       # first combo is non-net
    page["ol_window"].value = False                # 24-month zoom
    assert page["panes"]["onlevel"].object is not None

    net_key = next(k for k in scn.combo_keys()
                   if scn.row(k).get("netp") is not None)
    page["combo_sel"].value = net_key
    assert page["ol_basis"].visible is True
    assert "Program vs asserted" in page["ol_note"].object
    page["ol_basis"].value = "Logged program (D65)"
    assert page["panes"]["onlevel"].object is not None
    assert "PROGRAM BASIS" in page["ol_note"].object


@_needs_panel
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


# ------------------------------------------------- W4c: indicated level

def test_indicated_levels_are_the_stated_arithmetic():
    """The system's FIRST computed indication (the workbook's indication
    fields are carry-through by design, D111), so the definitions live in
    one place and this pins both of them."""
    from app import compute, importers
    from app.views.onlevel import indicated_levels
    from src import engine
    scn = importers.from_workbook(WB_PROP)
    row = scn.lr_rows[0]
    ci = compute.combo_inputs(scn, row)
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    target = row["target"]

    rate, bridge = indicated_levels(ci, res, target)
    assert rate == pytest.approx(res.crl_ind * ci.lr_proj / target, abs=1e-12)
    assert bridge == pytest.approx(rate * res.a_mod_p * ci.a_other, abs=1e-12)
    # the full bridge asks rate to do LESS exactly when the mod path helps
    assert (bridge > rate) == (res.a_mod_p * ci.a_other > 1.0)


def test_indicated_level_dashes_rather_than_guesses():
    """No target, no indication — and never a divide by zero."""
    from app import compute, importers
    from app.views.onlevel import indicated_levels, onlevel_frame
    from src import engine
    scn = importers.from_workbook(WB_PROP)
    row = scn.lr_rows[0]
    ci = compute.combo_inputs(scn, row)
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    for bad in (None, 0.0, -0.1, "", "0.6"):
        assert indicated_levels(ci, res, bad) == (None, None), bad
    fr = onlevel_frame(ci, res)                     # target defaults to None
    assert fr.ind_rate is None and fr.ind_bridge is None
    assert not any("INDICATED LEVEL" in n for n in fr.notes)


def test_the_frame_carries_the_level_and_discloses_its_footing():
    from app import compute, importers
    from app.views.onlevel import onlevel_frame
    from src import engine
    scn = importers.from_workbook(WB_PROP)
    row = scn.lr_rows[0]
    ci = compute.combo_inputs(scn, row)
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    fr = onlevel_frame(ci, res, target=row["target"])
    assert fr.ind_rate is not None and fr.ind_bridge is not None
    notes = " ".join(fr.notes)
    assert "INDICATED LEVEL" in notes and "rate-only" in notes
    # the honest footing: CRL_ind counts only the CONSIDERED changes while
    # the bands compound every logged one
    assert "flagged IN the indication" in notes


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("holoviews") is None,
    reason="holoviews not installed (system interpreter — app venv runs this)")
def test_both_renderers_show_the_level_where_it_belongs():
    """The on-level diagram's y-axis is TERM ELAPSED, so it carries a
    labeled comparison; earn_lag's y-axis IS the rate index, so it carries
    the literal rules."""
    from app import compute, importers
    from app.views import earning as v_earning
    from app.views.onlevel import onlevel, onlevel_frame
    from src import engine
    scn = importers.from_workbook(WB_PROP)
    row = scn.lr_rows[0]
    ci = compute.combo_inputs(scn, row)
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    fr = onlevel_frame(ci, res, target=row["target"])

    import holoviews as hv
    plot = onlevel(ci, res, frame=fr)
    texts = " ".join(
        str(v) for el in plot.traverse(lambda e: e, [hv.Labels])
        for v in el.dimension_values("text"))
    assert "indicated rate level" in texts
    assert f"{fr.ind_rate:.4f}" in texts and f"{fr.ind_bridge:.4f}" in texts
    # the per-year annotation carries the gap in the exhibit's own terms
    assert "vs indicated" in texts

    lag = v_earning.earn_lag(res, ind_rate=fr.ind_rate,
                             ind_bridge=fr.ind_bridge)
    lines = [el for el in lag if isinstance(el, hv.HLine)]
    assert len(lines) == 2
    assert {round(float(el.data), 10) for el in lines} == {
        round(fr.ind_rate, 10), round(fr.ind_bridge, 10)}
    # ...and without a target the chart is exactly what it was before
    assert not [el for el in v_earning.earn_lag(res)
                if isinstance(el, hv.HLine)]


def test_coinciding_lines_say_why_rather_than_look_broken():
    """On a net combo the engine pins A_mod = 1, so the full-bridge line
    lands exactly on the rate-only one. A reader seeing one line where the
    note promised two deserves the reason, not a mystery."""
    from app import compute, importers
    from app.views.onlevel import onlevel_frame
    from src import engine
    scn = importers.from_workbook(WB_PROP)
    row = next(r for r in scn.lr_rows if r.get("netp") is not None)
    ci = compute.combo_inputs(scn, row)
    res = engine.run_bridge(scn.plan_year, ci, "monthly")
    fr = onlevel_frame(ci, res, target=row["target"])
    assert fr.ind_bridge == pytest.approx(fr.ind_rate, abs=1e-12)
    notes = " ".join(fr.notes)
    assert "COINCIDE" in notes and "net mode pins A_mod = 1" in notes
