"""W3d: the per-combo flow depth — arrangement pinned to the engine.

The dashboard frame must be a VERBATIM read of the EngineResult series
and of ``program_flow_by_month`` rows (identity of arrangement, 1e-15);
runway re-derives independently; the carryover ledger's shares are
brute-forced from the engine's own w(k)·ec(k, year) masses with the
day-blend done by hand, and its totals reconcile to
``res.yoy_earned_p1`` to floating point. The reconciliation strip's
identity check must hold exactly (the lrf_ident twin).
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


def _setup(wb=WB_PROP, want_changes=2, net=False):
    from src import engine

    from app import compute, importers
    scn = importers.from_workbook(wb)
    for key in scn.combo_keys():
        row = scn.row(key)
        ci = compute.combo_inputs(scn, row)
        if net and not ci.net_mode:
            continue
        if not net and (ci.net_mode or len(ci.rate_changes) < want_changes):
            continue
        res = engine.run_bridge(scn.plan_year, ci, "monthly")
        pf = engine.program_flow_by_month(scn.plan_year, ci)
        return scn, key, ci, res, pf
    raise AssertionError("no suitable combo")


# ------------------------------------------------------- dashboard frame

def test_dashboard_series_are_engine_result_verbatim():
    from app.views.flow import dashboard_frame
    scn, _k, ci, res, pf = _setup()
    df = dashboard_frame(res, pf, ci)
    assert len(df) == 36
    for _i, r in df.iterrows():
        mi = r["mi"]
        assert r["written"] == res.inforce_index[res.cohort_mi.index(mi)]
        assert r["earned"] == res.earned_index_m[
            res.earned_month_mi.index(mi)]
        assert r["runway"] == pytest.approx(
            r["written"] / r["earned"] - 1.0, abs=1e-12)


def test_program_leg_columns_are_pf_rows_verbatim():
    from app.views.flow import _LEG_KEYS, dashboard_frame
    scn, _k, ci, res, pf = _setup()
    df = dashboard_frame(res, pf, ci)
    legs = list(pf.prior_rows) + list(pf.rows)
    for i in range(36):
        for k in _LEG_KEYS:
            assert df.iloc[i][k] == legs[i].get(k), (i, k)
    # YoY of the frame's own earned series, independently
    for i in range(12, 36):
        want = df["earned"].iloc[i] / df["earned"].iloc[i - 12] - 1.0
        assert df["yoy_earned_rate"].iloc[i] == pytest.approx(want,
                                                             abs=1e-12)


def test_mod_columns_none_when_off_or_net():
    from app import compute, importers
    from app.views.flow import dashboard_frame
    from src import engine

    scn, _k, ci, res, pf = _setup(net=True)
    df = dashboard_frame(res, pf, ci)
    assert df["wmod"].iloc[0] is None and df["price_e"].iloc[0] is None
    # net rows carry Δ vs the year's own assertion
    p = scn.plan_year
    for _i, r in df.iterrows():
        if r["year"] < p or r["delivered"] is None:
            assert r["delta_vs_net"] is None or r["year"] >= p
            continue
        x = ci.net_sel_p if r["year"] == p else ci.net_x1
        assert r["delta_vs_net"] == pytest.approx(r["delivered"] - x,
                                                  abs=1e-15)

    import dataclasses
    scn2, _k2, ci2, _res2, _pf2 = _setup()
    ci_off = dataclasses.replace(ci2, mod_adjustment_enabled=False)
    res_off = engine.run_bridge(scn2.plan_year, ci_off, "monthly")
    pf_off = engine.program_flow_by_month(scn2.plan_year, ci_off)
    df_off = dashboard_frame(res_off, pf_off, ci_off)
    assert df_off["wmod"].iloc[0] is None
    assert not pf_off.mod_on
    assert all(r["mod_leg"] is None for _j, r in df_off.iterrows())


# ------------------------------------------------------- carryover ledger

@pytest.mark.parametrize("wb", [WB_PROP, WB_IM],
                         ids=["annual-term", "six-month-term"])
def test_ledger_shares_brute_forced_and_totals_reconcile(wb):
    from src import engine

    from app.views.flow import ledger_frame
    scn, _k, ci, res, _pf = _setup(wb)
    eng = engine.MonthlyEngine(scn.plan_year, ci)
    ldf, totals = ledger_frame(eng, ci, res)
    p = scn.plan_year
    assert len(ldf) == len(ci.rate_changes)

    for r in ldf.itertuples():
        for yr, got in ((p, r.share_p), (p + 1, r.share_p1)):
            num = den = 0.0
            mo = r.eff.year * 12 + r.eff.month - 1
            eom = (dt.date(r.eff.year + (r.eff.month == 12),
                           r.eff.month % 12 + 1, 1) - dt.timedelta(days=1))
            frac = (eom - r.eff + dt.timedelta(days=1)).days / eom.day
            for k in eng.cohorts:
                m = eng.w(k) * eng.ec(k, yr)
                den += m
                if k > mo:
                    num += m
                elif k == mo:
                    num += frac * m
            assert got == pytest.approx(num / den, abs=1e-12), (r.seq, yr)
        assert r.contribution == pytest.approx(
            math.exp(math.log(1 + r.pct) * (r.share_p1 - r.share_p)) - 1,
            abs=1e-15)

    # (1 + combined) x (1 + residual) - 1 == the engine's own carryover
    assert (1 + totals["combined"]) * (1 + totals["residual"]) - 1 == \
        pytest.approx(res.yoy_earned_p1, abs=1e-14)
    assert totals["total"] == res.yoy_earned_p1          # linked, not derived


def test_ledger_hand_case_full_month_action():
    """One +10% action effective 1/1/P on a hand-built combo: everything
    written on/after Jan P carries it — share_p is the plan-year cohorts'
    entire share of CY P earning, day-blend = 1.0 on the first."""
    from src.engine import ComboInputs, ModInputs, MonthlyEngine, \
        RateChange, month_index, run_bridge

    from app.views.flow import ledger_frame
    mods = ModInputs(m_ind=0.85, m0=0.85, m0_asof=dt.date(2026, 9, 30),
                     m1=0.85)
    ci = ComboInputs(mods=mods, rate_changes=(
        RateChange(dt.date(2027, 1, 1), 0.10, "taken", considered=True),),
        term_months=12, lr_proj=0.65)
    res = run_bridge(2027, ci, "monthly")
    eng = MonthlyEngine(2027, ci)
    ldf, totals = ledger_frame(eng, ci, res)
    jan = month_index(2027, 1)
    for yr, col in ((2027, "share_p"), (2028, "share_p1")):
        den = sum(eng.w(k) * eng.ec(k, yr) for k in eng.cohorts)
        num = sum(eng.w(k) * eng.ec(k, yr) for k in eng.cohorts if k >= jan)
        assert ldf[col].iloc[0] == pytest.approx(num / den, abs=1e-12)
    assert totals["net_mode"] is False


def test_ledger_net_mode_flag():
    from src import engine

    from app.views.flow import ledger_frame
    scn, _k, ci, res, _pf = _setup(net=True)
    eng = engine.MonthlyEngine(scn.plan_year, ci)
    _ldf, totals = ledger_frame(eng, ci, res)
    assert totals["net_mode"] is True


# ------------------------------------------------- walk table + reconciliation

def test_walk_table_is_flow_rows_verbatim():
    from src import engine

    from app.views.delivery import walk_table_frame
    scn, key, ci, _res, _pf = _setup()
    flow = engine.lr_flow_by_month(scn.plan_year, ci)
    target = 0.62
    df = walk_table_frame(flow, target)
    assert len(df) == 24
    for i, r in enumerate(flow.rows):
        assert df.iloc[i]["lr"] == r["lr"]
        assert df.iloc[i]["a_rate"] == r["a_rate"]
        assert df.iloc[i]["den"] == r["den"]
        if r["lr"] is None:
            assert df.iloc[i]["vs_target"] is None       # dead months dash
        else:
            assert df.iloc[i]["vs_target"] == pytest.approx(
                (r["lr"] - target) * 100.0, abs=1e-12)


def test_recon_identity_holds_to_floating_point():
    from src import engine

    from app.views.delivery import headroom, recon_rows
    scn, _key, ci, _res, _pf = _setup()
    flow = engine.lr_flow_by_month(scn.plan_year, ci)
    ident = flow.mean_p - flow.head_p * (1 + flow.cog_tilt_p) \
        * (1 + flow.convexity_p) / (1 + flow.cov_rate_price)
    assert ident == pytest.approx(0.0, abs=1e-12)
    labels = [k for k, _v in recon_rows(flow)]
    assert any("identity check" in k for k in labels)

    pts, bad = headroom(flow, flow.breakeven_trend + 0.01)
    assert pts == pytest.approx(1.0, abs=1e-9) and bad
    pts2, bad2 = headroom(flow, flow.breakeven_trend - 0.02)
    assert pts2 == pytest.approx(-2.0, abs=1e-9) and not bad2


# ------------------------------------------------------------- venv render

@_needs_panel
def test_combo_page_flow_tab_renders():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page

    session = PlanSession()
    page = combo_page.build(session)
    scn = importers.from_workbook(WB_PROP)
    session.replace_config(scn)

    assert page["panes"]["flow_index"].object is not None
    assert page["panes"]["flow_runway"].object is not None
    assert page["panes"]["flow_delivery"].object is not None
    assert "Carryover ledger" in page["panes"]["flow_ledger"].object
    assert "breakeven" in page["panes"]["walk_chips"].object
    assert len(page["walk_tbl"].value) == 24
    assert "identity check" in page["panes"]["recon"].object

    # a NET combo must not blank the tab — mod/price collapses to a note
    net_key = next(k for k in scn.combo_keys()
                   if scn.row(k).get("netp") is not None)
    page["combo_sel"].value = net_key
    assert page["panes"]["flow_delivery"].object is not None
    assert "net path already combines" in page["panes"]["flow_note"].object
    assert page["panes"]["flow_mod"].visible is False
