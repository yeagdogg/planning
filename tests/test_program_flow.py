"""Program flow decomposition (D59): written-basis YoY legs of the program
as logged, the taken-only locked leg, and the plan-year ratio averages.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import datetime as dt
from dataclasses import replace

import pytest

from src.engine import (
    ComboInputs,
    ModInputs,
    RateChange,
    combined_flow_by_month,
    month_index,
    program_flow_by_month,
)

P = 2027

FLAT_MODS = ModInputs(m_ind=0.85, m0=0.85, m0_asof=dt.date(P - 1, 9, 30), m1=0.85)
SLOPED_MODS = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.89)


def _combo(changes=(), mods=SLOPED_MODS, **kw) -> ComboInputs:
    return ComboInputs(
        lr_proj=0.65, lr_basis="current", mods=mods,
        rate_changes=tuple(changes), **kw,
    )


def _row(res, year, month):
    mi = month_index(year, month)
    return next(r for r in res.rows if r["mi"] == mi)


def _prior(res, month):
    mi = month_index(P - 1, month)
    return next(r for r in res.prior_rows if r["mi"] == mi)


def test_anniversary_collapse_hand_constants():
    """Single taken +8% @ 7/1/(P-1), flat mods: the YoY rate leg is +8% until
    the filing's anniversary and exactly 0 after it (a 7/1 change day-blends
    at p = 1, so its own month is fully post-change)."""
    res = program_flow_by_month(
        P, _combo([RateChange(dt.date(P - 1, 7, 1), 0.08, "taken", True)],
                  mods=FLAT_MODS))
    for m in range(1, 7):
        row = _row(res, P, m)
        assert row["rate_leg"] == pytest.approx(0.08, abs=1e-12)
        assert row["mod_leg"] == pytest.approx(0.0, abs=1e-12)
        assert row["delivered"] == pytest.approx(0.08, abs=1e-12)
        assert row["locked_leg"] == pytest.approx(0.08, abs=1e-12)
        assert row["planned_residual"] == pytest.approx(0.0, abs=1e-12)
    for m in range(7, 13):
        assert _row(res, P, m)["rate_leg"] == pytest.approx(0.0, abs=1e-12)
    for m in range(1, 13):  # P+1: fully annualized, nothing left to deliver
        assert _row(res, P + 1, m)["delivered"] == pytest.approx(0.0, abs=1e-12)


def test_locked_leg_ignores_planned_in_mixed_month():
    """D58 case: a planned row EARLIER in the same cohort month as a taken row
    must not disturb the taken-only day-blend, and the locked series must be
    identical with or without the planned row."""
    taken_only = (
        RateChange(dt.date(P - 2, 7, 1), 0.04, "taken", True),
        RateChange(dt.date(P - 1, 6, 20), 0.02, "taken", False),
    )
    mixed = taken_only + (RateChange(dt.date(P - 1, 6, 5), 0.03, "planned", False),)
    res_mixed = program_flow_by_month(P, _combo(mixed))
    res_taken = program_flow_by_month(P, _combo(taken_only))
    # hand day-blend for June (P-1): first TAKEN change 6/20 -> p = 11/30
    p = 11 / 30
    expected_june = 1.02 / ((1 - p) + p * 1.02) - 1.0
    assert _row(res_mixed, P, 6)["locked_leg"] == pytest.approx(expected_june, abs=1e-12)
    for rm, rt in zip(res_mixed.rows, res_taken.rows):
        assert rm["locked_leg"] == pytest.approx(rt["locked_leg"], abs=1e-15)
    # sanity: the planned row DOES move the full rate leg in that month
    assert _row(res_mixed, P, 6)["rate_leg"] != pytest.approx(
        _row(res_taken, P, 6)["rate_leg"], abs=1e-6)


def test_achievement_and_midmonth_blend():
    """Planned +5% at 80% achievement enters at +4%, and a 4/15 effective date
    day-blends its own month at p = 16/30."""
    res_ach = program_flow_by_month(
        P, _combo([RateChange(dt.date(P, 4, 15), 0.05, "planned", False, 0.8)],
                  mods=FLAT_MODS))
    res_flat = program_flow_by_month(
        P, _combo([RateChange(dt.date(P, 4, 15), 0.04, "planned", False, 1.0)],
                  mods=FLAT_MODS))
    assert res_ach.rows == res_flat.rows
    p = 16 / 30
    assert _row(res_ach, P, 4)["rate_leg"] == pytest.approx(p * 0.04, abs=1e-12)
    assert _row(res_ach, P, 5)["rate_leg"] == pytest.approx(0.04, abs=1e-12)
    # no taken rows: locked leg 0, the whole leg is planned residual
    assert _row(res_ach, P, 4)["locked_leg"] == pytest.approx(0.0, abs=1e-12)
    assert _row(res_ach, P, 4)["planned_residual"] == pytest.approx(p * 0.04, abs=1e-12)


def test_net_mode_ignored_on_program_basis():
    """D59: program flow shows the log AS LOGGED — a net selection changes
    nothing here, and the delivered path is NOT the flat (1+x) assertion."""
    changes = (
        RateChange(dt.date(P - 2, 7, 1), 0.04, "taken", True),
        RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False),
        RateChange(dt.date(P, 4, 1), 0.05, "planned", False),
    )
    res_net = program_flow_by_month(P, _combo(changes, net_sel_p=0.10))
    res_off = program_flow_by_month(P, _combo(changes, net_sel_p=None))
    assert res_net.rows == res_off.rows
    assert res_net.avg_delivered_ratio == pytest.approx(
        res_off.avg_delivered_ratio, abs=1e-15)
    deliv = [r["delivered"] for r in res_net.rows[:12]]
    assert max(abs(d - 0.10) for d in deliv) > 0.01  # the gap is the feature


def test_mod_off_rate_only():
    combo = _combo([RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False)],
                   mod_adjustment_enabled=False)
    res = program_flow_by_month(P, combo)
    assert not res.mod_on and res.avg_mod_ratio is None
    for row in res.rows:
        assert row["mod_leg"] is None
        assert row["delivered"] == pytest.approx(row["rate_leg"], abs=1e-15)
    assert res.avg_delivered_ratio == pytest.approx(res.avg_rate_ratio, abs=1e-15)


def test_term_independent():
    """Written-basis YoY has no policy term in it — term 6 == term 12."""
    changes = [RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False)]
    res6 = program_flow_by_month(P, _combo(changes, term_months=6))
    res12 = program_flow_by_month(P, _combo(changes, term_months=12))
    assert res6.rows == res12.rows
    assert res6.avg_delivered_ratio == res12.avg_delivered_ratio


def test_seasonality_moves_only_averages():
    season = (0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5)
    changes = [RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False)]
    res_u = program_flow_by_month(P, _combo(changes))
    res_s = program_flow_by_month(P, _combo(changes, seasonality=season))
    for ru, rs in zip(res_u.rows, res_s.rows):
        for fld in ("rate_leg", "mod_leg", "delivered", "locked_leg"):
            assert ru[fld] == pytest.approx(rs[fld], abs=1e-15)
    assert res_u.avg_rate_ratio != pytest.approx(res_s.avg_rate_ratio, abs=1e-6)
    # the average IS the w-weighted mean of the monthly ratios
    hand = (sum(r["w"] * (1.0 + r["rate_leg"]) for r in res_s.rows[:12])
            / sum(r["w"] for r in res_s.rows[:12]))
    assert res_s.avg_rate_ratio == pytest.approx(hand, abs=1e-15)


CHG_A = (RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False),)
CHG_B = (RateChange(dt.date(P - 1, 4, 1), 0.03, "taken", False),
         RateChange(dt.date(P, 7, 1), 0.06, "planned", False))


def test_combined_single_combo_degenerates():
    combo = _combo(CHG_A, plan_ep=120.0)
    single = program_flow_by_month(P, combo)
    comb = combined_flow_by_month(P, [combo])
    for cr, sr in zip(comb.rows, single.rows[:12]):
        assert cr["rate_leg"] == pytest.approx(sr["rate_leg"], abs=1e-15)
        assert cr["mod_leg"] == pytest.approx(sr["mod_leg"], abs=1e-15)
        assert cr["delivered"] == pytest.approx(sr["delivered"], abs=1e-15)
    assert comb.avg_delivered_ratio == pytest.approx(single.avg_delivered_ratio,
                                                     abs=1e-15)


def test_combined_within_state_is_ep_only():
    """Shared seasonality (state-keyed) cancels out of the weights: the
    EP x w combination equals the plain EP-weighted mean of the legs."""
    season = (0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5)
    ca = _combo(CHG_A, seasonality=season, plan_ep=100.0)
    cb = _combo(CHG_B, seasonality=season, plan_ep=300.0)
    comb = combined_flow_by_month(P, [ca, cb])
    fa, fb = program_flow_by_month(P, ca), program_flow_by_month(P, cb)
    for j in range(12):
        hand = (100.0 * fa.rows[j]["delivered"] + 300.0 * fb.rows[j]["delivered"]) / 400.0
        assert comb.rows[j]["delivered"] == pytest.approx(hand, abs=1e-15)


def test_combined_mixed_modeff():
    """Mod-off combos are excluded from the mod leg but included in delivered
    (at their rate leg) — the honest book statistic."""
    ca = _combo(CHG_A, plan_ep=100.0)
    cb = _combo(CHG_B, mod_adjustment_enabled=False, plan_ep=100.0)
    comb = combined_flow_by_month(P, [ca, cb])
    fa, fb = program_flow_by_month(P, ca), program_flow_by_month(P, cb)
    for j in range(12):
        assert comb.rows[j]["mod_leg"] == pytest.approx(fa.rows[j]["mod_leg"],
                                                        abs=1e-15)
        hand = (fa.rows[j]["delivered"] + fb.rows[j]["rate_leg"]) / 2.0
        assert comb.rows[j]["delivered"] == pytest.approx(hand, abs=1e-15)
    # no mod-on combo at all -> mod leg is None
    both_off = combined_flow_by_month(
        P, [replace(ca, mod_adjustment_enabled=False), cb])
    assert all(r["mod_leg"] is None for r in both_off.rows)
    assert both_off.avg_mod_ratio is None


def test_combined_zero_ep_excluded():
    ca = _combo(CHG_A, plan_ep=100.0)
    ghost = _combo(CHG_B, plan_ep=0.0)
    ghost2 = _combo(CHG_B)                      # plan_ep None
    comb = combined_flow_by_month(P, [ca, ghost, ghost2])
    solo = combined_flow_by_month(P, [ca])
    assert comb.rows == solo.rows
    with pytest.raises(ValueError):
        combined_flow_by_month(P, [ghost, ghost2])


def test_combined_mix_effect_delivered_is_exact():
    """Under mix, EP-weighted rate x EP-weighted mod need not equal the
    EP-weighted delivered — delivered is the exact statistic (D60)."""
    mods_hot = ModInputs(m_ind=0.85, m0=0.80, m0_asof=dt.date(P - 1, 9, 30), m1=0.95)
    ca = _combo(CHG_A, mods=mods_hot, plan_ep=100.0)
    cb = _combo(CHG_B, plan_ep=900.0)
    comb = combined_flow_by_month(P, [ca, cb])
    gaps = [abs((1 + r["rate_leg"]) * (1 + r["mod_leg"]) - (1 + r["delivered"]))
            for r in comb.rows]
    assert max(gaps) > 1e-6                     # the mix effect is real...
    fa, fb = program_flow_by_month(P, ca), program_flow_by_month(P, cb)
    for j in range(12):                          # ...and delivered stays exact
        hand = (100.0 * fa.rows[j]["delivered"] + 900.0 * fb.rows[j]["delivered"]) / 1000.0
        assert comb.rows[j]["delivered"] == pytest.approx(hand, abs=1e-15)


def test_algebra_identities_rich_combo():
    """Locks the leg algebra the workbook cells reproduce: legs multiply to
    delivered, and locked x residual recombines to the full rate leg."""
    mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.90,
                     m_prior=0.84)
    changes = (
        RateChange(dt.date(P - 2, 7, 1), 0.04, "taken", True),
        RateChange(dt.date(P - 1, 6, 5), 0.03, "planned", False, 0.7),
        RateChange(dt.date(P - 1, 6, 20), 0.02, "taken", False),
        RateChange(dt.date(P, 4, 1), 0.05, "planned", False),
        RateChange(dt.date(P, 4, 20), -0.01, "taken", False),
    )
    res = program_flow_by_month(P, _combo(changes, mods=mods))
    for row in res.rows:
        assert (1.0 + row["rate_leg"]) * (1.0 + row["mod_leg"]) == pytest.approx(
            1.0 + row["delivered"], abs=1e-12)
        assert (1.0 + row["locked_leg"]) * (1.0 + row["planned_residual"]) == \
            pytest.approx(1.0 + row["rate_leg"], abs=1e-12)
    assert len(res.rows) == 24


# ---------------------------------------------------------------------------
# D68: the prior year — what is CURRENTLY flowing, alongside the plan year
# ---------------------------------------------------------------------------


def test_prior_rows_are_the_twelve_months_before_the_plan_year():
    res = program_flow_by_month(P, _combo(CHG_B))
    assert len(res.prior_rows) == 12
    assert [r["mi"] for r in res.prior_rows] == [
        month_index(P - 1, m) for m in range(1, 13)]
    # additive: the existing family must not have moved by a single index
    assert len(res.rows) == 24 and res.rows[0]["mi"] == month_index(P, 1)


def test_prior_year_hand_constants():
    """One taken +8% at 7/1/(P-2), flat mods. The year now flowing (P-1) sees
    +8% through June — its year-ago base is pre-change — then 0 once the base
    months are themselves post-change. The plan year sees nothing at all."""
    res = program_flow_by_month(
        P, _combo([RateChange(dt.date(P - 2, 7, 1), 0.08, "taken", True)],
                  mods=FLAT_MODS))
    for m in range(1, 7):
        assert _prior(res, m)["rate_leg"] == pytest.approx(0.08, abs=1e-12)
        assert _prior(res, m)["delivered"] == pytest.approx(0.08, abs=1e-12)
    for m in range(7, 13):
        assert _prior(res, m)["rate_leg"] == pytest.approx(0.0, abs=1e-12)
    for m in range(1, 13):
        assert _row(res, P, m)["rate_leg"] == pytest.approx(0.0, abs=1e-12)
    assert res.avg_delivered_ratio_prior == pytest.approx(1.04, abs=1e-12)
    assert res.avg_delivered_ratio == pytest.approx(1.0, abs=1e-12)


def test_prior_weights_equal_plan_year_weights_by_calendar_month():
    """Seasonality is a function of MONTH, so w repeats year over year.

    The workbook leans on exactly this: it publishes ONE epw family and reads
    it for both years' All-view aggregation (D68). If this ever stopped
    holding, that reuse would silently mis-weight the prior year."""
    season = (0.4, 0.6, 0.8, 1.0, 1.2, 1.4, 1.6, 1.4, 1.2, 1.0, 0.8, 0.6)
    res = program_flow_by_month(P, _combo(CHG_B, seasonality=season))
    for j in range(12):
        assert res.prior_rows[j]["w"] == pytest.approx(res.rows[j]["w"], abs=1e-15)


def test_prior_averages_are_w_weighted_means_of_prior_ratios():
    season = (0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5)
    res = program_flow_by_month(P, _combo(CHG_B, seasonality=season))
    for attr, fld in (("avg_rate_ratio_prior", "rate_leg"),
                      ("avg_mod_ratio_prior", "mod_leg"),
                      ("avg_delivered_ratio_prior", "delivered")):
        hand = (sum(r["w"] * (1.0 + r[fld]) for r in res.prior_rows)
                / sum(r["w"] for r in res.prior_rows))
        assert getattr(res, attr) == pytest.approx(hand, abs=1e-15)


def test_prior_year_is_term_independent_too():
    res6 = program_flow_by_month(P, _combo(CHG_B, term_months=6))
    res12 = program_flow_by_month(P, _combo(CHG_B, term_months=12))
    assert res6.prior_rows == res12.prior_rows


def test_prior_mod_off_rate_only():
    res = program_flow_by_month(P, _combo(CHG_A, mod_adjustment_enabled=False))
    assert res.avg_mod_ratio_prior is None
    for row in res.prior_rows:
        assert row["mod_leg"] is None
        assert row["delivered"] == pytest.approx(row["rate_leg"], abs=1e-15)


def test_combined_prior_ties_the_ep_weighted_per_combo_prior():
    ca = _combo(CHG_A, plan_ep=100.0)
    cb = _combo(CHG_B, plan_ep=300.0)
    comb = combined_flow_by_month(P, [ca, cb])
    fa, fb = program_flow_by_month(P, ca), program_flow_by_month(P, cb)
    assert len(comb.prior_rows) == 12
    assert comb.prior_rows[0]["mi"] == month_index(P - 1, 1)
    for j in range(12):
        wa, wb = 100.0 * fa.prior_rows[j]["w"], 300.0 * fb.prior_rows[j]["w"]
        hand = (wa * fa.prior_rows[j]["delivered"]
                + wb * fb.prior_rows[j]["delivered"]) / (wa + wb)
        assert comb.prior_rows[j]["delivered"] == pytest.approx(hand, abs=1e-15)
    hand_avg = (100.0 * fa.avg_delivered_ratio_prior
                + 300.0 * fb.avg_delivered_ratio_prior) / 400.0
    assert comb.avg_delivered_ratio_prior == pytest.approx(hand_avg, abs=1e-15)


def test_combined_prior_excludes_zero_ep():
    ca = _combo(CHG_A, plan_ep=100.0)
    comb = combined_flow_by_month(P, [ca, _combo(CHG_B, plan_ep=0.0)])
    solo = combined_flow_by_month(P, [ca])
    assert comb.prior_rows == solo.prior_rows
