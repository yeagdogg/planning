"""Net delivery decomposition (D57): closed-form suggestion, affine rate leg,
dual M_1' solve, and the live-row filter semantics.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import datetime as dt
from dataclasses import replace

import pytest

from src.engine import (
    ComboInputs,
    ModInputs,
    MonthlyEngine,
    RateChange,
    _live_changes,
    month_index,
    net_delivery_by_month,
    net_delivery_components,
    net_program_plan_lr,
    required_m1,
    run_bridge,
    solve_rate_for_target,
    suggest_net_rate,
)

P = 2027


def _combo(changes=(), net=0.10, mods=None, **kw) -> ComboInputs:
    mods = mods or ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.89)
    return ComboInputs(
        lr_proj=0.65, lr_basis="current", mods=mods,
        rate_changes=tuple(changes), net_sel_p=net, **kw,
    )


SHOWCASE = [
    RateChange(dt.date(P - 2, 7, 1), 0.04, "taken", True),
    RateChange(dt.date(P - 1, 10, 1), 0.10, "taken", False),
    RateChange(dt.date(P - 1, 7, 1), 0.065, "planned", False),   # live (pre-1/1/P)
    RateChange(dt.date(P, 4, 1), 0.05, "planned", False),        # superseded
]


def _delivered_avg(comp, r):
    """Written-weighted delivered net with mods on the projected path."""
    num = sum(
        comp.w[j] * (comp.a[j] + comp.b[j] * r)
        * ((comp.mproj[j] / comp.mw12[j]) if comp.mod_on else 1.0) / comp.w12[j]
        for j in range(12))
    return num / comp.sum_w


def test_live_filter_keeps_pre_plan_planned_rows():
    live = _live_changes(P, _combo(SHOWCASE))
    effs = {rc.effective for rc in live}
    assert dt.date(P - 1, 7, 1) in effs          # planned before 1/1/P stays live
    assert dt.date(P, 4, 1) not in effs          # planned on/after 1/1/P superseded
    assert dt.date(P - 1, 10, 1) in effs


@pytest.mark.parametrize("eff", [
    dt.date(P, 1, 1), dt.date(P, 4, 1), dt.date(P, 5, 15), dt.date(P, 12, 31),
    dt.date(P, 10, 15),   # co-occupies October with nothing; anniversary month
])
def test_suggested_rate_reproduces_target(eff):
    combo = _combo(SHOWCASE)
    r, comp = suggest_net_rate(P, combo, eff)
    assert _delivered_avg(comp, r) == pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


@pytest.mark.parametrize("eff", [
    dt.date(P, 3, 1),
    dt.date(P, 5, 15),
    dt.date(P, 6, 20),    # co-occupies June with a live taken row (D31 case)
    dt.date(P, 6, 5),     # earlier than the June taken row -> p splits at D
])
@pytest.mark.parametrize("r", [-0.03, 0.0, 0.05, 0.12])
def test_affine_equals_brute_force_injection(eff, r):
    changes = SHOWCASE + [RateChange(dt.date(P, 6, 10), 0.02, "taken", False)]
    combo = _combo(changes)
    comp = net_delivery_components(P, combo, eff)
    live = _live_changes(P, combo)
    brute = MonthlyEngine(
        P, combo, changes_override=live + (RateChange(eff, r, "taken", False),))
    for j in range(12):
        mi = month_index(P, 1) + j
        assert comp.a[j] + comp.b[j] * r == pytest.approx(
            brute.written_index(mi), abs=1e-12), f"month {j + 1}"


def test_mod_off_rate_only():
    combo = _combo(SHOWCASE, mod_adjustment_enabled=False)
    r, comp = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    assert not comp.mod_on and comp.warnings
    assert _delivered_avg(comp, r) == pytest.approx(1.10, abs=1e-12)
    with pytest.raises(ValueError):
        required_m1(P, combo, dt.date(P, 4, 1), r)


@pytest.mark.parametrize("mods", [
    ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.89),
    ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.89,
              m_prior=0.85),
    # as-of INSIDE the plan year: early plan months sit on the prior segment
    ModInputs(m_ind=0.85, m0=0.87, m0_asof=dt.date(P, 6, 30), m1=0.90,
              m_prior=0.84),
])
def test_required_m1_delivers_target_on_new_path(mods):
    combo = _combo(SHOWCASE, mods=mods)
    eff = dt.date(P, 4, 1)
    r = 0.03                                   # a deliberately-too-small filing
    m1p = required_m1(P, combo, eff, r)
    comp = net_delivery_components(P, combo, eff)
    new_path = MonthlyEngine(
        P, replace(combo, mods=replace(mods, m1=m1p, m2=None)),
        changes_override=_live_changes(P, combo))
    num = sum(comp.w[j] * (comp.a[j] + comp.b[j] * r)
              * new_path.written_mod(comp.months[j]) / comp.mw12[j] / comp.w12[j]
              for j in range(12))
    assert num / comp.sum_w == pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


def test_term_six_same_definition():
    combo = _combo(SHOWCASE, term_months=6)
    r, comp = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    assert _delivered_avg(comp, r) == pytest.approx(1.10, abs=1e-12)


def test_seasonality_weights_enter_the_solve():
    season = (0.5, 0.5, 0.5, 0.5, 0.5, 0.5, 1.5, 1.5, 1.5, 1.5, 1.5, 1.5)
    r_u, _ = suggest_net_rate(P, _combo(SHOWCASE), dt.date(P, 7, 1))
    r_s, comp_s = suggest_net_rate(
        P, _combo(SHOWCASE, seasonality=season), dt.date(P, 7, 1))
    assert r_u != pytest.approx(r_s, abs=1e-6)   # weights genuinely matter
    assert _delivered_avg(comp_s, r_s) == pytest.approx(1.10, abs=1e-12)


def test_guards():
    with pytest.raises(ValueError):              # D outside the plan year
        net_delivery_components(P, _combo(SHOWCASE), dt.date(P - 1, 10, 1))
    with pytest.raises(ValueError):              # no net selection
        net_delivery_components(P, _combo(SHOWCASE, net=None), dt.date(P, 4, 1))
    # zero written weight on/after D -> C1 = 0 -> no leverage
    season = tuple([1.0] * 11 + [0.0])
    with pytest.raises(ValueError):
        suggest_net_rate(P, _combo(SHOWCASE, seasonality=season),
                         dt.date(P, 12, 15))


def test_exact_equivalence_no_history_flat_mods():
    """With no history, flat mods, and D = 1/1: r* = x exactly and BOOKING the
    program reproduces the net-mode plan LR to machine precision."""
    mods = ModInputs(m_ind=0.85, m0=0.85, m0_asof=dt.date(P - 1, 9, 30), m1=0.85)
    combo = _combo((), mods=mods)
    r, comp = suggest_net_rate(P, combo, dt.date(P, 1, 1))
    assert r == pytest.approx(0.10, abs=1e-12)
    m1p = required_m1(P, combo, dt.date(P, 1, 1), r)
    assert m1p == pytest.approx(0.85, abs=1e-9)  # flat path stays flat
    lr_prog = net_program_plan_lr(P, combo, dt.date(P, 1, 1), r, m1p)
    lr_net = run_bridge(P, combo, "monthly").cy_lr_p
    assert lr_prog == pytest.approx(lr_net, abs=1e-12)


def _brute_force_solver_r(combo, target, eff, lo=-0.4, hi=0.4):
    """Numerically invert the D13 counterfactual (taken rows + ONE planned
    action at eff) through the full forward engine — the independent yardstick
    for the closed form."""
    taken = tuple(rc for rc in combo.rate_changes if rc.status == "taken")

    def f(r):
        cb = replace(combo, rate_changes=taken
                     + (RateChange(eff, r, "planned", False, 1.0),))
        return run_bridge(P, cb, "monthly").cy_lr_p - target

    a, b = lo, hi
    fa = f(a)
    for _ in range(200):
        mid = (a + b) / 2.0
        fm = f(mid)
        if fa * fm <= 0:
            b = mid
        else:
            a, fa = mid, fm
    return (a + b) / 2.0


@pytest.mark.parametrize("eff", [dt.date(P, 4, 1), dt.date(P, 6, 25)])
def test_solver_mixed_status_month_matches_brute_force(eff):
    """D58 oracle pin: PLANNED rows earlier in the same cohort month as TAKEN
    rows (one mixed month in P-1, one in P containing the solve date itself)
    must not disturb the taken-only base. The closed form must equal a
    brute-force bisection of the same D13 counterfactual through the full
    forward engine."""
    mixed = (
        RateChange(dt.date(P - 2, 7, 1), 0.04, "taken", True),
        RateChange(dt.date(P - 1, 6, 5), 0.03, "planned", False),  # steals rl_first
        RateChange(dt.date(P - 1, 6, 20), 0.02, "taken", False),   # must still blend
        RateChange(dt.date(P, 6, 5), 0.025, "planned", False),     # mixed month in P
        RateChange(dt.date(P, 6, 20), 0.015, "taken", False),
    )
    combo = _combo(mixed, net=None)
    target = 0.62
    closed = solve_rate_for_target(P, combo, target, eff).required_change
    brute = _brute_force_solver_r(combo, target, eff)
    assert closed == pytest.approx(brute, abs=1e-9)


def test_by_month_rows_are_coherent():
    combo = _combo(SHOWCASE)
    r, comp = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    rows = net_delivery_by_month(P, combo, dt.date(P, 4, 1), r)
    assert len(rows) == 12
    for row in rows:
        assert (1.0 + row["rate_leg"]) * (1.0 + row["price_leg_required"]) == \
            pytest.approx(1.10, abs=1e-12)
        assert row["m_required"] > 0
    # the October anniversary of the +10% taken filing: the rate leg collapses
    # and the pricing burden spikes (D39's compounding made visible)
    oct_row, mar_row = rows[9], rows[2]
    assert oct_row["rate_leg"] < mar_row["rate_leg"] - 0.05
    assert oct_row["price_leg_required"] > mar_row["price_leg_required"] + 0.05
