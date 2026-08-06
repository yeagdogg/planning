"""Net delivery decomposition (D57): closed-form suggestion, affine rate leg,
dual M_1' solve, and the live-row filter semantics.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import datetime as dt
from dataclasses import replace

import pytest

from src import engine
from src.engine import (
    ComboInputs,
    ModInputs,
    MonthlyEngine,
    RateChange,
    _live_changes,
    month_index,
    net_delivery_by_month,
    net_delivery_by_month_p1,
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
        lr_proj=0.65, mods=mods,
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


def test_planned_change_pickup():
    """D63: Net Delivery's default filing — the FIRST planned row inside the
    plan year, at its EFFECTIVE percent (filed x achievement)."""
    assert engine.planned_change_in_plan_year(P, _combo(SHOWCASE)) == (
        dt.date(P, 4, 1), 0.05, 0.05, 1)
    # achievement is applied to the change in force, filed % reported separately
    ach = [RateChange(dt.date(P, 6, 1), 0.05, "planned", False, 0.8)]
    assert engine.planned_change_in_plan_year(P, _combo(ach)) == (
        dt.date(P, 6, 1), pytest.approx(0.04), 0.05, 1)
    # earliest wins, and the count reports what was NOT adopted
    two = [RateChange(dt.date(P, 9, 1), 0.03, "planned", False),
           RateChange(dt.date(P, 3, 1), 0.02, "planned", False)]
    got = engine.planned_change_in_plan_year(P, _combo(two))
    assert got[0] == dt.date(P, 3, 1) and got[3] == 2
    # taken rows and planned rows outside the plan year are never adopted
    assert engine.planned_change_in_plan_year(P, _combo([
        RateChange(dt.date(P, 5, 1), 0.05, "taken", False),
        RateChange(dt.date(P - 1, 7, 1), 0.06, "planned", False),
        RateChange(dt.date(P + 1, 2, 1), 0.07, "planned", False)])) is None
    assert engine.planned_change_in_plan_year(P, _combo()) is None


def test_planned_pickup_reproduces_delivery():
    """Adopting the planned filing is just (D, r) into the existing decomposition
    — the delivered path must equal a brute-force injection of that same row."""
    combo = _combo(SHOWCASE)
    d, r, _, _ = engine.planned_change_in_plan_year(P, combo)
    comp = net_delivery_components(P, combo, d)
    live = _live_changes(P, combo)
    brute = MonthlyEngine(P, combo,
                          changes_override=live + (RateChange(d, r, "taken", False),))
    for j in range(12):
        assert comp.a[j] + comp.b[j] * r == pytest.approx(
            brute.written_index(month_index(P, 1) + j), abs=1e-12)


def test_program_basis_plan_lr():
    """D65: the explicit rate x mod counterfactual. Identical to the headline
    when no net selection is in force; different, and net-independent, when
    one is."""
    plain = _combo(SHOWCASE, net=None)
    lp, lp1 = engine.program_basis_plan_lr(P, plain)
    base = run_bridge(P, plain, "monthly")
    assert lp == pytest.approx(base.cy_lr_p, abs=1e-12)
    assert lp1 == pytest.approx(base.cy_lr_p1, abs=1e-12)
    # asserting a target changes the headline but never the program basis
    for x in (0.05, 0.10, 0.25):
        netc = _combo(SHOWCASE, net=x)
        assert engine.program_basis_plan_lr(P, netc) == (lp, lp1)
        assert run_bridge(P, netc, "monthly").cy_lr_p != pytest.approx(lp, abs=1e-6)


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


# ---------------------------------------------------------------------------
# D102: the FOLLOWING year, carryover only. No second filing is assumed in P+1
# — nobody plans a rate change two years out — so the exhibit answers what THIS
# year's decision leaves for next year.
# ---------------------------------------------------------------------------


def test_p1_rows_close_exactly_to_the_p1_target():
    combo = _combo(SHOWCASE, net=0.10, net_sel_p1=0.04)
    r, _ = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    rows = net_delivery_by_month_p1(P, combo, dt.date(P, 4, 1), r)
    assert len(rows) == 12
    for row in rows:
        # the defining property: rate x pricing lands on the P+1 target, not P's
        assert (1.0 + row["rate_leg"]) * (1.0 + row["price_leg_required"]) == \
            pytest.approx(1.04, abs=1e-12)
        assert row["x"] == 0.04


def test_p1_target_falls_back_to_the_p_selection_when_blank():
    r, _ = suggest_net_rate(P, _combo(SHOWCASE), dt.date(P, 4, 1))
    blank = net_delivery_by_month_p1(P, _combo(SHOWCASE), dt.date(P, 4, 1), r)
    assert all(row["x"] == 0.10 for row in blank)
    for row in blank:
        assert (1.0 + row["rate_leg"]) * (1.0 + row["price_leg_required"]) == \
            pytest.approx(1.10, abs=1e-12)
    # and an explicit P+1 selection overrides it
    setp1 = net_delivery_by_month_p1(
        P, _combo(SHOWCASE, net_sel_p1=0.0), dt.date(P, 4, 1), r)
    assert all(row["x"] == 0.0 for row in setp1)


def test_p1_rate_leg_is_the_plan_year_filing_rolling_through():
    """With a clean log and one filing on 1 July P, the P+1 months before the
    anniversary still compare against a pre-filing base and carry the whole
    step; from July the filing sits on both sides and the leg goes to zero.
    That IS the carryover the exhibit exists to show."""
    combo = _combo(net=0.10)                       # no rate changes at all
    r = 0.08
    rows = net_delivery_by_month_p1(P, combo, dt.date(P, 7, 1), r)
    for row in rows[:6]:                           # Jan..Jun P+1
        assert row["rate_leg"] == pytest.approx(r, abs=1e-12)
    for row in rows[6:]:                           # Jul..Dec P+1
        assert row["rate_leg"] == pytest.approx(0.0, abs=1e-12)
    # so the pricing burden is entirely in the back half of the year
    assert rows[0]["price_leg_required"] < rows[11]["price_leg_required"]


def test_p1_reports_no_pricing_leg_when_the_mod_adjustment_is_off():
    combo = _combo(SHOWCASE, mod_adjustment_enabled=False)
    r, _ = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    rows = net_delivery_by_month_p1(P, combo, dt.date(P, 4, 1), r)
    assert all(row["price_leg_required"] is None for row in rows)
    assert all(row["m_required"] is None for row in rows)


def test_p1_mod_base_is_the_plan_year_written_mod():
    """m_required is stated against the year-ago mod, which for a P+1 month is
    the plan-year written mod — the same mproj the P decomposition carries."""
    combo = _combo(SHOWCASE)
    r, comp = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    rows = net_delivery_by_month_p1(P, combo, dt.date(P, 4, 1), r)
    for j, row in enumerate(rows):
        assert row["m_base"] == pytest.approx(comp.mproj[j], abs=1e-15)
        assert row["m_required"] == pytest.approx(
            comp.mproj[j] * (1.0 + row["x"]) / (1.0 + row["rate_leg"]), abs=1e-15)


# ---------------------------------------------------------------------------
# The mod ask as a dated ACTION (D75) — required_mod_step
#
# required_m1 answers with a LEVEL, which is the right shape only while the
# mod path is a drift line. These tests pin the action form: apply the solved
# step to the log, and the combo delivers its own target exactly.
# ---------------------------------------------------------------------------


def _with_step(combo, d_m, c):
    """``combo`` with one more mod step of ``c`` on ``d_m``."""
    return replace(combo, mod_changes=engine.base_plus(combo.mod_changes, d_m, c))


def _delivered_with_step(combo, eff, r, d_m, c):
    """Round-trip: written-weighted delivered net once the step is LOGGED."""
    return _delivered_avg(net_delivery_components(P, _with_step(combo, d_m, c), eff), r)


MOD_LOG = (
    RateChange(dt.date(P - 1, 7, 1), -0.02, "taken", False),   # history (drift owns)
    RateChange(dt.date(P, 2, 1), -0.015, "taken", False),      # plan year
    RateChange(dt.date(P, 6, 10), -0.01, "planned", False),    # read WHOLE (D57)
)


@pytest.mark.parametrize("d_m", [
    dt.date(P, 1, 1), dt.date(P, 3, 1), dt.date(P, 6, 10),   # co-occupies with a step
    dt.date(P, 6, 20), dt.date(P, 7, 15), dt.date(P, 11, 1),
])
@pytest.mark.parametrize("log", [(), MOD_LOG])
def test_solved_step_reproduces_the_target(d_m, log):
    combo = _combo(SHOWCASE, mod_changes=log)
    r, _ = suggest_net_rate(P, combo, dt.date(P, 4, 1))
    res = engine.required_mod_step(P, combo, dt.date(P, 4, 1), r, mod_eff=d_m)
    assert _delivered_with_step(combo, dt.date(P, 4, 1), r, d_m, res.required_step) == \
        pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


@pytest.mark.parametrize("r", [-0.05, 0.0, 0.03, 0.11])
def test_solved_step_reproduces_the_target_at_any_filing(r):
    """The step closes whatever gap the filing leaves — including a filing
    that already overshoots, where the required step turns negative."""
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    res = engine.required_mod_step(P, combo, dt.date(P, 4, 1), r, mod_eff=dt.date(P, 4, 1))
    assert _delivered_with_step(combo, dt.date(P, 4, 1), r, dt.date(P, 4, 1),
                                res.required_step) == \
        pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


def test_delivery_is_exactly_affine_in_one_plus_c():
    """The closed form is a linearisation only if delivery really is affine in
    (1 + c). Two probes fix the line; a third has to land on it."""
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff, d_m, r = dt.date(P, 4, 1), dt.date(P, 5, 1), 0.06
    y0 = _delivered_with_step(combo, eff, r, d_m, 0.0)
    y1 = _delivered_with_step(combo, eff, r, d_m, 0.10)
    for c in (-0.07, 0.035, 0.25):
        assert _delivered_with_step(combo, eff, r, d_m, c) == \
            pytest.approx(y0 + (y1 - y0) * c / 0.10, abs=1e-12)


def test_step_defaults_to_the_filings_own_date():
    """One filing carries both levers unless the caller says otherwise."""
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff = dt.date(P, 5, 15)
    r, _ = suggest_net_rate(P, combo, eff)
    assert engine.required_mod_step(P, combo, eff, r).required_step == \
        engine.required_mod_step(P, combo, eff, r, mod_eff=eff).required_step


def test_once_the_step_is_logged_nothing_more_is_required():
    """The fixed point: solve, log the answer, and a later solve asks for 0."""
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff = dt.date(P, 4, 1)
    r, _ = suggest_net_rate(P, combo, eff)
    first = engine.required_mod_step(P, combo, eff, r, mod_eff=dt.date(P, 4, 1))
    done = _with_step(combo, dt.date(P, 4, 1), first.required_step)
    for later in (dt.date(P, 4, 1), dt.date(P, 8, 1), dt.date(P, 11, 1)):
        again = engine.required_mod_step(P, done, eff, r, mod_eff=later)
        assert again.required_step == pytest.approx(0.0, abs=1e-12)


def test_achievement_grosses_up_what_you_direct():
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff = dt.date(P, 4, 1)
    res = engine.required_mod_step(P, combo, eff, 0.02, achievement=0.8)
    assert res.directed_equivalent == pytest.approx(res.required_step / 0.8, abs=1e-15)
    # and directing that much at 80% achievement is what actually delivers
    assert _delivered_with_step(combo, eff, 0.02, eff,
                                res.directed_equivalent * 0.8) == \
        pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


def test_later_dates_cost_more_and_carry_less():
    """Slipping the date leaves less of the year to earn the step, so the ask
    grows and the share it can move shrinks — the feasibility cliff."""
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff = dt.date(P, 4, 1)
    rows = [engine.required_mod_step(P, combo, eff, 0.0, mod_eff=dt.date(P, m, 1))
            for m in range(1, 13)]
    assert all(a.required_step < b.required_step - 1e-9
               for a, b in zip(rows, rows[1:]))          # r = 0 leaves a rate gap
    assert all(a.post_share > b.post_share + 1e-9 for a, b in zip(rows, rows[1:]))
    assert rows[0].feasible and not rows[-1].feasible


def test_m_end_is_the_year_end_mod_the_step_produces():
    combo = _combo(SHOWCASE, mod_changes=MOD_LOG)
    eff = dt.date(P, 4, 1)
    r, _ = suggest_net_rate(P, combo, eff)
    res = engine.required_mod_step(P, combo, eff, r, mod_eff=dt.date(P, 5, 1))
    done = MonthlyEngine(P, _with_step(combo, dt.date(P, 5, 1), res.required_step))
    assert res.m_end == pytest.approx(done.derived_m1(), abs=1e-12)
    assert res.mod_base == pytest.approx(done.mod_base, abs=1e-12)


def test_explicit_m_end_prior_is_the_level_the_step_compounds_on():
    """D70's re-anchor is in force even with an empty log, because ASKING for
    a step commits the combo to the stepped regime."""
    mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(P - 1, 9, 30), m1=0.89,
                     m_end_prior=0.875)
    combo = _combo(SHOWCASE, mods=mods)
    res = engine.required_mod_step(P, combo, dt.date(P, 4, 1), 0.05)
    assert res.mod_base == pytest.approx(0.875, abs=1e-15)
    assert _delivered_with_step(combo, dt.date(P, 4, 1), 0.05, dt.date(P, 4, 1),
                                res.required_step) == \
        pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)


def test_out_of_bounds_year_end_mod_is_flagged_not_hidden():
    """A target that needs an unfileable mod still returns the number — with
    the reason it cannot be filed attached."""
    # the two bounds are independent: +61.9% is an unreasonable STEP that still
    # lands on a fileable mod (1.340); +111.4% lands at 1.750, which is neither
    steep = engine.required_mod_step(
        P, _combo(SHOWCASE, net=0.60, mod_changes=MOD_LOG), dt.date(P, 4, 1), 0.02)
    assert steep.required_step > 0.15 and not steep.feasible
    assert any("reasonability" in w for w in steep.warnings)
    assert not any("filed range" in w for w in steep.warnings)
    wild = engine.required_mod_step(
        P, _combo(SHOWCASE, net=1.00, mod_changes=MOD_LOG), dt.date(P, 4, 1), 0.02)
    assert wild.m_end > 1.5 and not wild.feasible
    assert any("filed range" in w for w in wild.warnings)


def test_refuses_when_the_mod_adjustment_is_off():
    combo = _combo(SHOWCASE, mod_adjustment_enabled=False)
    with pytest.raises(ValueError, match="mod adjustment is off"):
        engine.required_mod_step(P, combo, dt.date(P, 4, 1), 0.05)


def test_refuses_a_step_dated_outside_the_plan_year():
    combo = _combo(SHOWCASE)
    with pytest.raises(ValueError, match="must lie inside plan year"):
        engine.required_mod_step(P, combo, dt.date(P, 4, 1), 0.05,
                                 mod_eff=dt.date(P - 1, 12, 1))


def test_refuses_without_a_net_selection():
    combo = _combo(SHOWCASE, net=None)
    with pytest.raises(ValueError, match="requires a net rate selection"):
        engine.required_mod_step(P, combo, dt.date(P, 4, 1), 0.05)


def test_mod_step_survives_a_degenerate_as_of_date():
    """"Mods as of 12/31 of the prior year" leaves no history segment for the
    D70 switch anchor to sit on. The oracle drops it; the solve must still
    round-trip, because the workbook builds that same path from anchor cells
    and a zero-width segment there is a divide-by-zero."""
    for asof in (dt.date(P - 1, 12, 31), dt.date(P, 6, 30)):
        for m_prior in (None, 0.83):
            mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=asof, m1=0.89,
                             m_prior=m_prior, m_end_prior=0.875)
            combo = _combo(SHOWCASE, mods=mods, mod_changes=MOD_LOG)
            eff = dt.date(P, 4, 1)
            r, _ = suggest_net_rate(P, combo, eff)
            res = engine.required_mod_step(P, combo, eff, r)
            assert res.mod_base == pytest.approx(0.875, abs=1e-15)
            assert _delivered_with_step(combo, eff, r, eff, res.required_step) == \
                pytest.approx(1.0 + combo.net_sel_p, abs=1e-12)
