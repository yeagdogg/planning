"""Stepped schedule-mod actions (D70).

A mod change is a dated percent that stays in force until superseded — the
same mechanic as a rate change, applied to the schedule mod. Drift owns
history through 12/31/(P-1); the step log owns the plan year onward.

The most important test in this file is
``test_january_step_reproduces_the_endpoint_average``: it pins the fact that
the shop's current method (average the two year-end mods) is EXACTLY right for
a single step on 1/1/P at a 12-month term, which is what makes the two
processes reconcilable and localises the whole difference to effective date
and term.

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
    add_months,
    mi_mid_ordinal,
    month_index,
    program_flow_by_month,
    run_bridge,
    validate_inputs,
)

P = 2027
M_END_PRIOR = 0.850


def _mods(**kw) -> ModInputs:
    base = dict(m_ind=0.850, m0=0.850, m0_asof=dt.date(P - 1, 9, 30), m1=0.850,
                m_end_prior=M_END_PRIOR)
    base.update(kw)
    return ModInputs(**base)


def _combo(mod_changes=(), mods=None, **kw) -> ComboInputs:
    return ComboInputs(lr_proj=0.65, lr_basis="current", mods=mods or _mods(),
                       mod_changes=tuple(mod_changes), **kw)


def _step(month, day, pct, status="taken", ach=None):
    return RateChange(dt.date(P, month, day), pct, status, False,
                      *( (ach,) if ach is not None else () ))


# ---------------------------------------------------------------------------
# the identity that makes the old process reconcilable
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("c", [-0.05, -0.02, 0.03, 0.06])
def test_january_step_reproduces_the_endpoint_average(c):
    """12-month term: a single step on 1/1/P earns exactly half into CY P, so
    the earned mod equals the average of the two year-end mods — which is the
    current process's arithmetic, exactly."""
    eng = MonthlyEngine(P, _combo([_step(1, 1, c)]))
    theirs = (M_END_PRIOR + M_END_PRIOR * (1.0 + c)) / 2.0
    assert eng.earned_mod_cy(P) == pytest.approx(theirs, abs=1e-12)


@pytest.mark.parametrize("term,share", [(12, 0.5), (6, 0.75), (3, 0.875)])
def test_earned_share_of_a_january_step_by_term(term, share):
    """Share of CY P that a 1/1 step actually earns. The endpoint-average
    assumes 50% always — true only at a 12-month term. Shorter terms turn the
    book over faster, so a plan-year action earns in MORE, and the current
    process under-credits it (6-month: 75% real vs 50% assumed)."""
    c = 0.10
    eng = MonthlyEngine(P, _combo([_step(1, 1, c)], term_months=term))
    got = (eng.earned_mod_cy(P) / M_END_PRIOR - 1.0) / c
    assert got == pytest.approx(share, abs=1e-12)


def test_later_steps_earn_strictly_less_into_the_plan_year():
    prev = None
    for m in range(1, 13):
        eng = MonthlyEngine(P, _combo([_step(m, 1, 0.04)]))
        earned = eng.earned_mod_cy(P)
        if prev is not None:
            assert earned < prev
        prev = earned
    # a December step is nearly worthless to the plan year
    assert prev == pytest.approx(M_END_PRIOR, abs=2e-3)


# ---------------------------------------------------------------------------
# the step mechanic itself
# ---------------------------------------------------------------------------


def test_written_mod_steps_at_the_date_and_holds():
    eng = MonthlyEngine(P, _combo([_step(4, 1, -0.03)]))
    for m in (1, 2, 3):
        assert eng.written_mod(month_index(P, m)) == pytest.approx(M_END_PRIOR, abs=1e-12)
    for m in (4, 7, 12):
        assert eng.written_mod(month_index(P, m)) == pytest.approx(
            M_END_PRIOR * 0.97, abs=1e-12)
    # ...and stays in force through P+1 with no further action
    assert eng.written_mod(month_index(P + 1, 6)) == pytest.approx(
        M_END_PRIOR * 0.97, abs=1e-12)


def test_mid_month_step_day_blends_like_a_rate_change():
    """A 4/15 mod step blends its own month at p = 16/30 — the same D31 rule
    the rate leg uses, because both call blended_step_index."""
    eng = MonthlyEngine(P, _combo([_step(4, 15, 0.04)]))
    p = 16 / 30
    expect = M_END_PRIOR * ((1 - p) + p * 1.04)
    assert eng.written_mod(month_index(P, 4)) == pytest.approx(expect, abs=1e-12)
    assert eng.written_mod(month_index(P, 5)) == pytest.approx(
        M_END_PRIOR * 1.04, abs=1e-12)


def test_steps_compound():
    eng = MonthlyEngine(P, _combo([_step(3, 1, 0.02), _step(9, 1, 0.03)]))
    assert eng.written_mod(month_index(P, 6)) == pytest.approx(
        M_END_PRIOR * 1.02, abs=1e-12)
    assert eng.written_mod(month_index(P, 10)) == pytest.approx(
        M_END_PRIOR * 1.02 * 1.03, abs=1e-12)
    assert eng.derived_m1() == pytest.approx(M_END_PRIOR * 1.02 * 1.03, abs=1e-12)


def test_planned_mod_at_achievement_equals_a_smaller_taken_step():
    """Mod actions are usually not fully achieved — the achievement field
    means the same thing it does on a rate row."""
    part = MonthlyEngine(P, _combo([_step(4, 1, 0.05, "planned", 0.7)]))
    full = MonthlyEngine(P, _combo([_step(4, 1, 0.035, "taken")]))
    assert [part.written_mod(k) for k in part.cohorts] == pytest.approx(
        [full.written_mod(k) for k in full.cohorts], abs=1e-15)


def test_path_is_continuous_across_the_new_year():
    """Drift hands off to the step log at 12/31/(P-1) with no jump: with no
    step in January the first plan-year cohort sits at m_end_prior."""
    eng = MonthlyEngine(P, _combo([_step(7, 1, 0.05)]))
    assert eng.written_mod(month_index(P, 1)) == pytest.approx(M_END_PRIOR, abs=1e-12)
    assert eng.written_mod(month_index(P - 1, 12)) == pytest.approx(
        M_END_PRIOR, abs=2e-3)


def test_history_still_drifts_before_the_plan_year():
    """Steps must not flatten the observed run-up to the current year end."""
    mods = _mods(m0=0.880, m_end_prior=0.850)
    eng = MonthlyEngine(P, _combo([_step(4, 1, 0.02)], mods=mods))
    sep = eng.written_mod(month_index(P - 1, 9))
    dec = eng.written_mod(month_index(P - 1, 12))
    assert 0.850 < dec < sep < 0.885          # still sliding through the old year


# ---------------------------------------------------------------------------
# invariants that must survive the change
# ---------------------------------------------------------------------------


def test_no_mod_rows_is_bit_identical_to_the_drift_build():
    """The whole de-risking argument: an empty log changes nothing, and
    m_end_prior is inert without steps."""
    drift = ComboInputs(lr_proj=0.65, lr_basis="current",
                        mods=ModInputs(m_ind=0.85, m0=0.86,
                                       m0_asof=dt.date(P - 1, 9, 30), m1=0.90),
                        rate_changes=(RateChange(dt.date(P - 1, 7, 1), 0.06, "taken", True),))
    with_field = replace(drift, mods=replace(drift.mods, m_end_prior=0.83))
    a, b = run_bridge(P, drift, "monthly"), run_bridge(P, with_field, "monthly")
    assert a.cy_lr_p == b.cy_lr_p and a.cy_lr_p1 == b.cy_lr_p1
    assert a.a_mod_p == b.a_mod_p and a.e_cy == b.e_cy


def test_under_a_net_selection_the_step_is_inert_but_the_re_anchor_is_not():
    """D39 x D70, and the one interaction that surprises people.

    A plan-year mod action cannot reach a NET-mode plan LR: the net path takes
    over from 1/1/P, so A_mod is 1 and the mod leg is superseded. What DOES
    move the number is the RE-ANCHOR — logging the first action makes
    M_endPrior live, and the drift path now lands on it at 12/31/(P-1) instead
    of running on to M_1, which reshapes the prior-year mod the net path is
    seeded from. Put M_endPrior back on the old drift line and the plan LR
    returns exactly, which is what localises the whole effect to the anchor.
    """
    mods = _mods(m0=0.860, m1=0.900, m_end_prior=0.880)
    net = _combo(mods=mods, net_sel_p=0.08)
    x1 = mods.m0_asof.toordinal() + 1
    x2 = dt.date(P, 12, 31).toordinal() + 1
    on_line = mods.m0 + (mods.m1 - mods.m0) / (x2 - x1) * (
        dt.date(P - 1, 12, 31).toordinal() + 1 - x1)

    stepped = _combo([_step(5, 1, -0.03)], mods=replace(mods, m_end_prior=on_line),
                     net_sel_p=0.08)
    a, b = run_bridge(P, net, "monthly"), run_bridge(P, stepped, "monthly")
    assert a.a_mod_p == 1.0 and b.a_mod_p == 1.0        # net supersedes the mod leg
    assert b.cy_lr_p == pytest.approx(a.cy_lr_p, abs=1e-12)

    # ...and the test is not vacuous: the step really does move the earned mod,
    # it just cannot reach the plan LR while the net selection is in force
    assert MonthlyEngine(P, stepped).earned_mod_cy(P) != pytest.approx(
        MonthlyEngine(P, net).earned_mod_cy(P), abs=1e-6)

    # off net mode the same action lands on the plan LR
    off = replace(stepped, net_sel_p=None)
    assert run_bridge(P, off, "monthly").cy_lr_p != pytest.approx(
        run_bridge(P, replace(net, net_sel_p=None), "monthly").cy_lr_p, abs=1e-6)


def test_monthly_and_continuous_agree_on_a_stepped_path():
    combo = _combo([_step(4, 1, 0.03), _step(10, 1, 0.02)],
                   rate_changes=(RateChange(dt.date(P - 1, 7, 1), 0.05, "taken", True),))
    mo = run_bridge(P, combo, "monthly")
    co = run_bridge(P, combo, "continuous")
    assert mo.cy_lr_p == pytest.approx(co.cy_lr_p, abs=0.001)   # within 0.10 pts


def test_mod_leg_of_program_flow_stays_term_independent():
    """Written-basis YoY has no term in it — true for the mod leg too."""
    changes = [_step(4, 1, 0.03)]
    r6 = program_flow_by_month(P, _combo(changes, term_months=6))
    r12 = program_flow_by_month(P, _combo(changes, term_months=12))
    assert [r["mod_leg"] for r in r6.rows] == pytest.approx(
        [r["mod_leg"] for r in r12.rows], abs=1e-15)


def test_a_mod_moves_the_plan_lr_in_the_right_direction():
    """Mod UP = price up = lower plan LR."""
    up = run_bridge(P, _combo([_step(1, 1, 0.04)]), "monthly")
    flat = run_bridge(P, _combo(), "monthly")
    down = run_bridge(P, _combo([_step(1, 1, -0.04)]), "monthly")
    assert up.cy_lr_p < flat.cy_lr_p < down.cy_lr_p


# ---------------------------------------------------------------------------
# validation
# ---------------------------------------------------------------------------


def test_mod_row_before_the_plan_year_is_refused():
    bad = _combo([RateChange(dt.date(P - 1, 6, 1), 0.03, "taken", False)])
    with pytest.raises(ValueError) as e:
        validate_inputs(P, bad)
    assert "twice" in str(e.value)


def test_missing_year_end_estimate_is_advised_not_fatal():
    combo = _combo([_step(4, 1, 0.02)], mods=_mods(m_end_prior=None))
    w = validate_inputs(P, combo)
    assert any("year-end mod" in s for s in w)
    assert MonthlyEngine(P, combo).mod_base is not None   # still usable


def test_mod_steps_with_the_adjustment_off_are_advised_inert():
    combo = _combo([_step(4, 1, 0.02)], mod_adjustment_enabled=False)
    w = validate_inputs(P, combo)
    assert any("deliver nothing" in s for s in w)
    assert run_bridge(P, combo, "monthly").a_mod_p == pytest.approx(1.0, abs=1e-15)


def test_two_mod_steps_in_one_month_are_advised():
    combo = _combo([_step(4, 5, 0.02), _step(4, 20, 0.01)])
    assert any("share cohort month" in s for s in validate_inputs(P, combo))


# ---------------------------------------------------------------------------
# The degenerate as-of date: "mods as of 12/31 of the prior year"
#
# D70 moves the drift path's forward anchor back to 12/31/(P-1). An as-of date
# on or after that leaves no history segment at all — the engine drops the
# switch anchor rather than building a zero-width one. Ordinary thing to type,
# so it is pinned rather than validated away.
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("asof", [dt.date(P - 1, 12, 31), dt.date(P, 6, 30)])
def test_asof_at_or_after_the_switch_leaves_a_flat_history(asof):
    """M_prior blank: nothing anchors the history but M_0 itself."""
    mods = _mods(m0=0.900, m0_asof=asof, m_end_prior=0.870)
    eng = MonthlyEngine(P, _combo([RateChange(dt.date(P, 3, 1), 0.02, "taken", False)],
                                  mods=mods))
    for k in range(month_index(P - 2, 1), month_index(P, 1)):
        assert eng.written_mod(k) == pytest.approx(0.900, abs=1e-12)


@pytest.mark.parametrize("asof", [dt.date(P - 1, 12, 31), dt.date(P, 6, 30)])
def test_asof_at_or_after_the_switch_keeps_the_prior_line(asof):
    """M_prior present: the M_prior -> M_0 line, extrapolated, and nothing
    snaps to m_end_prior — there is no segment left for it to anchor."""
    mods = _mods(m0=0.900, m_prior=0.860, m0_asof=asof, m_end_prior=0.870)
    eng = MonthlyEngine(P, _combo([RateChange(dt.date(P, 3, 1), 0.02, "taken", False)],
                                  mods=mods))
    x0, x1 = float(add_months(asof, -12).toordinal() + 1), float(asof.toordinal() + 1)
    slope = (0.900 - 0.860) / (x1 - x0)
    for k in range(month_index(P - 2, 1), month_index(P, 1)):
        want = 0.860 + slope * (mi_mid_ordinal(k) - x0)
        assert eng.written_mod(k) == pytest.approx(want, abs=1e-12)


def test_degenerate_asof_still_steps_on_m_end_prior():
    """The plan year is unaffected: the log still compounds on M_endPrior."""
    mods = _mods(m0=0.900, m0_asof=dt.date(P - 1, 12, 31), m_end_prior=0.870)
    eng = MonthlyEngine(P, _combo([RateChange(dt.date(P, 1, 1), 0.02, "taken", False)],
                                  mods=mods))
    assert eng.mod_base == pytest.approx(0.870, abs=1e-15)
    assert eng.written_mod(month_index(P, 6)) == pytest.approx(0.870 * 1.02, abs=1e-12)


def test_blank_m_end_prior_extends_the_drift_it_does_not_stop_at_m0():
    """The contract the workbook has to match: a blank M_endPrior means the
    planner has not RESTATED the level, not that the mod stops drifting. The
    engine reads the unswitched anchor path at 12/31/(P-1); M_0 would silently
    discard the drift already entered."""
    mods = _mods(m0=0.840, m0_asof=dt.date(P - 1, 9, 30), m1=0.845, m_end_prior=None)
    eng = MonthlyEngine(P, _combo([RateChange(dt.date(P, 4, 1), 0.02, "taken", False)],
                                  mods=mods))
    want = eng.mod_path.value(eng.x_switch)
    assert eng.mod_base == pytest.approx(want, abs=1e-15)
    assert eng.mod_base > 0.840 + 1e-5          # strictly past M_0 — drift is real
    assert eng.mod_base < 0.845                 # but has not reached M_1 either


def test_blank_m_end_prior_with_m_prior_and_a_degenerate_as_of():
    """Both edge cases at once: no restated level AND no history segment. The
    fallback lands on the M_prior -> M_0 line, which is the only line left."""
    mods = _mods(m0=0.900, m_prior=0.860, m0_asof=dt.date(P, 6, 30), m_end_prior=None)
    eng = MonthlyEngine(P, _combo([RateChange(dt.date(P, 7, 1), 0.02, "taken", False)],
                                  mods=mods))
    x0 = float(add_months(mods.m0_asof, -12).toordinal() + 1)
    x1 = float(mods.m0_asof.toordinal() + 1)
    xs = float(dt.date(P - 1, 12, 31).toordinal() + 1)
    want = 0.860 + (0.900 - 0.860) / (x1 - x0) * (xs - x0)
    assert eng.mod_base == pytest.approx(want, abs=1e-12)
