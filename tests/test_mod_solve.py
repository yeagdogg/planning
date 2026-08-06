"""Solving for a dated mod step (D73).

The mod leg inverts on the same algebra as the rate leg, so the standard is
the same one Mode A is held to: feed the solved answer back into the engine
and land on the target to machine precision, not approximately.

The exhibit these tests protect is the timing table — take mod in March and
you need +2.0%, wait until May and the same target costs +3.2%, and by
October there is not enough year left to earn it at any sane number.

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
    mod_timing_table,
    run_bridge,
    solve_mod_for_target,
)

P = 2027


def _combo(**kw) -> ComboInputs:
    mods = kw.pop("mods", None) or ModInputs(
        m_ind=0.850, m0=0.860, m0_asof=dt.date(P - 1, 9, 30), m1=0.890,
        m_end_prior=0.870)
    base = dict(
        lr_proj=0.65, mods=mods,
        rate_changes=(RateChange(dt.date(P - 1, 7, 1), 0.06, "taken", True),))
    base.update(kw)
    return ComboInputs(**base)


def _apply(combo, eff, c, status="taken"):
    return replace(combo, mod_changes=combo.mod_changes
                   + (RateChange(eff, c, status, False),))


# ---------------------------------------------------------------------------
# the inversion
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("target", [0.60, 0.62, 0.64, 0.66, 0.70])
@pytest.mark.parametrize("month", [1, 3, 5, 8])
def test_solved_step_hits_the_target_exactly(target, month):
    """Round-trip: solve, feed the answer back through the full bridge."""
    combo = _combo()
    eff = dt.date(P, month, 1)
    res = solve_mod_for_target(P, combo, target, eff)
    got = run_bridge(P, _apply(combo, eff, res.required_step), "monthly").cy_lr_p
    assert got == pytest.approx(target, abs=1e-12)


def test_inversion_holds_with_actions_already_logged():
    """The solve is incremental — it stacks on whatever is already in force."""
    combo = _apply(_combo(), dt.date(P, 2, 1), 0.015)
    eff = dt.date(P, 7, 1)
    res = solve_mod_for_target(P, combo, 0.635, eff)
    got = run_bridge(P, _apply(combo, eff, res.required_step), "monthly").cy_lr_p
    assert got == pytest.approx(0.635, abs=1e-12)


def test_planned_actions_are_excluded_from_the_base():
    """D13's convention, one dimension over: a PLANNED mod action is what you
    are solving for, so it cannot also be part of the position you solve from.
    Adding one must not change the answer."""
    combo = _combo()
    withp = _apply(combo, dt.date(P, 6, 1), 0.04, status="planned")
    eff = dt.date(P, 3, 1)
    a = solve_mod_for_target(P, combo, 0.64, eff).required_step
    b = solve_mod_for_target(P, withp, 0.64, eff).required_step
    assert a == pytest.approx(b, abs=1e-15)


def test_mid_month_date_inverts_too():
    """The D31 day-blend is inside the closed form, not an approximation."""
    combo = _combo()
    eff = dt.date(P, 4, 17)
    res = solve_mod_for_target(P, combo, 0.638, eff)
    got = run_bridge(P, _apply(combo, eff, res.required_step), "monthly").cy_lr_p
    assert got == pytest.approx(0.638, abs=1e-12)


def test_solve_inverts_even_when_the_combo_had_no_actions_at_all():
    """The re-anchor trap: with an empty log, the first action moves M_endPrior
    into play and reshapes the prior year. Linearising on a bare drift engine
    would be subtly wrong, so the base carries a zero-magnitude step."""
    combo = _combo()
    assert combo.mod_changes == ()
    eff = dt.date(P, 5, 1)
    res = solve_mod_for_target(P, combo, 0.63, eff)
    got = run_bridge(P, _apply(combo, eff, res.required_step), "monthly").cy_lr_p
    assert got == pytest.approx(0.63, abs=1e-12)


def test_achievement_grosses_the_directed_step_up():
    res = solve_mod_for_target(P, _combo(), 0.64, dt.date(P, 3, 1), achievement=0.70)
    assert res.directed_equivalent == pytest.approx(res.required_step / 0.70, abs=1e-15)
    # and directing that much at 70% achievement lands on the target
    got = run_bridge(P, _apply(_combo(), dt.date(P, 3, 1),
                               res.directed_equivalent), "monthly")
    # (the engine applies achievement itself only on planned rows; taken rows
    # take the number as given, so the directed figure overshoots by 1/0.70)
    assert got.cy_lr_p < 0.64


def test_planned_row_at_achievement_reproduces_the_directed_solve():
    """Directing c/a on a PLANNED row, which earns at a x directed, lands on
    the target — the round trip the Net Delivery gross-up relies on."""
    combo, eff, a = _combo(), dt.date(P, 3, 1), 0.70
    res = solve_mod_for_target(P, combo, 0.64, eff, achievement=a)
    planned = replace(combo, mod_changes=(
        RateChange(eff, res.directed_equivalent, "planned", False, a),))
    assert run_bridge(P, planned, "monthly").cy_lr_p == pytest.approx(0.64, abs=1e-12)


# ---------------------------------------------------------------------------
# the timing exhibit
# ---------------------------------------------------------------------------


def test_later_dates_cost_strictly_more():
    rows = mod_timing_table(P, _combo(), 0.640)
    steps = [r["required_step"] for r in rows]
    assert all(b > a for a, b in zip(steps, steps[1:])), steps
    # the headline comparison: March is materially cheaper than May
    assert rows[2]["required_step"] < rows[4]["required_step"]


def test_the_share_left_to_earn_falls_through_the_year():
    shares = [r["post_share"] for r in mod_timing_table(P, _combo(), 0.640)]
    assert all(b < a for a, b in zip(shares, shares[1:]))
    assert shares[0] == pytest.approx(0.5, abs=0.01)      # a 1/1 step earns ~half


def test_there_is_a_feasibility_cliff_late_in_the_year():
    rows = mod_timing_table(P, _combo(), 0.640)
    assert rows[0]["feasible"] and rows[2]["feasible"]
    assert not rows[11]["feasible"]                       # December cannot carry it
    first_bad = next(i for i, r in enumerate(rows) if not r["feasible"])
    assert all(not r["feasible"] for r in rows[first_bad:])   # no coming back


def test_every_row_of_the_timing_table_actually_inverts():
    combo, target = _combo(), 0.645
    for row in mod_timing_table(P, combo, target):
        c = row["required_step"]
        if c != c:                                        # nan
            continue
        got = run_bridge(P, _apply(combo, row["effective"], c), "monthly").cy_lr_p
        assert got == pytest.approx(target, abs=1e-12), row["month"]


# ---------------------------------------------------------------------------
# refusals
# ---------------------------------------------------------------------------


def test_mod_adjustment_off_is_called_out_and_infeasible():
    res = solve_mod_for_target(P, _combo(mod_adjustment_enabled=False), 0.64,
                               dt.date(P, 3, 1))
    assert not res.feasible
    assert any("OFF" in w for w in res.warnings)


def test_net_selection_is_called_out_and_infeasible():
    res = solve_mod_for_target(P, _combo(net_sel_p=0.08), 0.64, dt.date(P, 3, 1))
    assert not res.feasible
    assert any("supersedes" in w for w in res.warnings)


def test_absurd_year_end_mod_is_flagged():
    res = solve_mod_for_target(P, _combo(), 0.50, dt.date(P, 9, 1))
    assert any("filed range" in w or "reasonability" in w for w in res.warnings)


def test_direction_matches_intuition():
    """A LOWER target needs MORE mod (mod up = price up = lower LR)."""
    lo = solve_mod_for_target(P, _combo(), 0.62, dt.date(P, 3, 1)).required_step
    hi = solve_mod_for_target(P, _combo(), 0.66, dt.date(P, 3, 1)).required_step
    assert lo > 0 > hi
