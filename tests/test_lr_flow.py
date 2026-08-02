"""LR flow: the monthly plan loss ratio, and the race between trend and earn-in.

The exhibit's whole claim is that a monthly walk whose weighted mean sits a
known, decomposable distance from the CY headline is more useful than one forced
to tie exactly. These tests pin both halves of that: the anchor that makes the
walk average to the headline, and the identity that accounts for what is left.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import dataclasses
import math

import pytest

from src import engine as E


@pytest.fixture(scope="module")
def base():
    p, combo = E.worked_example()
    return p, dataclasses.replace(combo, net_trend=0.05)


def _variants(p, combo):
    """The four regimes the monthly walk has to be right in."""
    seas_q4 = (0.0,) * 9 + (6.0, 3.0, 3.0)
    return [
        ("mod on", combo),
        ("trend off", dataclasses.replace(combo, net_trend=0.0)),
        ("net selection", dataclasses.replace(combo, net_sel_p=0.05)),
        ("mod adjustment off",
         dataclasses.replace(combo, mod_adjustment_enabled=False)),
        ("short term, lumpy writings",
         dataclasses.replace(combo, term_months=6, seasonality=seas_q4)),
    ]


# ---------------------------------------------------------------------------
# the anchor
# ---------------------------------------------------------------------------


def test_anchor_makes_both_years_average_to_their_headline(base):
    """7/1/P is the only anchor where BOTH years work.

    The plan-year headline carries no trend, so the plan-year exponents must
    average to 0. The P+1 headline carries exactly one trend step, so the P+1
    exponents must average to 1. Half a month either way and one of them breaks
    — month CENTRES average to 6.0 and month NUMBERS to 6.5, and only the
    second is the 7/1 boundary.
    """
    p, combo = base
    f = E.lr_flow_by_month(p, combo)
    d = [r["delta"] for r in f.rows]
    assert sum(d[:12]) / 12 == pytest.approx(0.0, abs=1e-15)
    assert sum(d[12:]) / 12 == pytest.approx(1.0, abs=1e-15)
    # and the anchor really is the 7/1 boundary, not July's midpoint
    assert E.LR_FLOW_ANCHOR == 6.5


def test_a_half_month_slip_would_cost_real_basis_points(base):
    """Why the anchor gets a test rather than a comment."""
    p, combo = base
    t = 1.0 + combo.net_trend
    # the same walk anchored on July's CENTRE instead of the 7/1 boundary
    off_p = sum(t ** ((mm - 6.0) / 12.0) for mm in range(1, 13)) / 12
    off_p1 = sum(t ** ((mm - 6.0) / 12.0) for mm in range(13, 25)) / 12
    assert abs(off_p - 1.0) * 1e4 > 15          # ~19 bp of factor in P
    assert abs(off_p1 / t - 1.0) * 1e4 > 15     # and again in P+1


# ---------------------------------------------------------------------------
# the residual, and that it is accounted for rather than hidden
# ---------------------------------------------------------------------------


@pytest.mark.parametrize("tag", [v[0] for v in _variants(*E.worked_example())])
def test_weighted_mean_reconstructs_from_the_decomposition(base, tag):
    """mean = head x E_den[T^delta] / (1 + c_EM), exactly, in every regime.

    This is the claim the exhibit makes when it shows a residual instead of
    forcing a tie. If the identity ever stops holding, the residual has stopped
    being an explanation and become an error.
    """
    p, combo = base
    c = dict(_variants(p, combo))[tag]
    f = E.lr_flow_by_month(p, c)
    rebuilt = f.head_p * f.trend_mean_p / (1.0 + f.cov_rate_price)
    assert f.mean_p == pytest.approx(rebuilt, rel=1e-12)


def test_trend_off_isolates_the_rate_price_covariance(base):
    """With no trend the only thing left is the covariance, so the residual
    must BE it — the headline divides by two separate averages, the walk
    divides month by month by their product."""
    p, combo = base
    f = E.lr_flow_by_month(p, dataclasses.replace(combo, net_trend=0.0))
    assert f.convexity_p == pytest.approx(0.0, abs=1e-15)
    assert f.resid_p == pytest.approx(-f.cov_rate_price / (1 + f.cov_rate_price),
                                      rel=1e-9)


def test_covariance_vanishes_when_there_is_no_separate_price_leg(base):
    """Under a net selection, or with the mod adjustment off, there is only one
    leg — so there is nothing for the rate to covary with."""
    p, combo = base
    for c in (dataclasses.replace(combo, net_sel_p=0.05),
              dataclasses.replace(combo, mod_adjustment_enabled=False)):
        f = E.lr_flow_by_month(p, c)
        assert not f.mod_on
        assert f.cov_rate_price == pytest.approx(0.0, abs=1e-15)


def test_annual_term_has_no_centre_of_gravity_tilt(base):
    """Earned exposure per month is uniform for an annual term REGARDLESS of
    written seasonality, so the fixed 7/1 anchor is the exposure centre of
    gravity and the tilt term is structurally zero. A short term breaks that,
    and the tilt is how the exhibit says so."""
    p, combo = base
    seas = (0.4, 0.4, 0.4, 0.6, 0.6, 0.8, 0.8, 1.0, 1.4, 2.2, 1.8, 1.6)
    for c in (combo, dataclasses.replace(combo, seasonality=seas)):
        assert E.lr_flow_by_month(p, c).cog_tilt_p == pytest.approx(0.0, abs=1e-12)
    lumpy = dataclasses.replace(combo, term_months=6, seasonality=seas)
    assert abs(E.lr_flow_by_month(p, lumpy).cog_tilt_p) > 1e-4


def test_the_residual_stays_small_where_the_model_is_well_posed(base):
    """A few basis points on an annual term — small enough to disclose rather
    than solve, which is the argument for disclosing it."""
    p, combo = base
    f = E.lr_flow_by_month(p, combo)
    assert abs(f.resid_p) < 5e-4
    assert abs(f.resid_p1) < 5e-4


# ---------------------------------------------------------------------------
# the race
# ---------------------------------------------------------------------------


def test_breakeven_trend_actually_flattens_the_year(base):
    """t* is the net trend at which trend exactly offsets rate and mod earn-in,
    so January and December land on the same loss ratio. Solved in closed form
    from the two months' factors; verified by re-running the walk at it."""
    p, combo = base
    t_star = E.lr_flow_by_month(p, combo).breakeven_trend
    assert t_star is not None
    flat = E.lr_flow_by_month(p, dataclasses.replace(combo, net_trend=t_star))
    assert flat.rows[0]["lr"] == pytest.approx(flat.rows[11]["lr"], rel=1e-12)


def test_below_breakeven_the_year_falls_and_above_it_rises(base):
    p, combo = base
    t_star = E.lr_flow_by_month(p, combo).breakeven_trend
    lo = E.lr_flow_by_month(p, dataclasses.replace(combo, net_trend=t_star - 0.03))
    hi = E.lr_flow_by_month(p, dataclasses.replace(combo, net_trend=t_star + 0.03))
    assert lo.rows[11]["lr"] < lo.rows[0]["lr"]      # earn-in wins
    assert hi.rows[11]["lr"] > hi.rows[0]["lr"]      # trend wins


def test_a_net_selection_breaks_even_at_its_own_target(base):
    """Under a net selection the combined price renews at the target, so the
    year is flat exactly when trend equals it. A good independent check that
    the mod leg is not being double-counted on top of the net path."""
    p, combo = base
    f = E.lr_flow_by_month(p, dataclasses.replace(combo, net_sel_p=0.05))
    assert f.breakeven_trend == pytest.approx(0.05, abs=1e-3)


# ---------------------------------------------------------------------------
# months that earn nothing
# ---------------------------------------------------------------------------


def test_months_that_earn_nothing_are_gaps_not_zeros(base):
    """A six-month term written only in Q4 genuinely earns nothing in parts of
    the year. A zero there would read as a 0% loss ratio and drag every
    average; the walk reports None and names the months."""
    p, combo = base
    lumpy = dataclasses.replace(combo, term_months=6,
                                seasonality=(0.0,) * 9 + (6.0, 3.0, 3.0))
    f = E.lr_flow_by_month(p, lumpy)
    assert f.dead_months, "expected some months to earn nothing"
    for r in f.rows:
        if r["month"] in f.dead_months:
            assert r["lr"] is None and r["weight"] == 0.0
        else:
            assert r["lr"] is not None and r["lr"] > 0.0
    assert not math.isnan(f.mean_p)          # the mean skips them, not chokes


def test_every_row_is_the_bridge_evaluated_at_that_month(base):
    """No second engine: each month is LR_current x trend x A_rate x A_mod x
    A_other, the same product the headline assembles."""
    p, combo = base
    f = E.lr_flow_by_month(p, combo)
    res = E.run_bridge(p, combo, "monthly")
    for r in f.rows:
        if r["lr"] is None:
            continue
        want = (res.lr_current * r["trend_factor"] * r["a_rate"] * r["a_mod"]
                * combo.a_other)
        assert r["lr"] == pytest.approx(want, rel=1e-15)
