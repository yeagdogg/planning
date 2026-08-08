"""Oracle test suite (brief §9): worked example, property tests, identities.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import datetime as dt
import math
from dataclasses import replace

import pytest

from src.engine import (
    ComboInputs,
    ModInputs,
    MonthlyEngine,
    PlannedActual,
    RateChange,
    attribution,
    continuous_earned_index,
    continuous_earned_mod,
    continuous_net_index,
    crl_indication,
    month_index,
    run_bridge,
    solve_rate_for_target,
    solver_timing_table,
    validate_inputs,
    worked_example,
)

TAKEN, PLANNED = "taken", "planned"


def _combo(changes=(), seasonality=None, term=12, **kw) -> ComboInputs:
    """Convenience combo with a flat 0.85 mod path (A_mod = 1) unless overridden."""
    mods = kw.pop(
        "mods",
        ModInputs(m_ind=0.85, m0=0.85, m0_asof=dt.date(2026, 9, 30), m1=0.85),
    )
    defaults = dict(lr_proj=0.65)
    defaults.update(kw)
    return ComboInputs(
        mods=mods, rate_changes=tuple(changes), seasonality=seasonality, term_months=term, **defaults
    )


# ---------------------------------------------------------------------------
# §9 worked example
# ---------------------------------------------------------------------------


class TestWorkedExampleContinuous:
    """Continuous closed-form must match the brief's stated values to 1e-6."""

    def setup_method(self):
        self.p, self.combo = worked_example()
        self.res = run_bridge(self.p, self.combo, convention="continuous")

    def test_crl(self):
        assert self.res.crl_ind == pytest.approx(1.1024, abs=1e-9)

    def test_e_cy(self):
        # areas: 0.125 @ 1.0400, 0.59375 @ 1.1024, 0.28125 @ 1.157520
        assert self.res.e_cy[2027] == pytest.approx(1.1101025, abs=1e-6)

    def test_a_rate(self):
        assert self.res.a_rate_p == pytest.approx(1.1024 / 1.1101025, abs=1e-6)
        assert round(self.res.a_rate_p, 5) == 0.99306

    def test_earned_mod_identity_value(self):
        # linear-mod identity: written mod at 1/1/2027 = 0.860 + 3mo x 0.002 = 0.866
        assert self.res.m_earned[2027] == pytest.approx(0.8660, abs=1e-6)

    def test_a_mod(self):
        assert self.res.a_mod_p == pytest.approx(0.850 / 0.866, abs=1e-6)
        assert round(self.res.a_mod_p, 5) == 0.98152

    def test_cy_plan_lr(self):
        expected = 0.65 * (1.1024 / 1.1101025) * (0.850 / 0.866)
        assert self.res.cy_lr_p == pytest.approx(expected, abs=1e-9)
        assert round(self.res.cy_lr_p, 4) == 0.6336  # 63.36%

    def test_p_plus_1(self):
        # CY 2028: 0.03125 @ 1.1024 (Jan-Mar 2027 writings), 0.96875 @ 1.157520
        e_2028 = 0.03125 * 1.1024 + 0.96875 * 1.157520
        assert self.res.e_cy[2028] == pytest.approx(e_2028, abs=1e-6)
        assert self.res.a_rate_p1 == pytest.approx(1.1024 / e_2028, abs=1e-6)


class TestWorkedExampleMonthly:
    """Monthly convention: within ±0.10 LR points of continuous (brief §9)."""

    def setup_method(self):
        self.p, self.combo = worked_example()
        self.mon = run_bridge(self.p, self.combo, convention="monthly")
        self.cont = run_bridge(self.p, self.combo, convention="continuous")

    def test_cy_lr_within_tolerance(self):
        assert abs(self.mon.cy_lr_p - self.cont.cy_lr_p) < 0.0010  # 0.10 pts

    def test_index_close(self):
        assert self.mon.e_cy[2027] == pytest.approx(self.cont.e_cy[2027], abs=7e-4)
        assert self.mon.e_cy[2028] == pytest.approx(self.cont.e_cy[2028], abs=7e-4)

    def test_mod_close(self):
        assert self.mon.m_earned[2027] == pytest.approx(0.8660, abs=2e-4)

    def test_no_warnings(self):
        assert self.mon.warnings == []


# ---------------------------------------------------------------------------
# §3.2 directionality properties
# ---------------------------------------------------------------------------


class TestDirectionality:
    def test_planned_increase_not_considered_lowers_cy_lr(self):
        """Planned increase, considered=N  =>  A_rate < 1."""
        changes = [RateChange(dt.date(2027, 4, 1), 0.05, PLANNED, considered=False)]
        for conv in ("monthly", "continuous"):
            res = run_bridge(2027, _combo(changes), convention=conv)
            assert res.a_rate_p < 1.0, conv
            assert res.cy_lr_p < 0.65, conv

    def test_considered_partially_earned_raises_cy_lr(self):
        """Taken+considered but not fully earned in CY P  =>  A_rate > 1."""
        changes = [RateChange(dt.date(2026, 7, 1), 0.06, TAKEN, considered=True)]
        for conv in ("monthly", "continuous"):
            res = run_bridge(2027, _combo(changes), convention=conv)
            assert res.a_rate_p > 1.0, conv

    def test_rising_mods_give_a_mod_below_1(self):
        """Mods rising above M_ind  =>  A_mod < 1."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        for conv in ("monthly", "continuous"):
            res = run_bridge(2027, _combo(mods=mods), convention=conv)
            assert res.a_mod_p < 1.0, conv

    def test_fully_earned_considered_change_is_neutral(self):
        """A considered change taken long ago is fully earned: A_rate = 1."""
        changes = [RateChange(dt.date(2024, 1, 1), 0.08, TAKEN, considered=True)]
        res = run_bridge(2027, _combo(changes))
        assert res.a_rate_p == pytest.approx(1.0, abs=1e-12)


# ---------------------------------------------------------------------------
# §3.3.5 linear-mod identity
# ---------------------------------------------------------------------------


class TestLinearModIdentity:
    def test_monthly_identity(self):
        """Uniform writings, annual term, linear path: CY earned mod = written mod at 1/1."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        eng = MonthlyEngine(2027, _combo(mods=mods))
        jan1 = float(dt.date(2027, 1, 1).toordinal())  # start-of-day 1/1
        assert eng.earned_mod_cy(2027) == pytest.approx(eng.mod_path.value(jan1), abs=2e-4)

    def test_continuous_identity_exact(self):
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        m_bar = continuous_earned_mod(2027, mods, 2027, 12)
        # month-coordinate written mod at 1/1/2027: 0.860 + 3/15 x 0.03 = 0.866
        assert m_bar == pytest.approx(0.866, abs=1e-9)


# ---------------------------------------------------------------------------
# Monthly vs continuous convergence
# ---------------------------------------------------------------------------


class TestConvergence:
    LOGS = [
        [],
        [RateChange(dt.date(2026, 3, 1), 0.03, TAKEN, True)],
        [
            RateChange(dt.date(2025, 10, 1), -0.02, TAKEN, True),
            RateChange(dt.date(2026, 6, 1), 0.045, TAKEN, True),
            RateChange(dt.date(2027, 2, 1), 0.07, PLANNED, False),
        ],
        [
            RateChange(dt.date(2026, 1, 1), 0.10, TAKEN, True),
            RateChange(dt.date(2027, 9, 1), 0.05, PLANNED, False, achievement=0.8),
            RateChange(dt.date(2028, 4, 1), 0.03, PLANNED, False),
        ],
    ]

    @pytest.mark.parametrize("log", LOGS)
    @pytest.mark.parametrize("term", [12, 6])
    def test_first_of_month_logs_converge(self, log, term):
        combo = _combo(log, term=term)
        for year in (2026, 2027, 2028):
            mon = MonthlyEngine(2027, combo).earned_index_cy(year)
            cont = continuous_earned_index(combo.rate_changes, year, term)
            assert mon == pytest.approx(cont, abs=1.5e-3), (year, term)

    def test_midmonth_effective_date(self):
        log = [RateChange(dt.date(2027, 4, 15), 0.05, PLANNED, False)]
        combo = _combo(log)
        mon = MonthlyEngine(2027, combo).earned_index_cy(2027)
        cont = continuous_earned_index(combo.rate_changes, 2027, 12)
        assert mon == pytest.approx(cont, abs=1.5e-3)


# ---------------------------------------------------------------------------
# Degenerate inputs (§3.5)
# ---------------------------------------------------------------------------


class TestDegenerate:
    def test_no_changes_flat_mods(self):
        res = run_bridge(2027, _combo())
        assert res.crl_ind == 1.0
        assert res.a_rate_p == pytest.approx(1.0, abs=1e-12)
        assert res.a_mod_p == pytest.approx(1.0, abs=1e-12)
        assert res.cy_lr_p == pytest.approx(0.65, abs=1e-12)
        assert res.cy_lr_p1 == pytest.approx(0.65, abs=1e-12)

    def test_mod_adjustment_disabled(self):
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        res = run_bridge(2027, _combo(mods=mods, mod_adjustment_enabled=False))
        assert res.a_mod_p == 1.0
        assert res.cy_lr_p == pytest.approx(0.65, abs=1e-12)

    def test_the_projected_lr_is_taken_at_current_level_as_entered(self):
        """v3.8 (D107): the basis toggle is gone, so lr_current IS the input.
        This replaces test_basis_proposed, which proved the x(1+s) conversion."""
        assert run_bridge(2027, _combo()).lr_current == pytest.approx(0.65, abs=1e-12)
        assert run_bridge(2027, _combo(lr_proj=0.702)).lr_current == pytest.approx(
            0.702, abs=1e-12)

    def test_net_trend_applies_to_p1_only(self):
        res = run_bridge(2027, _combo(net_trend=0.02))
        assert res.cy_lr_p == pytest.approx(0.65, abs=1e-12)
        assert res.cy_lr_p1 == pytest.approx(0.65 * 1.02, abs=1e-12)


# ---------------------------------------------------------------------------
# Mid-month blend mechanics (D4)
# ---------------------------------------------------------------------------


class TestMidMonthBlend:
    def test_blend_equals_exact_daily_single_change(self):
        log = [RateChange(dt.date(2027, 4, 15), 0.05, TAKEN, False)]
        eng = MonthlyEngine(2027, _combo(log))
        k = month_index(2027, 4)
        assert eng.written_index(k) == pytest.approx(eng.written_index_exact_daily(k), abs=1e-12)

    def test_first_of_month_gets_full_new_rate(self):
        log = [RateChange(dt.date(2027, 4, 1), 0.05, TAKEN, False)]
        eng = MonthlyEngine(2027, _combo(log))
        assert eng.written_index(month_index(2027, 4)) == pytest.approx(1.05, abs=1e-12)

    def test_blend_weights(self):
        # eff 4/16/2027: 15 days pre, 15 of 30 on/after -> p = 0.5
        log = [RateChange(dt.date(2027, 4, 16), 0.10, TAKEN, False)]
        eng = MonthlyEngine(2027, _combo(log))
        assert eng.written_index(month_index(2027, 4)) == pytest.approx(1.05, abs=1e-12)


# ---------------------------------------------------------------------------
# Earning profile / closed-form ec equivalence (D10 — underpins the Excel _calc)
# ---------------------------------------------------------------------------


class TestEarningProfile:
    @pytest.mark.parametrize("term", [12, 6, 1])
    def test_profile_sums_to_one(self, term):
        eng = MonthlyEngine(2027, _combo(term=term))
        k = month_index(2026, 5)
        assert sum(eng.e(k, m) for m in range(k, k + term + 1)) == pytest.approx(1.0, abs=1e-12)

    @pytest.mark.parametrize("term", [12, 6])
    def test_ec_closed_form_equals_matrix_sums(self, term):
        eng = MonthlyEngine(2027, _combo(term=term))
        for k in eng.cohorts:
            for year in (2025, 2026, 2027, 2028, 2029):
                explicit = sum(
                    eng.e(k, m) for m in range(month_index(year, 1), month_index(year, 12) + 1)
                )
                assert eng.ec(k, year) == pytest.approx(explicit, abs=1e-12), (k, year)

    def test_interior_cohort_fully_earned_within_three_years(self):
        eng = MonthlyEngine(2027, _combo())
        k = month_index(2026, 8)
        total = sum(eng.ec(k, y) for y in (2026, 2027, 2028))
        assert total == pytest.approx(1.0, abs=1e-12)


# ---------------------------------------------------------------------------
# Seasonality (§3.2.4)
# ---------------------------------------------------------------------------


class TestSeasonality:
    def test_flat_profile_equals_uniform(self):
        log = [RateChange(dt.date(2026, 7, 1), 0.06, TAKEN, True)]
        flat = run_bridge(2027, _combo(log, seasonality=tuple([1.0] * 12)))
        unif = run_bridge(2027, _combo(log))
        assert flat.e_cy[2027] == pytest.approx(unif.e_cy[2027], abs=1e-12)

    def test_concentration_shifts_earned_index(self):
        """All writings in January: a 7/1/2026 change is fully in every 2027 cohort
        (written 1/2027) and mostly earned by 2027 cohorts -> E_CY(2027) closer to
        the post-change level than under uniform writings."""
        log = [RateChange(dt.date(2026, 7, 1), 0.06, TAKEN, True)]
        jan_only = tuple([1.0] + [0.0] * 11)
        conc = run_bridge(2027, _combo(log, seasonality=jan_only))
        unif = run_bridge(2027, _combo(log))
        assert conc.e_cy[2027] > unif.e_cy[2027]


# ---------------------------------------------------------------------------
# CRL / validation
# ---------------------------------------------------------------------------


class TestValidationAndCrl:
    def test_crl_is_product_of_considered_rows(self):
        log = [
            RateChange(dt.date(2025, 7, 1), 0.04, TAKEN, True),
            RateChange(dt.date(2026, 7, 1), 0.06, TAKEN, True),
            RateChange(dt.date(2027, 4, 1), 0.05, PLANNED, False),
        ]
        assert crl_indication(log) == pytest.approx(1.04 * 1.06, abs=1e-12)

    def test_planned_considered_warns_but_counts(self):
        log = [RateChange(dt.date(2027, 4, 1), 0.05, PLANNED, considered=True, achievement=0.8)]
        combo = _combo(log)
        warnings = validate_inputs(2027, combo)
        assert any("considered" in w for w in warnings)
        assert crl_indication(log) == pytest.approx(1.04, abs=1e-12)  # 5% x 80%

    def test_achievement_scales_planned_only(self):
        planned = RateChange(dt.date(2027, 4, 1), 0.05, PLANNED, False, achievement=0.8)
        taken = RateChange(dt.date(2026, 4, 1), 0.05, TAKEN, False, achievement=0.8)
        assert planned.effective_pct == pytest.approx(0.04, abs=1e-12)
        assert taken.effective_pct == pytest.approx(0.05, abs=1e-12)

    def test_two_changes_one_month_warns(self):
        log = [
            RateChange(dt.date(2027, 4, 1), 0.02, TAKEN, False),
            RateChange(dt.date(2027, 4, 20), 0.02, TAKEN, False),
        ]
        assert any("share cohort month" in w for w in validate_inputs(2027, _combo(log)))

    def test_bad_term_raises(self):
        """The other input the engine cannot compute WITH rather than warn about.
        (The basis check that used to sit beside this went with D107.)"""
        with pytest.raises(ValueError):
            run_bridge(2027, _combo(term=0))
        with pytest.raises(ValueError):
            run_bridge(2027, _combo(term=13))


# ---------------------------------------------------------------------------
# E1 Solver
# ---------------------------------------------------------------------------


class TestSolver:
    def test_mode_a_reproduces_worked_example(self):
        """Acceptance §11.7: solving for the worked example's own CY LR at
        4/1/2027 must return +5.0%."""
        p, combo = worked_example()
        target = run_bridge(p, combo).cy_lr_p
        res = solve_rate_for_target(p, combo, target, dt.date(2027, 4, 1))
        assert res.required_change == pytest.approx(0.05, abs=1e-10)
        assert res.warnings == []

    def test_mode_a_roundtrip_arbitrary(self):
        p, combo = worked_example()
        for r, eff in [(0.03, dt.date(2027, 2, 1)), (-0.02, dt.date(2027, 6, 15))]:
            taken_only = tuple(rc for rc in combo.rate_changes if rc.status == TAKEN)
            probe = ComboInputs(
                **{
                    **combo.__dict__,
                    "rate_changes": taken_only + (RateChange(eff, r, PLANNED, False),),
                }
            )
            target = run_bridge(p, probe).cy_lr_p
            res = solve_rate_for_target(p, combo, target, eff)
            assert res.required_change == pytest.approx(r, abs=1e-10), (r, eff)

    def test_filed_equivalent(self):
        p, combo = worked_example()
        target = run_bridge(p, combo).cy_lr_p
        res = solve_rate_for_target(p, combo, target, dt.date(2027, 4, 1), achievement=0.8)
        assert res.filed_equivalent == pytest.approx(0.05 / 0.8, abs=1e-10)

    def test_late_year_warning(self):
        p, combo = worked_example()
        res = solve_rate_for_target(p, combo, 0.60, dt.date(2027, 12, 15))
        assert any("mathematically absurd" in w for w in res.warnings)
        assert any("reasonability" in w for w in res.warnings)

    def test_roundtrip_with_taken_change_sharing_effective_month(self):
        """Adversarial-review fix (D31): the closed form must invert the
        engine's first-in-month blend even when a taken change shares the
        solver's effective month (both orderings, mid-month dates)."""
        cases = [
            (dt.date(2027, 4, 10), dt.date(2027, 4, 20)),  # taken first
            (dt.date(2027, 4, 10), dt.date(2027, 4, 1)),   # unknown first
            (dt.date(2027, 4, 10), dt.date(2027, 4, 10)),  # same day
        ]
        for taken_eff, solve_eff in cases:
            taken = RateChange(taken_eff, 0.02, TAKEN, considered=False)
            truth = RateChange(solve_eff, 0.05, PLANNED, considered=False)
            probe = _combo([taken, truth])
            target = run_bridge(2027, probe).cy_lr_p
            res = solve_rate_for_target(2027, _combo([taken]), target, solve_eff)
            assert res.required_change == pytest.approx(0.05, abs=1e-10), (taken_eff, solve_eff)

    def test_c_post_zero_returns_nan_with_warning(self):
        p, combo = worked_example()
        res = solve_rate_for_target(p, combo, 0.60, dt.date(p + 1, 1, 1))
        assert math.isnan(res.required_change)
        assert any("No CY" in w for w in res.warnings)
        jan_only = _combo(seasonality=tuple([1.0] + [0.0] * 11))
        res2 = solve_rate_for_target(2027, jan_only, 0.60, dt.date(2027, 2, 1))
        assert math.isnan(res2.required_change)
        assert res2.warnings

    def test_mode_b_consistent_with_mode_a(self):
        p, combo = worked_example()
        rows = solver_timing_table(p, combo, r=0.05, target_cy_lr=0.64)
        apr = next(row for row in rows if row["month"] == 4)
        base = run_bridge(p, combo)
        assert apr["cy_lr"] == pytest.approx(base.cy_lr_p, abs=1e-12)
        # later effective date earns less of the increase -> higher CY LR
        lrs = [row["cy_lr"] for row in rows]
        assert all(b >= a - 1e-12 for a, b in zip(lrs, lrs[1:]))


# ---------------------------------------------------------------------------
# E2 Attribution
# ---------------------------------------------------------------------------


class TestAttribution:
    def test_reconciles_and_directions(self):
        p, combo = worked_example()
        actuals = [
            PlannedActual(
                planned_effective=dt.date(2027, 4, 1),
                actual_pct=0.04,  # achieved 4% instead of 5%
                actual_effective=dt.date(2027, 6, 1),  # two months late
            )
        ]
        mods_actual = ModInputs(m_ind=0.850, m0=0.860, m0_asof=dt.date(2026, 9, 30), m1=0.900)
        res = attribution(p, combo, actuals, mods_actual, actual_cy_lr=0.655)
        assert res.reconciles
        f = dict(res.factors)
        assert f["Rate magnitude"] > 1.0  # under-achieved increase -> higher LR
        assert f["Rate timing"] > 1.0  # later date -> less earned -> higher LR
        assert f["Mod drift"] < 1.0  # mods rose more than planned -> more premium
        lr_plan = res.steps[0][1]
        total_pts = sum(pts for _, pts in res.points)
        assert lr_plan + total_pts / 100.0 == pytest.approx(0.655, abs=1e-12)

    def test_no_deviation_is_pure_residual(self):
        p, combo = worked_example()
        actuals = [
            PlannedActual(dt.date(2027, 4, 1), actual_pct=0.05, actual_effective=dt.date(2027, 4, 1))
        ]
        base = run_bridge(p, combo)
        res = attribution(p, combo, actuals, combo.mods, actual_cy_lr=base.cy_lr_p + 0.01)
        f = dict(res.factors)
        assert f["Rate magnitude"] == pytest.approx(1.0, abs=1e-12)
        assert f["Rate timing"] == pytest.approx(1.0, abs=1e-12)
        assert f["Mod drift"] == pytest.approx(1.0, abs=1e-12)
        assert f["Loss-side residual"] > 1.0

    def test_mod_drift_uses_plan_m_ind(self):
        """D31: an actual-anchors object with a different m_ind must NOT
        manufacture spurious mod drift — M_ind is an indication property."""
        p, combo = worked_example()
        actuals = [
            PlannedActual(dt.date(2027, 4, 1), actual_pct=0.05, actual_effective=dt.date(2027, 4, 1))
        ]
        mods_act = ModInputs(m_ind=0.90, m0=combo.mods.m0, m0_asof=combo.mods.m0_asof,
                             m1=combo.mods.m1)
        base = run_bridge(p, combo)
        res = attribution(p, combo, actuals, mods_act, actual_cy_lr=base.cy_lr_p)
        assert dict(res.factors)["Mod drift"] == pytest.approx(1.0, abs=1e-12)

    def test_unmatched_actual_raises(self):
        p, combo = worked_example()
        bad = [PlannedActual(dt.date(2027, 4, 2), 0.05, dt.date(2027, 4, 2))]
        with pytest.raises(ValueError):
            attribution(p, combo, bad, combo.mods, actual_cy_lr=0.65)


# ---------------------------------------------------------------------------
# Net rate selection (D39)
# ---------------------------------------------------------------------------


class TestNetRateSelection:
    """Optional per-combo net rate selection: from 1/1/P each written cohort's
    combined net price index renews net_sel_p above its year-ago cohort,
    superseding planned rate rows and the projected mod path for those cohorts
    only. History earns exactly as modeled."""

    def _user_example(self, **kw):
        """The user's scenario: +10% taken 10/1/(P-1), +10% net from 1/1/P."""
        log = [RateChange(dt.date(2026, 10, 1), 0.10, TAKEN, considered=True)]
        return _combo(log, net_sel_p=0.10, **kw)

    def test_user_example_written_path(self):
        eng = MonthlyEngine(2027, self._user_example())  # flat mods: q = W
        # history: Q4 2026 cohorts carry the filing; earlier 2026 cohorts do not
        assert eng.index_in_force(month_index(2026, 9)) == pytest.approx(1.00, abs=1e-12)
        assert eng.index_in_force(month_index(2026, 10)) == pytest.approx(1.10, abs=1e-12)
        # Jan-Sep 2027 renew pre-filing cohorts: +10% lands them at 1.10 —
        # the 10/1 filing IS the vehicle that delivers the target for them
        assert eng.index_in_force(month_index(2027, 1)) == pytest.approx(1.10, abs=1e-12)
        assert eng.index_in_force(month_index(2027, 9)) == pytest.approx(1.10, abs=1e-12)
        # Q4 2027 renews already-increased cohorts: +10% net compounds to 1.21
        assert eng.index_in_force(month_index(2027, 10)) == pytest.approx(1.21, abs=1e-12)
        # P+1 carries the selection by default: another +10% on year-ago
        assert eng.index_in_force(month_index(2028, 1)) == pytest.approx(1.21, abs=1e-12)

    def test_history_cohorts_unchanged_by_net_mode(self):
        combo = self._user_example()
        eng_net = MonthlyEngine(2027, combo)
        eng_off = MonthlyEngine(2027, replace(combo, net_sel_p=None))
        for k in eng_net.cohorts:
            if k < month_index(2027, 1):
                assert eng_net.index_in_force(k) == pytest.approx(
                    eng_off.index_in_force(k), abs=1e-12)

    def test_planned_rows_and_mod_projection_superseded(self):
        """A planned change in P and a different M_1 must not move net-mode
        results for CY P earnings from 1/1/P cohorts (history unchanged)."""
        base = self._user_example()
        noisy = replace(
            base,
            rate_changes=base.rate_changes
            + (RateChange(dt.date(2027, 6, 1), 0.07, PLANNED, considered=False),),
        )
        r_base = run_bridge(2027, base)
        r_noisy = run_bridge(2027, noisy)
        assert r_noisy.e_cy[2027] == pytest.approx(r_base.e_cy[2027], abs=1e-12)
        assert any("superseded" in w for w in r_noisy.warnings)

    def test_a_mod_collapses_to_one(self):
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        res = run_bridge(2027, self._user_example(mods=mods))
        assert res.a_mod_p == 1.0 and res.a_mod_p1 == 1.0
        assert res.cy_lr_p == pytest.approx(
            res.lr_current * res.crl_ind / res.e_cy[2027], abs=1e-12)

    def test_historical_mod_drift_enters_combined_index(self):
        """Cohorts before 1/1/P carry W x M/M_ind; rising mods lift q above W."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        eng = MonthlyEngine(2027, self._user_example(mods=mods))
        k = month_index(2026, 11)
        assert eng.index_in_force(k) > eng.written_index(k)
        assert eng.index_in_force(k) == pytest.approx(
            eng.written_index(k) * eng.written_mod(k) / 0.85, abs=1e-12)

    def test_net_zero_is_flat_yoy(self):
        eng = MonthlyEngine(2027, replace(self._user_example(), net_sel_p=0.0))
        for m in range(1, 13):
            assert eng.index_in_force(month_index(2027, m)) == pytest.approx(
                eng.index_in_force(month_index(2026, m)), abs=1e-12)

    def test_p1_selection_override(self):
        carried = run_bridge(2027, self._user_example())
        flat_p1 = run_bridge(2027, replace(self._user_example(), net_sel_p1=0.0))
        assert flat_p1.e_cy[2028] < carried.e_cy[2028]
        eng = MonthlyEngine(2027, replace(self._user_example(), net_sel_p1=0.0))
        assert eng.index_in_force(month_index(2028, 3)) == pytest.approx(
            eng.index_in_force(month_index(2027, 3)), abs=1e-12)

    def test_monthly_vs_continuous_net(self):
        """The genuine monthly-vs-continuous convention gap is ~1e-5. The
        pre-W3b tolerance here was 1.5e-3 — wide enough to hide the Simpson
        segment-branching defect (~1e-3) entirely."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        combo = self._user_example(mods=mods)
        mon = run_bridge(2027, combo, "monthly")
        for year in (2026, 2027, 2028):
            cont = continuous_net_index(2027, combo, year)
            assert mon.e_cy[year] == pytest.approx(cont, abs=5e-5), year

    @staticmethod
    def _brute_net_index(combo, plan_year, year, per_month=256):
        """Independent midpoint integration of L(t)·q(t). Grid boundaries
        sit on integer month coordinates (every jump in these cases is a
        1st-of-month / close-of-month coordinate), so no cell straddles a
        discontinuity and the only error is O(h^2) curvature — ~1e-8 at
        256 cells/month, three orders below the tolerance and four below
        the pre-W3b Simpson defect this pins."""
        from src.engine import _cy_overlap_weight, continuous_net_q

        tau = float(combo.term_months)
        y0 = float(month_index(year, 1))
        y1 = float(month_index(year + 1, 1))
        q, _jumps, _kinks = continuous_net_q(plan_year, combo)
        lo = y0 - tau
        n = int(round((y1 - lo) * per_month))
        h = (y1 - lo) / n
        num = den = 0.0
        for i in range(n):
            t = lo + (i + 0.5) * h
            w = _cy_overlap_weight(t, y0, y1, tau)
            den += w
            num += w * q(t)
        return num / den

    def test_continuous_net_ties_brute_force(self):
        """The Simpson assembly vs an independent integral, all three years.
        Pre-W3b this failed at ~2e-4..1.4e-3: a segment ending exactly on
        1/1/P (always a breakpoint) took the recursion branch for its
        left-limit endpoint. The 10/1 history filing makes the sawtooth
        pronounced, which is exactly the shape that exposed it."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        combo = self._user_example(mods=mods)
        for year in (2026, 2027, 2028):
            got = continuous_net_index(2027, combo, year)
            want = self._brute_net_index(combo, 2027, year)
            assert got == pytest.approx(want, abs=1e-7), year

    def test_continuous_net_honors_the_mod_reanchor(self):
        """D70a: once a mod log exists, HISTORY re-anchors onto m_end_prior
        — which moves every P-1 cohort's q and, through the recursion, the
        whole net path. Pre-W3b the continuous side used a bare drift path
        and returned the no-steps value identically (~1e-2 off monthly).

        Monthly-vs-continuous rides a looser tolerance here than the
        no-steps test: the monthly engine interpolates mod anchors in
        DAY-ordinal space (_cod / mi_mid_ordinal) while every continuous
        twin interpolates the same anchor values in MONTH coordinates —
        a long-standing convention pair whose gap scales with drift slope,
        and the re-anchor tail (0.860 -> 0.875 over a quarter) is steep
        (~4e-4 observed). The at-convention truth is the brute-force tie."""
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30),
                         m1=0.89, m_end_prior=0.875)
        steps = (RateChange(dt.date(2027, 3, 1), 0.03, TAKEN,
                            considered=False),)
        combo = replace(self._user_example(mods=mods), mod_changes=steps)
        bare = self._user_example(mods=mods)          # same mods, no log
        mon = run_bridge(2027, combo, "monthly")
        for year in (2026, 2027, 2028):
            cont = continuous_net_index(2027, combo, year)
            assert cont == pytest.approx(
                self._brute_net_index(combo, 2027, year), abs=1e-7), year
            assert mon.e_cy[year] == pytest.approx(cont, abs=1e-3), year
        # and the re-anchor genuinely moves the number (else this test
        # could pass with the defect present)
        assert abs(continuous_net_index(2027, combo, 2026)
                   - continuous_net_index(2027, bare, 2026)) > 1e-3

    def test_modadj_off_is_rate_only_with_warning(self):
        mods = ModInputs(m_ind=0.85, m0=0.86, m0_asof=dt.date(2026, 9, 30), m1=0.89)
        combo = self._user_example(mods=mods, mod_adjustment_enabled=False)
        eng = MonthlyEngine(2027, combo)
        k = month_index(2026, 11)
        assert eng.index_in_force(k) == pytest.approx(eng.written_index(k), abs=1e-12)
        assert any("rate-only" in w for w in validate_inputs(2027, combo))

    def test_off_by_default(self):
        p, combo = worked_example()
        assert not combo.net_mode
        assert run_bridge(p, combo).a_mod_p != 1.0  # classic decomposition intact


# ---------------------------------------------------------------------------
# Term override (6-month LOB)
# ---------------------------------------------------------------------------


class TestTermOverride:
    def test_six_month_term_earns_faster(self):
        """A 6-month book earns rate changes in faster. A change effective
        7/1/(P-1) is FULLY earned by CY P on a 6-month book (A_rate = 1
        exactly), while the annual book still carries old-rate premium into P
        (A_rate > 1). A change effective 1/1/P is partially earned under both
        terms, but less of a shortfall on the 6-month book."""
        log_p_minus_1 = [RateChange(dt.date(2026, 7, 1), 0.06, TAKEN, True)]
        assert run_bridge(2027, _combo(log_p_minus_1, term=6)).a_rate_p == pytest.approx(
            1.0, abs=1e-12
        )
        assert run_bridge(2027, _combo(log_p_minus_1, term=12)).a_rate_p > 1.0

        log_jan_p = [RateChange(dt.date(2027, 1, 1), 0.06, TAKEN, True)]
        r12 = run_bridge(2027, _combo(log_jan_p, term=12))
        r6 = run_bridge(2027, _combo(log_jan_p, term=6))
        assert 1.0 < r6.a_rate_p < r12.a_rate_p

    def test_term_must_be_valid(self):
        with pytest.raises(ValueError):
            run_bridge(2027, _combo(term=0))
        with pytest.raises(ValueError):
            run_bridge(2027, _combo(term=13))
