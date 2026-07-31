"""Oracle engine: calendar-year plan loss ratio bridge.

This module is the authoritative Python implementation of the methodology in the
build brief (standard on-leveling / current-rate-level methodology, cf. Werner &
Modlin, *Basic Ratemaking*, Ch. 5, extended with a schedule-mod adjustment).
The generated Excel workbook must tie to this engine; the pytest suite in
``tests/test_engine.py`` locks the behavior down.

Two engines are implemented:

* **Monthly engine** (:class:`MonthlyEngine`) — the production convention:
  monthly written cohorts, mid-month writing, the 1/(2T) .. 1/T .. 1/(2T)
  earning profile, two-point pro-rata blending for mid-month effective dates,
  and piecewise-linear mod paths interpolated on date serials. The Excel
  workbook reproduces THIS engine exactly.

* **Continuous engine** (closed-form parallelogram, :func:`continuous_earned_index`
  / :func:`continuous_earned_mod`) — textbook geometry measured in month units
  (each month = 1/12 year). Used as an independent cross-check; the monthly
  engine must agree within tolerance under uniform/annual settings.

Timeline conventions (DECISIONS.md D2/D3):

* Rate-change effective dates are **start-of-day**: a change effective on the
  1st applies to the whole month's writings.
* Schedule-mod as-of dates are **close-of-day**: "M_0 as of 9/30" anchors the
  mod path at the 10/1 boundary. In serial terms the anchor x is ``date + 1``.

Cohort window (DECISIONS.md D1): Jan (P-2) .. Dec (P+1), 48 months, so that
E_CY(P-1) (needed for the year-over-year earned rate change) is computable.
Earned months are reported for Jan (P-1) .. Dec (P+1), 36 months.
"""

from __future__ import annotations

import datetime as dt
import math
from dataclasses import dataclass, field, replace
from typing import Iterable, Sequence

# ---------------------------------------------------------------------------
# Month/date helpers
# ---------------------------------------------------------------------------


def month_index(year: int, month: int) -> int:
    """Absolute month index: Jan of year Y -> Y*12, Dec -> Y*12 + 11."""
    return year * 12 + (month - 1)


def mi_year(mi: int) -> int:
    return mi // 12


def mi_month(mi: int) -> int:
    return mi % 12 + 1


def mi_first(mi: int) -> dt.date:
    """First day of the month with index ``mi``."""
    return dt.date(mi_year(mi), mi_month(mi), 1)


def mi_days(mi: int) -> int:
    """Number of days in month ``mi``."""
    return (mi_first(mi + 1) - mi_first(mi)).days


def mi_last(mi: int) -> dt.date:
    """Last day of the month with index ``mi``."""
    return mi_first(mi + 1) - dt.timedelta(days=1)


def mi_mid_ordinal(mi: int) -> float:
    """Mid-month instant as a (fractional) proleptic ordinal.

    Midpoint of [start-of-month, start-of-next-month). Matches the Excel
    formula (SOM + EOM + 1) / 2 up to the constant Excel-serial offset, which
    cancels in linear interpolation.
    """
    return (mi_first(mi).toordinal() + mi_first(mi + 1).toordinal()) / 2.0


def add_months(d: dt.date, n: int) -> dt.date:
    """Shift a date by ``n`` calendar months, clamping the day to month end."""
    y, m0 = divmod(d.month - 1 + n, 12)
    y += d.year
    m = m0 + 1
    day = min(d.day, (dt.date(y + (m == 12), m % 12 + 1, 1) - dt.timedelta(days=1)).day)
    return dt.date(y, m, day)


def month_coord(d: dt.date, end_of_day: bool = False) -> float:
    """Continuous-time coordinate in month units (each month = 1/12 year).

    Start-of-day: the 1st of a month sits exactly on the month boundary.
    End-of-day: the last day of a month sits exactly on the next boundary
    (used for mod as-of anchors, DECISIONS.md D2/D3).
    """
    mi = month_index(d.year, d.month)
    frac = (d.day - (0 if end_of_day else 1)) / mi_days(mi)
    return mi + frac


# ---------------------------------------------------------------------------
# Input dataclasses
# ---------------------------------------------------------------------------

STATUS_TAKEN = "taken"
STATUS_PLANNED = "planned"


@dataclass(frozen=True)
class RateChange:
    """One row of the rate change log."""

    effective: dt.date
    filed_pct: float  # decimal fraction, e.g. +5.0% -> 0.05
    status: str  # 'taken' | 'planned'
    considered: bool  # considered in the indication?
    achievement: float = 1.0  # applies to planned rows only
    comment: str = ""

    @property
    def effective_pct(self) -> float:
        """Filed % x achievement % for planned rows; filed % for taken rows."""
        if self.status == STATUS_PLANNED:
            return self.filed_pct * self.achievement
        return self.filed_pct


@dataclass(frozen=True)
class ModInputs:
    """Schedule-mod path anchors. All 'as-of' dates are close-of-day (D2)."""

    m_ind: float  # avg mod assumed in the indication
    m0: float  # current avg written mod
    m0_asof: dt.date  # as-of date for m0
    m1: float  # projected avg written mod at 12/31/P
    m_prior: float | None = None  # optional: avg mod ~12 months before m0
    m2: float | None = None  # optional: projected at 12/31/(P+1); default m1


@dataclass(frozen=True)
class ComboInputs:
    """All inputs for one combo (a BU x state within one LOB workbook).

    Net rate selection (optional; DECISIONS.md D39): when ``net_sel_p`` is not
    None, cohorts written on/after 1/1/P abandon the planned rate program and
    the projected mod path; instead each such cohort's COMBINED net price index
    q(k) = W(k) x M_w(k)/M_ind renews at ``net_sel_p`` above its year-ago
    cohort: q(k) = q(k-12) x (1 + x). History (cohorts before 1/1/P) stays
    fully modeled — it is the carryover half of CY P earnings. ``net_sel_p1``
    covers cohorts written in P+1 (None = carry ``net_sel_p``).
    """

    lr_proj: float  # projected loss ratio from the indication
    lr_basis: str  # 'current' | 'proposed'
    mods: ModInputs
    rate_changes: tuple[RateChange, ...] = ()
    sel_change: float = 0.0  # selected indication change s (for basis conversion)
    net_trend: float = 0.0  # annual net loss-over-premium trend for P+1
    a_other: float = 1.0  # manual adjustment factor
    a_other_label: str = ""
    mod_adjustment_enabled: bool = True
    term_months: int = 12
    seasonality: tuple[float, ...] | None = None  # 12 monthly written weights
    plan_ep: float | None = None  # plan earned premium (portfolio weighting)
    net_sel_p: float | None = None  # optional net rate selection for plan year P
    net_sel_p1: float | None = None  # optional net selection for P+1 (None = carry P's)

    @property
    def net_mode(self) -> bool:
        return self.net_sel_p is not None

    @property
    def net_x1(self) -> float:
        if self.net_sel_p1 is not None:
            return self.net_sel_p1
        return self.net_sel_p if self.net_sel_p is not None else 0.0


# ---------------------------------------------------------------------------
# Result container
# ---------------------------------------------------------------------------


@dataclass
class EngineResult:
    plan_year: int
    convention: str  # 'monthly' | 'continuous'
    crl_ind: float
    e_cy: dict  # year -> earned rate index (P-1, P, P+1)
    m_earned: dict  # year -> earned mod (P, P+1; P-1 too for monthly)
    a_rate_p: float
    a_rate_p1: float
    a_mod_p: float
    a_mod_p1: float
    lr_current: float
    cy_lr_p: float
    cy_lr_p1: float
    earned_chg_vs_ind: float  # E_CY(P)/CRL_ind - 1
    yoy_earned_p: float  # E_CY(P)/E_CY(P-1) - 1
    yoy_earned_p1: float  # E_CY(P+1)/E_CY(P) - 1
    warnings: list = field(default_factory=list)
    # Monthly-engine series for workbook verification (empty for continuous):
    cohort_mi: list = field(default_factory=list)  # absolute month indices
    inforce_index: list = field(default_factory=list)  # index in force per cohort (D39)
    written_index: list = field(default_factory=list)  # W_k per cohort
    written_mod: list = field(default_factory=list)  # M_w(k) per cohort
    written_weight: list = field(default_factory=list)  # w_k per cohort
    earned_month_mi: list = field(default_factory=list)  # reported months
    earned_index_m: list = field(default_factory=list)  # E_m per reported month
    earned_mod_m: list = field(default_factory=list)  # earned mod per month


# ---------------------------------------------------------------------------
# Validation
# ---------------------------------------------------------------------------


def validate_inputs(plan_year: int, combo: ComboInputs) -> list[str]:
    """Input sanity checks mirrored by the workbook's Checks panel.

    Returns a list of human-readable warning strings (empty = clean).
    Raises ValueError only for inputs the engine cannot compute with.
    """
    w: list[str] = []
    p = plan_year
    if combo.lr_basis not in ("current", "proposed"):
        raise ValueError(f"lr_basis must be 'current' or 'proposed', got {combo.lr_basis!r}")
    if not 1 <= combo.term_months <= 12:
        raise ValueError(f"term_months must be in 1..12, got {combo.term_months}")
    lo, hi = mi_first(month_index(p - 2, 1)), mi_last(month_index(p + 1, 12))
    by_month: dict[int, int] = {}
    for rc in combo.rate_changes:
        if rc.status not in (STATUS_TAKEN, STATUS_PLANNED):
            raise ValueError(f"rate change status must be taken/planned, got {rc.status!r}")
        if rc.effective < lo:
            w.append(
                f"Rate change {rc.effective} {rc.filed_pct:+.1%} predates the {lo}..{hi} "
                "engine window; it compounds into the base index for all cohorts."
            )
        elif rc.effective > hi:
            w.append(
                f"Rate change {rc.effective} {rc.filed_pct:+.1%} is after the {lo}..{hi} "
                "engine window; it never enters any cohort's written index and affects "
                "results only through CRL_ind if flagged considered."
            )
        if rc.status == STATUS_PLANNED and rc.considered:
            w.append(
                f"Planned change {rc.effective} {rc.filed_pct:+.1%} is flagged "
                "considered-in-indication; considered rows should be taken rows."
            )
        if not 0.0 <= rc.achievement <= 1.5:
            w.append(f"Achievement {rc.achievement:.0%} outside [0%, 150%] for {rc.effective}.")
        if 1.0 + rc.effective_pct <= 0.0:
            raise ValueError(f"Rate change {rc.effective} implies a non-positive index.")
        mi = month_index(rc.effective.year, rc.effective.month)
        by_month[mi] = by_month.get(mi, 0) + 1
    for mi, n in sorted(by_month.items()):
        if n > 1:
            w.append(
                f"{n} rate changes share cohort month {mi_first(mi):%Y-%m}; the intra-month "
                "split blends at the first change date (DECISIONS.md D4)."
            )
    m = combo.mods
    for label, val in (("M_ind", m.m_ind), ("M_0", m.m0), ("M_1", m.m1)):
        if not 0.5 <= val <= 1.5:
            w.append(f"{label} = {val:.3f} outside [0.5, 1.5].")
    if m.m_prior is not None and not 0.5 <= m.m_prior <= 1.5:
        w.append(f"M_prior = {m.m_prior:.3f} outside [0.5, 1.5].")
    if m.m2 is not None and not 0.5 <= m.m2 <= 1.5:
        w.append(f"M_2 = {m.m2:.3f} outside [0.5, 1.5].")
    if not m.m0_asof < dt.date(p, 12, 31):
        raise ValueError("M_0 as-of date must precede 12/31 of the plan year.")
    if combo.seasonality is not None:
        if len(combo.seasonality) != 12:
            raise ValueError("Seasonality vector must have 12 monthly weights.")
        if min(combo.seasonality) < 0 or sum(combo.seasonality) <= 0:
            raise ValueError("Seasonality weights must be >= 0 with a positive sum.")
    if combo.net_sel_p is not None:
        for label, val in (("net_sel_p", combo.net_sel_p),
                           ("net_sel_p1", combo.net_sel_p1)):
            if val is not None and not -0.5 <= val <= 1.0:
                w.append(f"Net rate selection {label} = {val:+.1%} outside [-50%, +100%].")
        if not combo.mod_adjustment_enabled:
            w.append(
                "Net rate selection with mod adjustment OFF: the target is treated as "
                "rate-only (historical mod drift excluded from the combined index)."
            )
        planned_in_p = [rc for rc in combo.rate_changes
                        if rc.status == STATUS_PLANNED
                        and rc.effective >= dt.date(p, 1, 1)]
        if planned_in_p:
            w.append(
                f"{len(planned_in_p)} planned rate change(s) on/after 1/1/{p} are "
                "superseded by the net rate selection for cohorts written from 1/1/P "
                "(they still affect nothing else)."
            )
    return w


# ---------------------------------------------------------------------------
# Schedule-mod path
# ---------------------------------------------------------------------------


class ModPath:
    """Piecewise-linear written-mod path on close-of-day anchors (D2/D6).

    Anchor x-coordinates are supplied by ``coord``: a callable mapping an
    as-of date to a float (close-of-day). Beyond the last anchor the last
    segment extends linearly; before the first anchor the first segment
    extends linearly (which is the M_0->M_1 line when M_prior is absent).
    """

    def __init__(self, plan_year: int, mods: ModInputs, coord):
        anchors: list[tuple[float, float]] = []
        if mods.m_prior is not None:
            anchors.append((coord(add_months(mods.m0_asof, -12)), mods.m_prior))
        anchors.append((coord(mods.m0_asof), mods.m0))
        anchors.append((coord(dt.date(plan_year, 12, 31)), mods.m1))
        m2 = mods.m2 if mods.m2 is not None else mods.m1
        anchors.append((coord(dt.date(plan_year + 1, 12, 31)), m2))
        xs = [a[0] for a in anchors]
        if any(x1 <= x0 for x0, x1 in zip(xs, xs[1:])):
            raise ValueError(f"Mod path anchors must be strictly increasing in time: {anchors}")
        self.anchors = anchors

    def value(self, x: float) -> float:
        a = self.anchors
        if x <= a[0][0]:
            i = 0
        elif x >= a[-1][0]:
            i = len(a) - 2
        else:
            i = max(j for j in range(len(a) - 1) if a[j][0] <= x)
        (x0, y0), (x1, y1) = a[i], a[i + 1]
        return y0 + (y1 - y0) * (x - x0) / (x1 - x0)


# ---------------------------------------------------------------------------
# Monthly (production-convention) engine
# ---------------------------------------------------------------------------


class MonthlyEngine:
    """Monthly written-index x earning-matrix engine (brief §3.2 / §3.3).

    ``changes_override`` substitutes a different rate-change list (used by the
    Solver, Scenarios, and Attribution) while keeping all other combo inputs.
    """

    def __init__(
        self,
        plan_year: int,
        combo: ComboInputs,
        changes_override: Sequence[RateChange] | None = None,
        mods_override: ModInputs | None = None,
    ):
        self.p = plan_year
        self.combo = combo
        self.changes = tuple(changes_override if changes_override is not None else combo.rate_changes)
        self.mods = mods_override if mods_override is not None else combo.mods
        self.t = combo.term_months
        self.c0 = month_index(plan_year - 2, 1)  # first cohort month (D1)
        self.c1 = month_index(plan_year + 1, 12)  # last cohort month
        self.cohorts = list(range(self.c0, self.c1 + 1))  # 48 months
        self.r0 = month_index(plan_year - 1, 1)  # first reported earned month
        self.r1 = self.c1  # last reported earned month
        self.mod_path = ModPath(plan_year, self.mods, coord=lambda d: float(d.toordinal() + 1))
        # Seasonality: relative weights by calendar month, normalized to mean 1.
        if combo.seasonality is not None:
            s = combo.seasonality
            scale = 12.0 / sum(s)
            self._season = [x * scale for x in s]
        else:
            self._season = None

    # -- written weights ----------------------------------------------------
    def w(self, k: int) -> float:
        if self._season is None:
            return 1.0
        return self._season[mi_month(k) - 1]

    # -- written rate index -------------------------------------------------
    def index_at(self, d: dt.date) -> float:
        """Cumulative rate index at start-of-day ``d`` (changes eff <= d)."""
        idx = 1.0
        for rc in self.changes:
            if rc.effective <= d:
                idx *= 1.0 + rc.effective_pct
        return idx

    def written_index(self, k: int) -> float:
        """Cohort-month written index with the two-point pro-rata blend (D4).

        No change inside the month: the full cumulative index. Otherwise the
        month's written weight is split at the FIRST in-month change date:
        weight p = (days on/after first change) / (days in month) at the
        end-of-month index, (1-p) at the pre-month index.
        """
        som, eom, dim = mi_first(k), mi_last(k), mi_days(k)
        w_pre = self.index_at(som - dt.timedelta(days=1))
        w_post = self.index_at(eom)
        in_month = [rc.effective for rc in self.changes if som <= rc.effective <= eom]
        if not in_month:
            return w_post
        first = min(in_month)
        p = ((eom - first).days + 1) / dim
        return (1.0 - p) * w_pre + p * w_post

    def written_index_exact_daily(self, k: int) -> float:
        """Exact daily-average index for the month (test cross-check only)."""
        som, dim = mi_first(k), mi_days(k)
        return sum(self.index_at(som + dt.timedelta(days=d)) for d in range(dim)) / dim

    # -- earning profile ----------------------------------------------------
    def e(self, k: int, m: int) -> float:
        """Fraction of cohort ``k`` earned in month ``m`` (mid-month convention)."""
        off = m - k
        t = self.t
        if off < 0 or off > t:
            return 0.0
        if off == 0 or off == t:
            return 1.0 / (2 * t)
        return 1.0 / t

    def ec(self, k: int, year: int) -> float:
        """Fraction of cohort ``k`` earned in calendar ``year`` (closed form, D10).

        Equals sum of e(k, m) over months m of the year — proved in tests.
        """
        t = self.t

        def cum(o: int) -> float:  # cumulative earned through month offset o
            return min(1.0, max(0.0, (o + 0.5) / t))

        y0, y1 = month_index(year, 1), month_index(year, 12)
        return cum(y1 - k) - cum(y0 - 1 - k)

    # -- written mod path ---------------------------------------------------
    def written_mod(self, k: int) -> float:
        return self.mod_path.value(mi_mid_ordinal(k))

    # -- net rate selection (D39) -------------------------------------------
    def combined_hist(self, k: int) -> float:
        """Historical combined net price index q(k) = W x (M_w / M_ind).

        The mod leg is included only when the combo's mod adjustment is in
        force; otherwise the net target is interpreted rate-only (advisory-
        flagged by validation).
        """
        q = self.written_index(k)
        if self.combo.mod_adjustment_enabled:
            q *= self.written_mod(k) / self.mods.m_ind
        return q

    def net_index(self, k: int) -> float:
        """Combined index under the net selection: modeled history before
        1/1/P; q(k) = q(k-12) x (1 + x) for cohorts written on/after 1/1/P."""
        jan_p = month_index(self.p, 1)
        jan_p1 = month_index(self.p + 1, 1)
        if k < jan_p:
            return self.combined_hist(k)
        x = self.combo.net_sel_p if k < jan_p1 else self.combo.net_x1
        return self.net_index(k - 12) * (1.0 + x)

    def index_in_force(self, k: int) -> float:
        """The index the calendar year actually earns: the combined net path
        in net mode, the explicit written rate index otherwise."""
        return self.net_index(k) if self.combo.net_mode else self.written_index(k)

    # -- aggregates ---------------------------------------------------------
    def _cy_weighted(self, year: int, values) -> float:
        num = den = 0.0
        for k in self.cohorts:
            wk = self.w(k)
            ekc = self.ec(k, year)
            if wk == 0.0 or ekc == 0.0:
                continue
            den += wk * ekc
            num += wk * ekc * values(k)
        if den == 0.0:
            raise ValueError(f"No earned exposure in CY {year} (check seasonality weights).")
        return num / den

    def earned_index_cy(self, year: int) -> float:
        return self._cy_weighted(year, self.index_in_force)

    def earned_mod_cy(self, year: int) -> float:
        return self._cy_weighted(year, self.written_mod)

    def _month_weighted(self, m: int, values) -> float:
        num = den = 0.0
        for k in range(max(self.c0, m - self.t), min(self.c1, m) + 1):
            wk_e = self.w(k) * self.e(k, m)
            den += wk_e
            num += wk_e * values(k)
        return num / den if den else float("nan")

    def earned_index_month(self, m: int) -> float:
        return self._month_weighted(m, self.index_in_force)

    def earned_mod_month(self, m: int) -> float:
        return self._month_weighted(m, self.written_mod)


# ---------------------------------------------------------------------------
# Continuous (closed-form parallelogram) engine
# ---------------------------------------------------------------------------


def _cy_overlap_weight(t: float, y0: float, y1: float, tau: float) -> float:
    """Earned overlap of a policy written at time t with [y0, y1], / term."""
    return max(0.0, min(t + tau, y1) - max(t, y0)) / tau


def _breakpoints(y0: float, y1: float, tau: float, extra: Iterable[float]) -> list[float]:
    pts = {y0 - tau, y0, y1 - tau, y1}
    pts.update(x for x in extra if y0 - tau < x < y1)
    return sorted(pts)


def continuous_earned_index(
    changes: Sequence[RateChange], year: int, term_months: int = 12
) -> float:
    """Closed-form parallelogram CY earned rate index, time in month units (D3).

    Uniform continuous writings; policies written at t earn uniformly over
    [t, t + term]. Exact piecewise integration (trapezoid on linear pieces).
    """
    tau = float(term_months)
    y0, y1 = float(month_index(year, 1)), float(month_index(year + 1, 1))
    coords = sorted(month_coord(rc.effective) for rc in changes)
    pts = _breakpoints(y0, y1, tau, coords)

    def index_at(t: float) -> float:
        idx = 1.0
        for rc in changes:
            if month_coord(rc.effective) <= t:
                idx *= 1.0 + rc.effective_pct
        return idx

    num = den = 0.0
    for a, b in zip(pts, pts[1:]):
        la = _cy_overlap_weight(a, y0, y1, tau)
        lb = _cy_overlap_weight(b, y0, y1, tau)
        area = (la + lb) / 2.0 * (b - a)  # exact: L is linear between breakpoints
        den += area
        num += area * index_at((a + b) / 2.0)  # index constant between breakpoints
    return num / den


def continuous_earned_mod(
    plan_year: int, mods: ModInputs, year: int, term_months: int = 12
) -> float:
    """Closed-form CY earned mod: exposure-weighted average of the linear path.

    Integrand L(t)·M(t) is quadratic between breakpoints -> Simpson is exact.
    """
    tau = float(term_months)
    y0, y1 = float(month_index(year, 1)), float(month_index(year + 1, 1))
    path = ModPath(plan_year, mods, coord=lambda d: month_coord(d, end_of_day=True))
    pts = _breakpoints(y0, y1, tau, (x for x, _ in path.anchors))

    num = den = 0.0
    for a, b in zip(pts, pts[1:]):
        h = b - a

        def f(t: float) -> float:
            return _cy_overlap_weight(t, y0, y1, tau)

        def g(t: float) -> float:
            return f(t) * path.value(t)

        mid = (a + b) / 2.0
        den += h / 6.0 * (f(a) + 4.0 * f(mid) + f(b))
        num += h / 6.0 * (g(a) + 4.0 * g(mid) + g(b))
    return num / den


def continuous_net_index(plan_year: int, combo: ComboInputs, year: int) -> float:
    """Closed-form CY earned COMBINED net price index under a net selection.

    q(t) = W(t) x M(t)/M_ind for t before 1/1/P (modeled history; the mod leg
    drops out when the combo's mod adjustment is off), and
    q(t) = q(t - 12) x (1 + x) thereafter (D39). Breakpoints include the
    historical kinks shifted forward 12 and 24 months so q is linear on every
    integration segment; Simpson is then exact for L(t) x q(t).
    """
    tau = float(combo.term_months)
    y0, y1 = float(month_index(year, 1)), float(month_index(year + 1, 1))
    jan_p = float(month_index(plan_year, 1))
    path = ModPath(plan_year, combo.mods, coord=lambda d: month_coord(d, end_of_day=True))
    changes = [(month_coord(rc.effective), rc.effective_pct) for rc in combo.rate_changes]

    def w_at(t: float) -> float:
        idx = 1.0
        for c, pct in changes:
            if c <= t:
                idx *= 1.0 + pct
        return idx

    def q_seg(t: float, ref: float) -> float:
        """q(t) with the (jumpy) rate leg sampled at ``ref`` — a point in the
        same continuity interval as t — so segment-endpoint evaluation never
        straddles a rate-change jump. The mod leg is continuous, so it is
        evaluated at t itself."""
        if t < jan_p:
            q = w_at(ref)
            if combo.mod_adjustment_enabled:
                q *= path.value(t) / combo.mods.m_ind
            return q
        x = combo.net_sel_p if t < jan_p + 12.0 else combo.net_x1
        return q_seg(t - 12.0, ref - 12.0) * (1.0 + x)

    base_bps = {c for c, _ in changes} | {x for x, _ in path.anchors} | {jan_p}
    shifted = set()
    for shift in (0.0, 12.0, 24.0):
        shifted |= {b + shift for b in base_bps}
    pts = _breakpoints(y0, y1, tau, shifted)

    num = den = 0.0
    for a, b in zip(pts, pts[1:]):
        h = b - a
        mid = (a + b) / 2.0

        def f(t: float) -> float:
            return _cy_overlap_weight(t, y0, y1, tau)

        den += h / 6.0 * (f(a) + 4.0 * f(mid) + f(b))
        num += h / 6.0 * (f(a) * q_seg(a, mid) + 4.0 * f(mid) * q_seg(mid, mid)
                          + f(b) * q_seg(b, mid))
    return num / den


# ---------------------------------------------------------------------------
# Top-level bridge
# ---------------------------------------------------------------------------


def crl_indication(changes: Sequence[RateChange]) -> float:
    """Cumulative rate level in the indication: product over considered rows."""
    crl = 1.0
    for rc in changes:
        if rc.considered:
            crl *= 1.0 + rc.effective_pct
    return crl


def lr_at_current_level(combo: ComboInputs) -> float:
    """Basis normalization (brief §3.4.1)."""
    if combo.lr_basis == "proposed":
        return combo.lr_proj * (1.0 + combo.sel_change)
    return combo.lr_proj


def run_bridge(plan_year: int, combo: ComboInputs, convention: str = "monthly") -> EngineResult:
    """Compute the full bridge for one combo under either convention."""
    warnings = validate_inputs(plan_year, combo)
    p = plan_year
    crl = crl_indication(combo.rate_changes)
    lr_cur = lr_at_current_level(combo)

    if convention == "monthly":
        eng = MonthlyEngine(p, combo)
        e_cy = {y: eng.earned_index_cy(y) for y in (p - 1, p, p + 1)}
        m_earned = {y: eng.earned_mod_cy(y) for y in (p - 1, p, p + 1)}
        earned_index_m = [eng.earned_index_month(m) for m in range(eng.r0, eng.r1 + 1)]
        n_gap = sum(1 for v in earned_index_m if math.isnan(v))
        if n_gap:
            warnings.append(
                f"{n_gap} reported months have zero earned exposure (short policy term "
                "combined with concentrated seasonality); monthly series show gaps there. "
                "Calendar-year aggregates are unaffected."
            )
        series = dict(
            cohort_mi=list(eng.cohorts),
            inforce_index=[eng.index_in_force(k) for k in eng.cohorts],
            written_index=[eng.written_index(k) for k in eng.cohorts],
            written_mod=[eng.written_mod(k) for k in eng.cohorts],
            written_weight=[eng.w(k) for k in eng.cohorts],
            earned_month_mi=list(range(eng.r0, eng.r1 + 1)),
            earned_index_m=earned_index_m,
            earned_mod_m=[eng.earned_mod_month(m) for m in range(eng.r0, eng.r1 + 1)],
        )
    elif convention == "continuous":
        if combo.seasonality is not None:
            warnings.append("Continuous cross-check assumes uniform writings; seasonality ignored.")
        if combo.net_mode:
            e_cy = {y: continuous_net_index(p, combo, y) for y in (p - 1, p, p + 1)}
        else:
            e_cy = {
                y: continuous_earned_index(combo.rate_changes, y, combo.term_months)
                for y in (p - 1, p, p + 1)
            }
        m_earned = {
            y: continuous_earned_mod(p, combo.mods, y, combo.term_months)
            for y in (p - 1, p, p + 1)
        }
        series = {}
    else:
        raise ValueError(f"convention must be 'monthly' or 'continuous', got {convention!r}")

    a_rate_p = crl / e_cy[p]
    a_rate_p1 = crl / e_cy[p + 1]
    if combo.net_mode:
        # Rate and price merge into the single combined factor A_net = CRL/E[q]
        # carried on the a_rate slot; A_mod is 1 by construction (D39).
        a_mod_p = a_mod_p1 = 1.0
    elif combo.mod_adjustment_enabled:
        a_mod_p = combo.mods.m_ind / m_earned[p]
        a_mod_p1 = combo.mods.m_ind / m_earned[p + 1]
    else:
        a_mod_p = a_mod_p1 = 1.0
    cy_lr_p = lr_cur * a_rate_p * a_mod_p * combo.a_other
    cy_lr_p1 = lr_cur * (1.0 + combo.net_trend) * a_rate_p1 * a_mod_p1 * combo.a_other

    return EngineResult(
        plan_year=p,
        convention=convention,
        crl_ind=crl,
        e_cy=e_cy,
        m_earned=m_earned,
        a_rate_p=a_rate_p,
        a_rate_p1=a_rate_p1,
        a_mod_p=a_mod_p,
        a_mod_p1=a_mod_p1,
        lr_current=lr_cur,
        cy_lr_p=cy_lr_p,
        cy_lr_p1=cy_lr_p1,
        earned_chg_vs_ind=e_cy[p] / crl - 1.0,
        yoy_earned_p=e_cy[p] / e_cy[p - 1] - 1.0,
        yoy_earned_p1=e_cy[p + 1] / e_cy[p] - 1.0,
        warnings=warnings,
        **series,
    )


# ---------------------------------------------------------------------------
# Enhancement E1: inverse plan solver (closed form)
# ---------------------------------------------------------------------------


@dataclass
class SolverResult:
    required_change: float  # r, the effective (achieved) change
    filed_equivalent: float  # r / achievement
    c_pre: float
    c_post: float
    denominator: float
    a_mod_p: float
    post_share: float  # c_post / (c_pre + c_post): earnable share of CY P
    warnings: list


def _solver_components(
    plan_year: int, combo: ComboInputs, eff_date: dt.date
) -> tuple[float, float, float, MonthlyEngine]:
    """C_pre, C_post, D over CY P with the planned program excluded (D13).

    W0 is the taken-rows-only index. The forward engine (written_index, D4)
    blends a cohort month at the FIRST in-month change date, valuing the
    post-split weight at the fully-compounded end-of-month index. With the
    unknown change r added at ``eff_date``, the effective-month cohort becomes

        (1-p)*W0_pre  +  p * W0_eom * (1+r),
        p = days on/after min(eff_date, first taken change in month) / days,

    which stays exactly linear in r. C_pre/C_post carry those coefficients so
    the closed form inverts the engine exactly — including when a taken change
    shares the effective month (adversarial-review fix; DECISIONS.md D31).
    """
    base = tuple(rc for rc in combo.rate_changes if rc.status == STATUS_TAKEN)
    eng = MonthlyEngine(plan_year, combo, changes_override=base)
    k_eff = month_index(eff_date.year, eff_date.month)
    som, eom, dim = mi_first(k_eff), mi_last(k_eff), mi_days(k_eff)
    in_month = [rc.effective for rc in base if som <= rc.effective <= eom]
    first = min([eff_date] + in_month)
    p_split = ((eom - first).days + 1) / dim
    w0_pre = eng.index_at(som - dt.timedelta(days=1))
    w0_eom = eng.index_at(eom)
    c_pre = c_post = den = 0.0
    for k in eng.cohorts:
        wk_ec = eng.w(k) * eng.ec(k, plan_year)
        if wk_ec == 0.0:
            continue
        den += wk_ec
        if k < k_eff:
            c_pre += wk_ec * eng.written_index(k)
        elif k > k_eff:
            c_post += wk_ec * eng.written_index(k)
        else:
            c_pre += wk_ec * (1.0 - p_split) * w0_pre
            c_post += wk_ec * p_split * w0_eom
    return c_pre, c_post, den, eng


def solve_rate_for_target(
    plan_year: int,
    combo: ComboInputs,
    target_cy_lr: float,
    eff_date: dt.date,
    achievement: float = 1.0,
    max_reasonable_rate: float = 0.25,
    min_post_share: float = 0.05,
) -> SolverResult:
    """Mode A: required planned change r at ``eff_date`` to hit ``target_cy_lr``.

    E_CY(P) is linear in r: E(r) = (C_pre + (1+r)·C_post) / D, so
    r = (E_needed·D − C_pre) / C_post − 1 with E_needed = CRL_ind / A_rate_needed.
    """
    c_pre, c_post, den, _ = _solver_components(plan_year, combo, eff_date)
    res = run_bridge(plan_year, combo)  # for A_mod, CRL, LR_current
    a_needed = target_cy_lr / (res.lr_current * res.a_mod_p * combo.a_other)
    e_needed = res.crl_ind / a_needed
    warnings: list[str] = []
    post_share = c_post / (c_pre + c_post)
    if post_share < min_post_share:
        warnings.append(
            f"Only {post_share:.1%} of CY {plan_year} earned exposure sits on/after "
            f"{eff_date}; late-year dates make targets mathematically absurd."
        )
    if c_post == 0.0:
        warnings.append(
            f"No CY {plan_year} earned exposure sits on/after {eff_date}; "
            "no rate change at this date can move the plan-year result."
        )
        r = float("nan")
    else:
        r = (e_needed * den - c_pre) / c_post - 1.0
        if abs(r) > max_reasonable_rate:
            warnings.append(
                f"Required change {r:+.1%} exceeds the ±{max_reasonable_rate:.0%} reasonability bound."
            )
    return SolverResult(
        required_change=r,
        filed_equivalent=r / achievement if achievement else float("nan"),
        c_pre=c_pre,
        c_post=c_post,
        denominator=den,
        a_mod_p=res.a_mod_p,
        post_share=post_share,
        warnings=warnings,
    )


def solver_timing_table(
    plan_year: int, combo: ComboInputs, r: float, target_cy_lr: float | None = None
) -> list[dict]:
    """Mode B: CY LR for the change r effective the 1st of each month of P."""
    res = run_bridge(plan_year, combo)
    rows = []
    for m in range(1, 13):
        eff = dt.date(plan_year, m, 1)
        c_pre, c_post, den, _ = _solver_components(plan_year, combo, eff)
        e = (c_pre + (1.0 + r) * c_post) / den
        lr = res.lr_current * (res.crl_ind / e) * res.a_mod_p * combo.a_other
        rows.append(
            dict(
                month=m,
                effective=eff,
                earned_index=e,
                cy_lr=lr,
                meets_target=(target_cy_lr is not None and lr <= target_cy_lr),
            )
        )
    return rows


# ---------------------------------------------------------------------------
# Net delivery decomposition (DECISIONS.md D57)
#
# Given a net rate selection x from 1/1/P and ONE new rate change r at a
# state-chosen effective date D inside the plan year, decompose the target's
# month-by-month delivery on renewals:
#
#   delivered(m) = q(m)/q(m-12) = [W(m)/W(m-12)] x [M_w(m)/M_w(m-12)]
#
# The rate leg's base is the LIVE rows — taken rows plus planned rows
# effective before 1/1/P (D39: only planned rows on/after 1/1/P are
# superseded; this is NOT the Solver's taken-only D13 convention). The new
# change enters affinely per month, W_new(m) = A(m) + B(m)·r, with the D31
# first-in-month day-blend, so both the suggested filed change and the
# required year-end mod solve in closed form.
# ---------------------------------------------------------------------------


def _live_changes(plan_year: int, combo: ComboInputs) -> tuple[RateChange, ...]:
    """Rows that stay live under a net selection (taken + planned < 1/1/P)."""
    jan1 = dt.date(plan_year, 1, 1)
    return tuple(rc for rc in combo.rate_changes
                 if rc.status == STATUS_TAKEN or rc.effective < jan1)


@dataclass
class NetDeliveryComponents:
    plan_year: int
    eff_date: dt.date
    x: float                       # net selection for P
    mod_on: bool                   # mod adjustment in force (else rate-only)
    months: list                   # 12 absolute month indices, Jan..Dec P
    w: list                        # written weights w(m)
    a: list                        # W_new(m) = a(m) + b(m) * r
    b: list
    w12: list                      # year-ago written index W(m-12), live rows
    mw12: list                     # year-ago written mod M_w(m-12)
    mproj: list                    # projected written mod at m (anchor path)
    p_blend: float                 # day-blend share in D's month (D31 rule)
    c0: float                      # sum kappa(m) * a(m)
    c1: float                      # sum kappa(m) * b(m)
    sum_w: float
    warnings: list = field(default_factory=list)


def net_delivery_components(
    plan_year: int, combo: ComboInputs, eff_date: dt.date
) -> NetDeliveryComponents:
    """Per-month affine components of delivered net under one new change at D."""
    if combo.net_sel_p is None:
        raise ValueError("net_delivery_components requires a net rate selection.")
    if not dt.date(plan_year, 1, 1) <= eff_date <= dt.date(plan_year, 12, 31):
        raise ValueError(
            f"Effective date {eff_date} must lie inside plan year {plan_year}: a "
            "prior-year change also enters the year-ago comparison base and the "
            "closed form stops being exact (enter it in the rate log instead)."
        )
    live = _live_changes(plan_year, combo)
    eng = MonthlyEngine(plan_year, combo, changes_override=live)
    m_d = month_index(eff_date.year, eff_date.month)
    mod_on = combo.mod_adjustment_enabled

    months, w, a, b, w12, mw12, mproj = [], [], [], [], [], [], []
    p_blend = 1.0
    c0 = c1 = sum_w = 0.0
    for j in range(12):
        mi = month_index(plan_year, 1) + j
        months.append(mi)
        wm = eng.w(mi)
        w.append(wm)
        w12_m = eng.written_index(mi - 12)
        w12.append(w12_m)
        mw12_m = eng.written_mod(mi - 12)
        mw12.append(mw12_m)
        mproj_m = eng.written_mod(mi)
        mproj.append(mproj_m)
        if mi < m_d:
            am, bm = eng.written_index(mi), 0.0
        elif mi > m_d:
            base = eng.written_index(mi)
            am, bm = base, base
        else:
            # month of D: the blend splits at the FIRST live in-month change
            # OR D, whichever is earlier (D31); the end-of-month leg carries r
            som, eom, dim = mi_first(mi), mi_last(mi), mi_days(mi)
            w_pre = eng.index_at(som - dt.timedelta(days=1))
            w_eom = eng.index_at(eom)
            in_month = [rc.effective for rc in live if som <= rc.effective <= eom]
            first = min(in_month + [eff_date])
            p_blend = ((eom - first).days + 1) / dim
            am = (1.0 - p_blend) * w_pre + p_blend * w_eom
            bm = p_blend * w_eom
        a.append(am)
        b.append(bm)
        ratio = (mproj_m / mw12_m) if mod_on else 1.0
        kappa = wm * ratio / w12_m
        c0 += kappa * am
        c1 += kappa * bm
        sum_w += wm
    return NetDeliveryComponents(
        plan_year=plan_year, eff_date=eff_date, x=combo.net_sel_p, mod_on=mod_on,
        months=months, w=w, a=a, b=b, w12=w12, mw12=mw12, mproj=mproj,
        p_blend=p_blend, c0=c0, c1=c1, sum_w=sum_w,
        warnings=[] if mod_on else [
            "Mod adjustment OFF: target interpreted rate-only; the entire "
            "delivery burden sits on the filing (D39)."],
    )


def suggest_net_rate(
    plan_year: int, combo: ComboInputs, eff_date: dt.date
) -> tuple[float, NetDeliveryComponents]:
    """Closed-form filed change at D so the written-weighted delivered net,
    with mods on the ALREADY-PROJECTED path, equals the target:

        sum w(m)·[A(m)+B(m)·r]·mu(m)/W(m-12) = (1+x)·sum w(m)
        =>  r* = ((1+x)·sum_w - C0) / C1
    """
    comp = net_delivery_components(plan_year, combo, eff_date)
    if comp.c1 == 0.0:
        raise ValueError(
            "No plan-year written exposure on/after the effective date — the "
            "filing cannot influence delivery (C1 = 0)."
        )
    r = ((1.0 + comp.x) * comp.sum_w - comp.c0) / comp.c1
    return r, comp


def required_m1(
    plan_year: int, combo: ComboInputs, eff_date: dt.date, r: float
) -> float:
    """Dual closed form: the year-end written mod M_1' that delivers the
    target given the filed change r. Plan-month mods are affine in M_1 along
    the anchor path, M(m) = alpha(m) + beta(m)·M_1, so M_1' solves in one sum.
    Requires the mod adjustment in force (rate-only mode has no mod ask)."""
    comp = net_delivery_components(plan_year, combo, eff_date)
    if not comp.mod_on:
        raise ValueError("required_m1 is undefined when the mod adjustment is off.")
    mods = combo.mods
    x_asof = float(mods.m0_asof.toordinal() + 1)               # close-of-day (D2)
    x_end = float(dt.date(plan_year, 12, 31).toordinal() + 1)
    x_prior = float(add_months(mods.m0_asof, -12).toordinal() + 1)
    num = (1.0 + comp.x) * comp.sum_w
    den = 0.0
    for j in range(12):
        xm = mi_mid_ordinal(comp.months[j])
        if mods.m_prior is not None and xm <= x_asof:
            # explicit prior segment: independent of M_1
            beta = 0.0
            alpha = mods.m_prior + (mods.m0 - mods.m_prior) * (xm - x_prior) / (x_asof - x_prior)
        else:
            # on (or extrapolated from) the M_0 -> M_1 segment (§3.3.1)
            beta = (xm - x_asof) / (x_end - x_asof)
            alpha = mods.m0 * (1.0 - beta)
        h = comp.w[j] * (comp.a[j] + comp.b[j] * r) / (comp.w12[j] * comp.mw12[j])
        num -= h * alpha
        den += h * beta
    if den == 0.0:
        raise ValueError("Mod path has no M_1 leverage over the plan year (den = 0).")
    return num / den


def net_delivery_by_month(
    plan_year: int, combo: ComboInputs, eff_date: dt.date, r: float
) -> list[dict]:
    """The 12-row delivery decomposition at a given filed change r."""
    comp = net_delivery_components(plan_year, combo, eff_date)
    rows = []
    for j in range(12):
        w_new = comp.a[j] + comp.b[j] * r
        rate_leg = w_new / comp.w12[j] - 1.0
        price_req = (1.0 + comp.x) / (1.0 + rate_leg) - 1.0 if comp.mod_on else None
        m_req = (comp.mw12[j] * (1.0 + comp.x) / (1.0 + rate_leg)
                 if comp.mod_on else None)
        ratio = (comp.mproj[j] / comp.mw12[j]) if comp.mod_on else 1.0
        rows.append(dict(
            mi=comp.months[j], w=comp.w[j], rate_leg=rate_leg,
            price_leg_required=price_req, m_required=m_req, m_proj=comp.mproj[j],
            delivered_at_proj=(1.0 + rate_leg) * ratio - 1.0,
        ))
    return rows


def net_program_plan_lr(
    plan_year: int, combo: ComboInputs, eff_date: dt.date, r: float,
    m1_prime: float | None,
) -> float:
    """CY P plan LR if the concrete program is BOOKED instead of asserting the
    net path: live history earns as modeled; plan-year cohorts carry
    W_new(m) x M'(m)/M_ind with M' on the (M_0 -> M_1') anchor line (or the
    rate leg alone when the mod adjustment is off). Comparable to the
    net-mode plan LR (combined factor, A_mod = 1)."""
    comp = net_delivery_components(plan_year, combo, eff_date)
    live = _live_changes(plan_year, combo)
    eng = MonthlyEngine(plan_year, combo, changes_override=live)
    mods = combo.mods
    x_asof = float(mods.m0_asof.toordinal() + 1)
    x_end = float(dt.date(plan_year, 12, 31).toordinal() + 1)
    jan_p = month_index(plan_year, 1)

    def q_prog(k: int) -> float:
        if k < jan_p:
            return eng.combined_hist(k)
        j = k - jan_p
        w_new = comp.a[j] + comp.b[j] * r
        if not comp.mod_on:
            return w_new
        beta = (mi_mid_ordinal(k) - x_asof) / (x_end - x_asof)
        m_prime = mods.m0 * (1.0 - beta) + beta * float(m1_prime)
        if mods.m_prior is not None and mi_mid_ordinal(k) <= x_asof:
            m_prime = eng.written_mod(k)   # explicit prior segment unchanged
        return w_new * m_prime / mods.m_ind

    num = den = 0.0
    for k in eng.cohorts:
        wk, ekc = eng.w(k), eng.ec(k, plan_year)
        if wk == 0.0 or ekc == 0.0:
            continue
        den += wk * ekc
        num += wk * ekc * q_prog(k)
    e_prog = num / den
    return (lr_at_current_level(combo) * crl_indication(combo.rate_changes)
            / e_prog * combo.a_other)


# ---------------------------------------------------------------------------
# Enhancement E2: plan-vs-actual attribution
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PlannedActual:
    """Actuals for one planned rate-log row (matched by planned effective date)."""

    planned_effective: dt.date
    actual_pct: float  # actual achieved total change (already achievement-adjusted)
    actual_effective: dt.date


@dataclass
class AttributionResult:
    steps: list  # list of (label, cy_lr_after_step)
    factors: list  # list of (label, multiplicative factor)
    points: list  # list of (label, step change in LR points)
    reconciles: bool


def attribution(
    plan_year: int,
    combo: ComboInputs,
    actuals: Sequence[PlannedActual],
    mods_actual: ModInputs,
    actual_cy_lr: float,
) -> AttributionResult:
    """Sequential multiplicative decomposition of (actual − plan) CY LR (§7 E2).

    Order (documented on-sheet): 1) rate magnitude at planned dates,
    2) rate timing at actual dates, 3) mod drift, 4) loss-side residual.
    Taken rows are assumed booked as logged. The mod-drift factor uses the
    PLAN M_ind (an indication property, not an actual) against the actual
    earned-mod path (adversarial-review hardening; DECISIONS.md D31).
    """
    planned_effs = {rc.effective for rc in combo.rate_changes if rc.status == STATUS_PLANNED}
    for a in actuals:
        if a.planned_effective not in planned_effs:
            raise ValueError(
                f"Attribution actual references planned effective {a.planned_effective}, "
                f"which matches no planned rate-log row {sorted(planned_effs)}."
            )
    by_eff = {a.planned_effective: a for a in actuals}

    def swap(rc: RateChange, use_actual_date: bool) -> RateChange:
        a = by_eff.get(rc.effective)
        if rc.status != STATUS_PLANNED or a is None:
            return rc
        return replace(
            rc,
            filed_pct=a.actual_pct,
            achievement=1.0,
            effective=a.actual_effective if use_actual_date else rc.effective,
        )

    l_mag = tuple(swap(rc, use_actual_date=False) for rc in combo.rate_changes)
    l_time = tuple(swap(rc, use_actual_date=True) for rc in combo.rate_changes)

    base = run_bridge(plan_year, combo)
    p = plan_year
    e_mag = MonthlyEngine(p, combo, changes_override=l_mag).earned_index_cy(p)
    eng_time = MonthlyEngine(p, combo, changes_override=l_time)
    e_time = eng_time.earned_index_cy(p)
    m_act = MonthlyEngine(p, combo, mods_override=mods_actual).earned_mod_cy(p)
    a_mod_act = (combo.mods.m_ind / m_act) if combo.mod_adjustment_enabled else 1.0

    lr0 = base.cy_lr_p
    lr1 = base.lr_current * (base.crl_ind / e_mag) * base.a_mod_p * combo.a_other
    lr2 = base.lr_current * (base.crl_ind / e_time) * base.a_mod_p * combo.a_other
    lr3 = base.lr_current * (base.crl_ind / e_time) * a_mod_act * combo.a_other
    lr4 = actual_cy_lr

    labels = ["Plan CY LR", "Rate magnitude", "Rate timing", "Mod drift", "Loss-side residual"]
    lrs = [lr0, lr1, lr2, lr3, lr4]
    factors = [(labels[i], lrs[i] / lrs[i - 1]) for i in range(1, 5)]
    points = [(labels[i], (lrs[i] - lrs[i - 1]) * 100.0) for i in range(1, 5)]
    product = lr0
    for _, f in factors:
        product *= f
    return AttributionResult(
        steps=list(zip(labels, lrs)),
        factors=factors,
        points=points,
        reconciles=math.isclose(product, actual_cy_lr, rel_tol=1e-12, abs_tol=1e-12),
    )


# ---------------------------------------------------------------------------
# The §9 worked example
# ---------------------------------------------------------------------------


def worked_example() -> tuple[int, ComboInputs]:
    """Brief §9: P=2027, indication eff 7/1/2026, LR 65% current basis."""
    combo = ComboInputs(
        lr_proj=0.65,
        lr_basis="current",
        sel_change=0.0,
        mods=ModInputs(
            m_ind=0.850,
            m0=0.860,
            m0_asof=dt.date(2026, 9, 30),
            m1=0.890,
        ),
        rate_changes=(
            RateChange(dt.date(2025, 7, 1), 0.04, STATUS_TAKEN, considered=True),
            RateChange(dt.date(2026, 7, 1), 0.06, STATUS_TAKEN, considered=True),
            RateChange(dt.date(2027, 4, 1), 0.05, STATUS_PLANNED, considered=False, achievement=1.0),
        ),
        net_trend=0.0,
        a_other=1.0,
        mod_adjustment_enabled=True,
        term_months=12,
        seasonality=None,
    )
    return 2027, combo
