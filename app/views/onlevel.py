"""The on-leveling rectangle — policy year converted to calendar year,
drawn (W2d, extended W3b).

John's ask, verbatim: "a rectangular area for plan year and plan year
plus 1 with the earned rate levels identified in each area — really just
a visualization of the whole reason this app exists: conversion from
policy year to calendar year." This is the textbook parallelogram
diagram, live for the selected combo:

- x: calendar time across CY P-1, CY P and CY P+1 (months since Jan P);
- y: fraction of the policy term elapsed — a policy written at t earns
  along the diagonal (t, 0) → (t + term, 1);
- every rate change draws its diagonal; the regions between diagonals
  carry the rate level, shaded and labeled per calendar-year strip;
- each calendar year is annotated with the PRODUCTION earned level (and
  A_rate where the engine publishes one — there is no A_rate for P-1).

NET COMBOS DRAW THE NET PATH (W3b — John's construction, verified): the
engine's net recursion q(k) = q(k-12) x (1 + x) on written cohorts IS a
diagonal set — a restart diagonal at each 1/1 (a sawtooth: the level
steps DOWN when the book out-earned the selection, and that is correct,
not an artifact) plus anniversary echoes of each filing in the trailing
year of history. Evaluator, jump set and kink set all come from
``engine.continuous_net_q`` — the app adds no math. Under a mod log the
level drifts INSIDE a band, so each per-year cell is labeled with q at
its writing-time centroid t = cx - term*cy; q is linear in t within a
band, so the exposure-weighted mean over the cell equals q at that
centroid and the region rebuild ties ``continuous_net_index`` exactly.

Geometry is the engine's CONTINUOUS cross-check convention (uniform
writings, seasonality not drawn) in the engine's own month-coordinate
space; the tie tests rebuild ``engine.continuous_earned_index`` (and
``continuous_net_index`` for net combos) from the region cells to 1e-9,
so the picture cannot drift from the math. When the production monthly
number differs visibly from the continuous one (seasonality, short
terms, the day-vs-month mod interpolation conventions), BOTH print —
disclosed, never hidden.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from src.engine import continuous_earned_index, continuous_net_q, \
    month_coord, month_index

from app.glue.theme import GREY_TEXT, NAVY, STEEL, STEEL_LIGHT, \
    WARN_AMBER
from . import ensure_hv


def _clip(poly: list, fn) -> list:
    """Sutherland–Hodgman: clip a convex polygon to fn(x, y) <= 0."""
    out: list = []
    n = len(poly)
    for i in range(n):
        a, b = poly[i], poly[(i + 1) % n]
        fa, fb = fn(*a), fn(*b)
        if fa <= 0:
            out.append(a)
            if fb > 0:
                t = fa / (fa - fb)
                out.append((a[0] + t * (b[0] - a[0]),
                            a[1] + t * (b[1] - a[1])))
        elif fb <= 0:
            t = fa / (fa - fb)
            out.append((a[0] + t * (b[0] - a[0]),
                        a[1] + t * (b[1] - a[1])))
    return out


def _area(poly: list) -> float:
    s = 0.0
    for i in range(len(poly)):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % len(poly)]
        s += x0 * y1 - x1 * y0
    return abs(s) / 2.0


def _centroid(poly: list) -> tuple:
    a = cx = cy = 0.0
    for i in range(len(poly)):
        x0, y0 = poly[i]
        x1, y1 = poly[(i + 1) % len(poly)]
        w = x0 * y1 - x1 * y0
        a += w
        cx += (x0 + x1) * w
        cy += (y0 + y1) * w
    if abs(a) < 1e-12:
        return poly[0]
    return cx / (3.0 * a), cy / (3.0 * a)


@dataclass
class OnlevelFrame:
    plan_year: int
    term: float
    net_mode: bool                      # True = drawing the NET path
    window: tuple                       # (wx0, 24.0) months-since-Jan-P
    regions: list = field(default_factory=list)
    diagonals: list = field(default_factory=list)
    years: list = field(default_factory=list)
    notes: list = field(default_factory=list)
    # W4c: the two indicated rate LEVELS, on the same index scale as the
    # bands (None when the row carries no target or projected LR)
    ind_rate: float | None = None       # rate-only indication
    ind_bridge: float | None = None     # the full bridge's crossing point


def indicated_levels(ci, res, target) -> tuple:
    """(rate-only, full-bridge) indicated rate LEVEL, or (None, None).

    This is the system's first computed indication — the workbook's
    indication fields are carry-through by design (D111), so nothing
    upstream defines one. Both readings share the same question ("what
    earned rate level prices this book to its target?") and differ in what
    they let help:

    - **rate-only** ``CRL_ind × LR_proj / target`` — the classic
      loss-ratio indication: rate alone does the work. LR_proj is quoted
      at CURRENT rate level (D107) and CRL_ind is that level on the
      exhibit's own 1.0-based index, so the product lands on the y-scale
      the bands live on.
    - **full bridge** multiplies in A_mod and A_other as well, marking
      where the PLAN LR crosses target given the mod path already
      projected. The gap between the two lines is exactly what the
      schedule mod and other adjustments are being asked to carry.

    Neither is a substitute for a filed indication; both are read off the
    inputs already in the book, and the exhibit says so. ``target`` is the
    row's target loss ratio — a reference input the engine never reads
    (D96), which is why it arrives as an argument rather than off ``ci``.
    """
    lr_proj = getattr(ci, "lr_proj", None)
    crl = getattr(res, "crl_ind", None)
    vals = (target, lr_proj, crl)
    if not all(isinstance(v, (int, float)) and not isinstance(v, bool)
               for v in vals):
        return None, None
    if target <= 0.0:
        return None, None
    rate_only = crl * lr_proj / target
    a_mod = getattr(res, "a_mod_p", 1.0) or 1.0
    a_other = getattr(ci, "a_other", 1.0) or 1.0
    return rate_only, rate_only * a_mod * a_other


def _year_cells(poly, strips, level_at):
    """Clip a band polygon into per-CY-strip cells; each cell carries its
    share of the strip's exposure (strip area is 12 for any term) and the
    level at its writing-time centroid."""
    cells = []
    for yr, xs in strips:
        yp = _clip(poly, lambda x, y, xs=xs: xs - x)
        yp = _clip(yp, lambda x, y, xe=xs + 12.0: x - xe)
        if len(yp) < 3 or _area(yp) < 1e-9:
            cells.append(dict(year=yr, share=0.0, level=None,
                              cx=None, cy=None))
            continue
        cx, cy = _centroid(yp)
        cells.append(dict(year=yr, share=_area(yp) / 12.0,
                          level=level_at(cx, cy), cx=cx, cy=cy))
    return cells


def onlevel_frame(ci, res, window: int = 36,
                  basis: str = "auto", target=None) -> OnlevelFrame:
    """Pure geometry: regions/diagonals/year annotations for one combo.
    x is measured in months since Jan of the plan year.

    ``window``: 36 (Jan P-1 .. Dec P+1, default) or 24 (Jan P .. Dec P+1).
    ``basis``: "auto" draws the net path for net combos; "program" draws
    the LOGGED rate program instead (the D65 program basis) — only
    meaningful for net combos, where the two geometries differ by exactly
    the program-vs-asserted story.

    ``target``: the row's target loss ratio. Given one, the frame carries
    the two indicated rate LEVELS (see ``indicated_levels``) so the
    picture can say not just what the book earned but what it needs.
    """
    p = res.plan_year
    tau = float(ci.term_months)
    base = float(month_index(p, 1))
    wx0 = -12.0 if window == 36 else 0.0
    wx1 = 24.0
    draw_net = bool(ci.net_mode) and basis != "program"
    fr = OnlevelFrame(plan_year=p, term=tau, net_mode=draw_net,
                      window=(wx0, wx1))
    strips = ([(p - 1, -12.0)] if wx0 <= -12.0 else []) \
        + [(p, 0.0), (p + 1, 12.0)]

    changes = sorted(((month_coord(rc.effective) - base, rc.effective_pct,
                       rc.effective) for rc in ci.rate_changes),
                     key=lambda t: t[0])
    rect = [(wx0, 0.0), (wx1, 0.0), (wx1, 1.0), (wx0, 1.0)]

    if draw_net:
        q, jumps, kinks = continuous_net_q(p, ci)
        # window coordinates; kinks only matter when the mod leg is live
        jset = {j - base for j in jumps}
        kset = {k - base for k in kinks} if ci.mod_adjustment_enabled \
            else set()

        def kindof(c):
            """Classify a bound; kinks lose to jumps on a coincidence."""
            if any(abs(c - j) < 1e-9 for j in jset):
                if abs(c) < 1e-9 or abs(c - 12.0) < 1e-9:
                    yr = p if abs(c) < 1e-9 else p + 1
                    x = ci.net_sel_p if abs(c) < 1e-9 else ci.net_x1
                    return "restart", (f"net restart 1/1/{yr} "
                                       f"({x:+.1%} vs year-ago)")
                for cc, _pct, d in changes:
                    if abs(c - cc) < 1e-9:
                        return "filing", str(d)
                    if abs(c - cc - 12.0) < 1e-9 or \
                            abs(c - cc - 24.0) < 1e-9:
                        return "echo", f"anniversary of {d} filing"
                return "filing", "rate change"
            return "kink", "mod-path anchor (drift)"

        lo_guard, hi_guard = wx0 - tau - 1.0, wx1 + 1.0
        bounds = sorted(c for c in jset | kset if lo_guard < c < hi_guard)
        merged: list = []
        for c in bounds:                       # dedupe coincident bounds
            if merged and abs(c - merged[-1]) < 1e-9:
                continue
            merged.append(c)
        bounds = [lo_guard] + merged + [hi_guard]
        openers = [("base", "base level")] + [kindof(c) for c in merged]

        def level_at(cx, cy):
            return q(base + (cx - tau * cy))

        superseded = sum(1 for c, _pct, _d in changes if c >= 0.0)
    else:
        bounds = [wx0 - tau - 1.0] + [c for c, _pct, _d in changes] \
            + [wx1 + 1.0]
        idx, levels = 1.0, [1.0]
        openers = [("base", "base level")]
        for _c, pct, d in changes:
            idx *= 1.0 + pct
            levels.append(idx)
            openers.append(("filing", str(d)))
        superseded = 0

    for k in range(len(bounds) - 1):
        lo, hi = bounds[k], bounds[k + 1]
        poly = _clip(rect, lambda x, y, lo=lo: lo - (x - tau * y))
        poly = _clip(poly, lambda x, y, hi=hi: (x - tau * y) - hi)
        if len(poly) < 3 or _area(poly) < 1e-9:
            continue
        if draw_net:
            cells = _year_cells(poly, strips, level_at)
        else:
            lv = levels[k]
            cells = _year_cells(poly, strips, lambda cx, cy, lv=lv: lv)
        tot = sum(c["share"] for c in cells)
        if tot <= 0.0:
            continue
        mean = sum(c["share"] * c["level"] for c in cells
                   if c["level"] is not None) / tot
        kind, opened = openers[k]
        by_year = {c["year"]: c for c in cells}
        fr.regions.append(dict(
            poly=poly, kind=kind, opened=opened, index=mean,
            label=f"{mean:.4f}", cells=cells,
            share_pm1=by_year.get(p - 1, {}).get("share", 0.0),
            share_p=by_year[p]["share"], share_p1=by_year[p + 1]["share"]))

    # ---- diagonals --------------------------------------------------------
    def add_diag(c, kind, date):
        y_lo = max(0.0, (wx0 - c) / tau)
        y_hi = min(1.0, (wx1 - c) / tau)
        if y_lo < y_hi:
            fr.diagonals.append(dict(
                x0=c + tau * y_lo, y0=y_lo, x1=c + tau * y_hi, y1=y_hi,
                kind=kind, date=date))

    if draw_net:
        seen: set = set()
        for c in sorted(jset):
            key = round(c, 9)
            if key in seen:
                continue
            seen.add(key)
            kind, opened = kindof(c)
            add_diag(c, kind, opened)
    else:
        for c, _pct, d in changes:
            add_diag(c, "filing", str(d))

    # ---- year annotations -------------------------------------------------
    net_cont = {}
    if draw_net:
        for yr, _xs in strips:
            net_cont[yr] = sum(
                c["share"] * c["level"] for r in fr.regions
                for c in r["cells"]
                if c["year"] == yr and c["level"] is not None)
    for yr, xs in strips:
        e_prod = res.e_cy[yr]
        a_prod = (res.a_rate_p if yr == p
                  else res.a_rate_p1 if yr == p + 1 else None)
        if draw_net:
            e_cont = net_cont.get(yr)
        else:
            e_cont = continuous_earned_index(ci.rate_changes, yr,
                                             ci.term_months)
        fr.years.append(dict(year=yr, x0=xs, x1=xs + 12.0,
                             e_cy=e_prod, a_rate=a_prod, e_cont=e_cont))
        if basis == "program" and ci.net_mode:
            continue                # divergence is the POINT of that view
        if e_cont is not None and abs(e_cont - e_prod) > 1e-3:
            fr.notes.append(
                f"CY {yr}: the diagram's uniform-writings geometry earns "
                f"{e_cont:.4f}; the production monthly convention earns "
                f"{e_prod:.4f} (seasonality / term / mod-interpolation "
                f"conventions) — the annotation shows the production "
                f"number.")

    # ---- notes ------------------------------------------------------------
    if draw_net:
        base_q = next((r["index"] for r in fr.regions
                       if r["kind"] == "base"), None)
        scale = (f" The base band sits near M/M_ind "
                 f"≈ {base_q:.3f}, not 1.000." if base_q is not None
                 and ci.mod_adjustment_enabled else "")
        fr.notes.append(
            "NET PATH: levels are the combined net price index "
            "q = rate × mod / M_ind — each 1/1 restarts the whole book at "
            "the selection above its year-ago level (a DOWN-step when the "
            "book out-earned the selection — correct, not an artifact), "
            "and each filing in the trailing year of history echoes once "
            "at its anniversary." + scale)
        if ci.mod_adjustment_enabled and ci.mod_changes:
            fr.notes.append(
                "Mod drift makes q slope INSIDE a band — each strip label "
                "is the exposure-weighted level of its cell (exact at the "
                "writing-time centroid).")
        if superseded:
            fr.notes.append(
                f"{superseded} logged filing(s) on/after 1/1/{p} are "
                f"superseded by the net selection and not drawn.")
    elif ci.net_mode:
        fr.notes.append(
            "PROGRAM BASIS (D65): drawing the LOGGED rate program that "
            "the net selection supersedes — the annotated earned levels "
            "remain the engine's net-path production values, so the gap "
            "between geometry and annotation IS the program-vs-asserted "
            "story.")

    # ---- what the book NEEDS, beside what it earned (W4c) -----------------
    fr.ind_rate, fr.ind_bridge = indicated_levels(ci, res, target)
    if fr.ind_rate is not None:
        fr.notes.append(
            f"INDICATED LEVEL (computed here, not carried): rate-only "
            f"{fr.ind_rate:.4f} = CRL_ind {res.crl_ind:.4f} × projected LR "
            f"{ci.lr_proj:.1%} / target {target:.1%} — the earned rate "
            f"level at which rate ALONE prices the book to target. The "
            f"full-bridge line {fr.ind_bridge:.4f} multiplies in A_mod and "
            f"A_other: where PLAN LR crosses target given the mod path "
            f"already projected. The gap between them is what the schedule "
            f"mod and other adjustments are being asked to carry. The "
            f"workbook's own indication fields are reference only (D111) — "
            f"nothing upstream defines this, so it is stated here.")
        same = abs(fr.ind_bridge - fr.ind_rate) < 1e-12
        fr.notes.append(
            "CRL_ind compounds the rate changes flagged IN the indication; "
            "the bands compound EVERY logged change — where those sets "
            "differ the indicated line and the band edges are on slightly "
            "different footings."
            + (" The two readings COINCIDE here: A_mod × A_other = 1.000, "
               "so nothing but rate is being asked to close the gap"
               + (" (net mode pins A_mod = 1 by construction — the mod leg "
                  "already rides inside the net target)."
                  if ci.net_mode else ".") if same else "")
            + (" On a net combo the bands carry the combined q (rate × mod "
               "/ M_ind) while these levels are rate-basis, so read the gap, "
               "not the crossing." if draw_net else ""))
    return fr


_DIAG_STYLE = {
    "filing": dict(color=NAVY, line_width=2, line_dash="solid"),
    "restart": dict(color=WARN_AMBER, line_width=2, line_dash="dashed"),
    "echo": dict(color=STEEL, line_width=1.6, line_dash="dotted"),
}


def onlevel(ci, res, frame: OnlevelFrame | None = None):
    """Render the frame: shaded index regions, kind-styled diagonals, year
    bounds, per-strip cell labels, and the per-year earned-level
    annotations. The P-1 strip is muted under a STEEL_LIGHT span."""
    hv = ensure_hv()
    from app.views.earning import _mix

    fr = frame if frame is not None else onlevel_frame(ci, res)
    p = fr.plan_year
    wx0, wx1 = fr.window
    lo = min((r["index"] for r in fr.regions), default=1.0)
    hi = max((r["index"] for r in fr.regions), default=1.0)

    def shade(ix):
        t = 0.15 if hi <= lo else 0.15 + 0.75 * (ix - lo) / (hi - lo)
        return _mix("#FFFFFF", STEEL, t)

    def cell_txt(r, yr):
        c = next((c for c in r["cells"] if c["year"] == yr), None)
        return "—" if c is None or c["level"] is None else f"{c['level']:.4f}"

    yrs = [y["year"] for y in fr.years]
    polys = hv.Polygons(
        [{("x", "y"): r["poly"], "level": r["index"],
          "Mean level": r["label"], "Opened by": r["opened"],
          **{f"Level in CY {yr}": cell_txt(r, yr) for yr in yrs},
          **{f"Share of CY {yr}":
             f"{next((c['share'] for c in r['cells'] if c['year'] == yr), 0.0):.1%}"
             for yr in yrs}}
         for r in fr.regions],
        vdims=["level", "Mean level", "Opened by"]
        + [f"Level in CY {yr}" for yr in yrs]
        + [f"Share of CY {yr}" for yr in yrs]).opts(
        color="level", cmap=[shade(r["index"]) for r in
                             sorted(fr.regions, key=lambda r: r["index"])]
        or ["#FFFFFF"],
        line_color="white", line_width=1.2, alpha=0.9, tools=["hover"],
        hover_tooltips=[("opened by", "@{Opened by}")]
        + [(f"level in CY {yr}", f"@{{Level in CY {yr}}}") for yr in yrs]
        + [(f"share of CY {yr}", f"@{{Share of CY {yr}}}") for yr in yrs],
        colorbar=False, show_legend=False)

    span = (hv.VSpan(wx0, 0.0).opts(color=STEEL_LIGHT, alpha=0.25)
            if wx0 < 0.0 else hv.Overlay([]))

    diags = hv.Overlay([
        hv.Curve([(d["x0"], d["y0"]), (d["x1"], d["y1"])]).opts(
            **_DIAG_STYLE.get(d["kind"], _DIAG_STYLE["filing"]))
        for d in fr.diagonals]) if fr.diagonals else hv.Overlay([])

    vlines = hv.Overlay([hv.VLine(x).opts(color=GREY_TEXT, line_width=1.4,
                                          line_dash="dotted")
                         for x in (wx0, 0.0, 12.0, 24.0) if x >= wx0])

    labels = hv.Labels(
        [(c["cx"], c["cy"], f"{c['level']:.4f}")
         for r in fr.regions for c in r["cells"]
         if c["share"] > 0.03 and c["level"] is not None],
        kdims=["x", "y"], vdims=["text"]).opts(
        text_font_size="9pt", text_color=NAVY, text_font_style="bold")

    def ann_txt(y):
        # W4c: the earned level next to what it would need to be. This
        # axis is TERM ELAPSED, so an indicated LEVEL cannot be a
        # position here — it is a comparison, and reads as one.
        gap = ("" if fr.ind_rate is None or not y["e_cy"]
               else f"   ({y['e_cy'] / fr.ind_rate - 1.0:+.1%} vs indicated)")
        if y["a_rate"] is None:
            return f"CY {y['year']}   earned level {y['e_cy']:.4f}{gap}"
        return (f"CY {y['year']}   earned level {y['e_cy']:.4f}   "
                f"A_rate {y['a_rate']:.4f}{gap}")

    ann = hv.Overlay([
        hv.Labels([(y["x0"] + 6.0, 1.07, ann_txt(y))],
                  kdims=["x", "y"], vdims=["text"]).opts(
            text_font_size="10pt",
            text_color=STEEL if y["a_rate"] is None else NAVY,
            text_font_style="bold")
        for y in fr.years])

    ind = hv.Overlay([])
    if fr.ind_rate is not None:
        ind = hv.Labels(
            [(wx0 + 0.4, 1.135,
              f"indicated rate level {fr.ind_rate:.4f} rate-only · "
              f"{fr.ind_bridge:.4f} full bridge  "
              f"(levels, not positions — this axis is term elapsed)")],
            kdims=["x", "y"], vdims=["text"]).opts(
            text_font_size="9pt", text_color=WARN_AMBER,
            text_font_style="bold", text_align="left")

    ticks = [(x, f"{'Jan' if x % 12 == 0 else 'Jul'} "
              f"{p + (x + (0 if x % 12 == 0 else -6)) // 12}")
             for x in range(int(wx0), 25, 6)]
    title = ("On-level view — policy year converted to calendar year "
             f"(term {int(fr.term)} mo"
             + (", NET: combined price path q" if fr.net_mode else "")
             + ")")
    return (polys * span * diags * vlines * labels * ann * ind).opts(
        responsive=True, height=380, xlabel="", ylabel="policy term elapsed",
        xticks=ticks, ylim=(-0.03, 1.19), xlim=(wx0 - 0.5, wx1 + 0.5),
        title=title, fontsize={"title": "9pt"})
