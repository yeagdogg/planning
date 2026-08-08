"""The on-leveling rectangle — policy year converted to calendar year,
drawn (W2d).

John's ask, verbatim: "a rectangular area for plan year and plan year
plus 1 with the earned rate levels identified in each area — really just
a visualization of the whole reason this app exists: conversion from
policy year to calendar year." This is the textbook parallelogram
diagram, live for the selected combo:

- x: calendar time across CY P and CY P+1 (months since Jan P);
- y: fraction of the policy term elapsed — a policy written at t earns
  along the diagonal (t, 0) → (t + term, 1);
- every rate change draws its diagonal; the regions between diagonals
  carry the cumulative rate index, shaded and labeled;
- each calendar year is annotated with the PRODUCTION earned level and
  A_rate from the EngineResult (the monthly-convention numbers every
  other exhibit uses).

Geometry is the engine's CONTINUOUS cross-check convention (uniform
writings, seasonality not drawn) in the engine's own month-coordinate
space; the tie test rebuilds ``engine.continuous_earned_index`` from the
region areas to 1e-9, so the picture cannot drift from the math. When
the production monthly number differs visibly from the continuous one
(seasonality, short terms), BOTH print — disclosed, never hidden. Net
combos draw the LOGGED program (the D65 program basis): the net path is
a year-over-year recursion, not a set of effective dates, and drawing it
as regions would be a lie — the banner says so.
"""
from __future__ import annotations

from dataclasses import dataclass, field

from src.engine import month_coord, month_index

from app.glue.theme import NAVY, STEEL
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
    net_mode: bool
    window: tuple                       # (0.0, 24.0) in months-since-Jan-P
    regions: list = field(default_factory=list)
    diagonals: list = field(default_factory=list)
    years: list = field(default_factory=list)
    notes: list = field(default_factory=list)


def onlevel_frame(ci, res) -> OnlevelFrame:
    """Pure geometry: regions/diagonals/year annotations for one combo.
    x is measured in months since Jan of the plan year."""
    p = res.plan_year
    tau = float(ci.term_months)
    base = float(month_index(p, 1))
    wx0, wx1 = 0.0, 24.0
    fr = OnlevelFrame(plan_year=p, term=tau, net_mode=bool(ci.net_mode),
                      window=(wx0, wx1))

    changes = sorted(((month_coord(rc.effective) - base, rc.effective_pct,
                       rc.effective) for rc in ci.rate_changes),
                     key=lambda t: t[0])
    # region bands: [lo, hi) carrying the cumulative index; the base band
    # opens far enough left that nothing earning in-window predates it
    bounds = [wx0 - tau - 1.0] + [c for c, _pct, _d in changes] + [wx1 + 1.0]
    idx, indices = 1.0, [1.0]
    opened = ["base level"]
    for _c, pct, d in changes:
        idx *= 1.0 + pct
        indices.append(idx)
        opened.append(str(d))

    rect = [(wx0, 0.0), (wx1, 0.0), (wx1, 1.0), (wx0, 1.0)]
    for k in range(len(bounds) - 1):
        lo, hi = bounds[k], bounds[k + 1]
        poly = _clip(rect, lambda x, y, lo=lo: lo - (x - tau * y))
        poly = _clip(poly, lambda x, y, hi=hi: (x - tau * y) - hi)
        if len(poly) < 3 or _area(poly) < 1e-9:
            continue
        shares = []
        for ys in (0.0, 12.0):
            yp = _clip(poly, lambda x, y, ys=ys: ys - x)
            yp = _clip(yp, lambda x, y, ye=ys + 12.0: x - ye)
            shares.append(_area(yp) / 12.0 if len(yp) >= 3 else 0.0)
        cx, cy = _centroid(poly)
        fr.regions.append(dict(
            poly=poly, index=indices[k], label=f"{indices[k]:.4f}",
            opened=opened[k], share_p=shares[0], share_p1=shares[1],
            cx=cx, cy=cy))

    for c, _pct, d in changes:
        y_lo = max(0.0, (wx0 - c) / tau)
        y_hi = min(1.0, (wx1 - c) / tau)
        if y_lo < y_hi:
            fr.diagonals.append(dict(
                x0=c + tau * y_lo, y0=y_lo, x1=c + tau * y_hi, y1=y_hi,
                date=str(d)))

    from src.engine import continuous_earned_index
    for j, yr in enumerate((p, p + 1)):
        e_prod = res.e_cy[yr]
        a_prod = res.a_rate_p if yr == p else res.a_rate_p1
        e_cont = continuous_earned_index(ci.rate_changes, yr,
                                         ci.term_months)
        fr.years.append(dict(year=yr, x0=j * 12.0, x1=(j + 1) * 12.0,
                             e_cy=e_prod, a_rate=a_prod, e_cont=e_cont))
        if not ci.net_mode and abs(e_cont - e_prod) > 1e-3:
            fr.notes.append(
                f"CY {yr}: the diagram's uniform-writings geometry earns "
                f"{e_cont:.4f}; the production monthly convention earns "
                f"{e_prod:.4f} (seasonality / term effects) — the "
                f"annotation shows the production number.")
    if ci.net_mode:
        fr.notes.append(
            "NET SELECTION IN FORCE: this draws the LOGGED rate program "
            "(the program basis, D65). The net path renews year-over-year "
            "and has no effective-date diagonals — the annotated earned "
            "levels are the engine's net-path values.")
    return fr


_XTICKS = None


def onlevel(ci, res, frame: OnlevelFrame | None = None):
    """Render the frame: shaded index regions, change diagonals, year
    bounds, region labels, and the per-year earned-level annotations."""
    hv = ensure_hv()
    from app.views.earning import _mix

    fr = frame if frame is not None else onlevel_frame(ci, res)
    p = fr.plan_year
    lo = min((r["index"] for r in fr.regions), default=1.0)
    hi = max((r["index"] for r in fr.regions), default=1.0)

    def shade(ix):
        t = 0.15 if hi <= lo else 0.15 + 0.75 * (ix - lo) / (hi - lo)
        return _mix("#FFFFFF", STEEL, t)

    polys = hv.Polygons(
        [{("x", "y"): r["poly"], "level": r["index"],
          "Rate level": r["label"], "Opened by": r["opened"],
          "Share of CY P": f"{r['share_p']:.1%}",
          "Share of CY P+1": f"{r['share_p1']:.1%}"}
         for r in fr.regions],
        vdims=["level", "Rate level", "Opened by", "Share of CY P",
               "Share of CY P+1"]).opts(
        color="level", cmap=[shade(r["index"]) for r in
                             sorted(fr.regions, key=lambda r: r["index"])]
        or ["#FFFFFF"],
        line_color="white", line_width=1.2, alpha=0.9, tools=["hover"],
        hover_tooltips=[("rate level", "@{Rate level}"),
                        ("opened by", "@{Opened by}"),
                        ("share of CY P", "@{Share of CY P}"),
                        ("share of CY P+1", "@{Share of CY P+1}")],
        colorbar=False, show_legend=False)

    diags = hv.Overlay([
        hv.Curve([(d["x0"], d["y0"]), (d["x1"], d["y1"])]).opts(
            color=NAVY, line_width=2)
        for d in fr.diagonals]) if fr.diagonals else hv.Overlay([])

    years = hv.Overlay([hv.VLine(x).opts(color="#666666", line_width=1.4,
                                         line_dash="dotted")
                        for x in (0.0, 12.0, 24.0)])

    labels = hv.Labels(
        [(r["cx"], r["cy"], r["label"]) for r in fr.regions
         if max(r["share_p"], r["share_p1"]) > 0.02],
        kdims=["x", "y"], vdims=["text"]).opts(
        text_font_size="9pt", text_color=NAVY, text_font_style="bold")

    ann = hv.Labels(
        [(y["x0"] + 6.0, 1.07,
          f"CY {y['year']}   earned level {y['e_cy']:.4f}   "
          f"A_rate {y['a_rate']:.4f}") for y in fr.years],
        kdims=["x", "y"], vdims=["text"]).opts(
        text_font_size="10pt", text_color=NAVY, text_font_style="bold")

    ticks = [(0, f"Jan {p}"), (6, f"Jul {p}"), (12, f"Jan {p + 1}"),
             (18, f"Jul {p + 1}"), (24, f"Jan {p + 2}")]
    title = ("On-level view — policy year converted to calendar year "
             f"(term {int(fr.term)} mo"
             + (", NET: drawing the logged program" if fr.net_mode else "")
             + ")")
    return (polys * diags * years * labels * ann).opts(
        responsive=True, height=380, xlabel="", ylabel="policy term elapsed",
        xticks=ticks, ylim=(-0.03, 1.16), xlim=(-0.5, 24.5),
        title=title, fontsize={"title": "9pt"})
