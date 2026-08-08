"""The race and the delivery — monthly LR walk + where premium lands (P3).

The walk IS ``engine.lr_flow_by_month`` — the same function the workbook's
LR Flow tab is verified against (D93) — reframed, never recomputed. The
band is ``MonthlyEngine.ec``: the closed-form share of each plan-year
cohort the plan year keeps vs hands to the year after (the D110 band row,
drawn).
"""
from __future__ import annotations

import pandas as pd

from src.engine import LrFlowResult, mi_month, mi_year, month_index

from app.glue.theme import FAIL_RED, NAVY, STEEL, STEEL_LIGHT
from . import ensure_hv
from .earning import _ym


def walk_frame(flow: LrFlowResult) -> pd.DataFrame:
    """The 24 engine rows, one per month Jan P .. Dec P+1. Dead months
    (nothing earned) carry None and plot as gaps, never zeros."""
    rows = []
    for r in flow.rows:
        ts = pd.Timestamp(mi_year(r["mi"]), mi_month(r["mi"]), 1)
        rows.append(dict(
            month=ts, label=f"{ts:%b %Y}", lr=r["lr"],
            a_rate=r["a_rate"], a_mod=r["a_mod"],
            trend_factor=r["trend_factor"], weight=r["weight"],
            lr_txt="—" if r["lr"] is None else f"{r['lr']:.1%}",
            fac_txt=("—" if r["a_rate"] is None else
                     f"{r['a_rate']:.4f} / {r['a_mod']:.4f} / "
                     f"{r['trend_factor']:.4f}")))
    return pd.DataFrame(rows)


def race_sentence(flow: LrFlowResult, trend: float) -> str:
    """The workbook's one-sentence verdict (LR Flow A6), in app dress."""
    jan, dec = flow.rows[0]["lr"], flow.rows[11]["lr"]
    if flow.breakeven_trend is None or jan is None or dec is None:
        return ("Rate and price do not move this year — the loss ratio "
                "walks with trend alone, and there is no race to run.")
    t_star = flow.breakeven_trend
    if abs(trend - t_star) < 0.0005:
        verdict = "is **FLAT**: trend and earn-in cancel"
    elif trend > t_star:
        verdict = (f"**DETERIORATES** from {jan:.1%} in January to "
                   f"{dec:.1%} in December — trend is winning")
    else:
        verdict = (f"**IMPROVES** from {jan:.1%} in January to "
                   f"{dec:.1%} in December — earn-in is winning")
    return (f"At {trend:.1%} net trend the year {verdict}. "
            f"They break even at {t_star:.1%} trend.")


def walk(flow: LrFlowResult, target: float | None = None):
    """The race, drawn: monthly plan LR across P and P+1, the target as a
    dashed line, the following year tinted (indicative, as everywhere)."""
    hv = ensure_hv()
    from bokeh.models import NumeralTickFormatter

    df = walk_frame(flow)
    p = flow.plan_year
    span = hv.VSpan(pd.Timestamp(p + 1, 1, 1),
                    pd.Timestamp(p + 2, 1, 1)).opts(
        color=STEEL_LIGHT, alpha=0.35)
    curve = hv.Curve(df, "month",
                     ["lr", "label", "lr_txt", "fac_txt"]).relabel(
        "Plan LR").opts(
        color=NAVY, line_width=2.5, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("plan LR", "@lr_txt"),
                        ("A_rate / A_mod / trend", "@fac_txt")])
    overlay = span * curve
    if target:
        overlay = overlay * hv.HLine(float(target)).opts(
            color=FAIL_RED, line_dash="dashed", line_width=1.5)
    return overlay.opts(
        responsive=True, height=320, ylabel="plan loss ratio", xlabel="",
        yformatter=NumeralTickFormatter(format="0%"),
        title=f"The race — plan LR by month, CY {p} then CY {p + 1}",
        fontsize={"title": "9pt"})


def band_frame(eng, plan_year: int) -> pd.DataFrame:
    """Where each plan-year cohort's premium lands: the share the plan year
    keeps (``ec(k, P)``) vs the share handed to P+1. Stacks to one for any
    term ≤ 12 — asserted by the tests, not silently assumed."""
    rows = []
    for j in range(12):
        k = month_index(plan_year, 1) + j
        for year in (plan_year, plan_year + 1):
            share = eng.ec(k, year)
            rows.append(dict(month=_ym(k), dest=f"earned in CY {year}",
                             share=share, w=eng.w(k),
                             share_txt=f"{share:.1%}"))
    return pd.DataFrame(rows)


def band(eng, plan_year: int):
    """The delivery band: how much of each month it writes the plan year
    actually keeps — the geometry under every A_rate in the book."""
    hv = ensure_hv()
    from bokeh.models import NumeralTickFormatter

    df = band_frame(eng, plan_year)
    bars = hv.Bars(df, ["month", "dest"], ["share", "share_txt", "w"])
    return bars.opts(
        stacked=True, responsive=True, height=320,
        color="dest", cmap={f"earned in CY {plan_year}": NAVY,
                            f"earned in CY {plan_year + 1}": STEEL},
        xrotation=90, xlabel="written month", ylabel="share of cohort",
        legend_position="top_right",
        yformatter=NumeralTickFormatter(format="0%"), tools=["hover"],
        hover_tooltips=[("written", "@month"), ("", "@dest"),
                        ("share", "@share_txt")],
        title=(f"Share earned within the plan year — ec(k, P): what "
               f"CY {plan_year} keeps of each month it writes"),
        fontsize={"title": "9pt"})


# --------------------------------------------------- W3d: the walk, deeper

def walk_table_frame(flow: LrFlowResult, target=None) -> pd.DataFrame:
    """The LR Flow walk TABLE: the 24 engine rows verbatim plus the one
    disclosed subtraction (vs target, in points). Dead months keep None
    everywhere (rendered em-dash, never zero)."""
    rows = []
    for r in flow.rows:
        ts = pd.Timestamp(mi_year(r["mi"]), mi_month(r["mi"]), 1)
        vs = (None if r["lr"] is None or target is None
              else (r["lr"] - float(target)) * 100.0)
        rows.append(dict(
            label=f"{ts:%b %Y}", trend_factor=r["trend_factor"],
            a_rate=r["a_rate"], a_mod=r["a_mod"], lr=r["lr"],
            vs_target=vs, den=r["den"], weight=r["weight"]))
    return pd.DataFrame(rows)


def headroom(flow: LrFlowResult, trend: float):
    """(points, bad): trend minus breakeven — positive means trend beats
    earn-in and the year deteriorates. None when there is no race."""
    if flow.breakeven_trend is None:
        return None, False
    pts = (trend - flow.breakeven_trend) * 100.0
    return pts, pts > 0


def recon_rows(flow: LrFlowResult) -> list:
    """The LR Flow reconciliation strip: premium-weighted mean vs the CY
    headline, the residual in bp decomposed into covariance + trend
    centre-of-gravity tilt + convexity, and the identity check — which
    must be zero to floating point, and is (mean = head x (1+tilt) x
    (1+convexity) / (1+cov))."""
    bp = 10_000.0
    ident = flow.mean_p - flow.head_p * (1.0 + flow.cog_tilt_p) \
        * (1.0 + flow.convexity_p) / (1.0 + flow.cov_rate_price)
    return [
        ("Premium-weighted mean, CY P", f"{flow.mean_p:.2%}"),
        ("CY P headline", f"{flow.head_p:.2%}"),
        ("Residual", f"{flow.resid_p * bp:+.1f} bp"),
        ("…of which rate × price covariance",
         f"{flow.cov_rate_price * bp:+.1f} bp"),
        ("…of which trend centre-of-gravity tilt",
         f"{flow.cog_tilt_p * bp:+.1f} bp"),
        ("…of which trend convexity (Jensen)",
         f"{flow.convexity_p * bp:+.1f} bp"),
        ("Unexplained (identity check — must be zero)",
         f"{ident * bp:+.2f} bp"),
        ("Mean vs headline, CY P+1",
         f"{flow.mean_p1:.2%} vs {flow.head_p1:.2%}"),
    ]


def recon_html(flow: LrFlowResult) -> str:
    rows = "".join(
        f"<tr><td style='padding:1px 10px 1px 0; color:#595959'>{k}</td>"
        f"<td style='text-align:right; font-weight:600'>{v}</td></tr>"
        for k, v in recon_rows(flow))
    return (f"<div style='font-variant-numeric:tabular-nums; "
            f"color:{NAVY}; font-size:0.9em'>"
            f"<div style='font-weight:700; margin:6px 0 2px'>Mean vs "
            f"headline — the D46 reconciliation (disclosed, never "
            f"normalised away)</div>"
            f"<table style='border-collapse:collapse'>{rows}</table></div>")
