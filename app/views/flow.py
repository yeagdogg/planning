"""The per-combo flow depth (W3d) — the Excel Flow Dashboard, reframed.

Everything here ARRANGES engine output, never re-derives it: the
36-month index/mod/price series read straight off the EngineResult
series (cohorts run Jan P-2 .. Dec P+1 and earned months Jan P-1 .. Dec
P+1, so the whole window is already published); the eleven program-leg
columns are COPIED from ``engine.program_flow_by_month`` rows verbatim
(prior_rows carry Jan..Dec P-1, rows carry the 24 YoY-able months); the
carryover ledger mirrors the workbook's own cell arithmetic over the
same w(k)·ec(k, year) masses the RE block publishes — per action, in
log space, with the day-blend on the effective month — and its total is
LINKED to ``res.yoy_earned_p1``, never recomputed.

Unearned runway = written / earned - 1 — "the most plannable number in
the file": rate locked on the books but not yet earned. It converges to
zero as filed actions finish earning and jumps when new rate lands.
"""
from __future__ import annotations

import math

import pandas as pd

from src.engine import mi_month, mi_year, month_index

from app.glue.theme import NAVY, STEEL, STEEL_LIGHT, FAIL_RED, chart_colors
from . import ensure_hv

_LEG_KEYS = ("rate_leg", "mod_leg", "delivered", "locked_leg",
             "planned_residual", "locked_mod_leg", "planned_mod_residual",
             "locked_delivered", "planned_delivered_residual")


def _label(mi: int) -> str:
    return f"{pd.Timestamp(mi_year(mi), mi_month(mi), 1):%b %Y}"


def dashboard_frame(res, pf, combo) -> pd.DataFrame:
    """36 rows, Jan P-1 .. Dec P+1 — the Flow Dashboard's monthly detail.

    written/earned (and the mod twins) come off the EngineResult series;
    the program legs are ``pf.prior_rows + pf.rows`` verbatim. Price
    columns exist only when the mod adjustment is live and the combo is
    NOT net (the net path already combines rate and price — drawing a
    second combined series would double-count, and the page says so).
    """
    p = res.plan_year
    m0 = month_index(p - 1, 1)
    mods_on = bool(combo.mod_adjustment_enabled) and not combo.net_mode
    m_ind = combo.mods.m_ind
    legs = list(pf.prior_rows) + list(pf.rows)
    assert len(legs) == 36 and legs[0]["mi"] == m0, "leg window mismatch"

    rows = []
    for i in range(36):
        mi = m0 + i
        written = res.inforce_index[res.cohort_mi.index(mi)]
        earned = res.earned_index_m[res.earned_month_mi.index(mi)]
        wmod = (res.written_mod[res.cohort_mi.index(mi)]
                if mods_on else None)
        emod = (res.earned_mod_m[res.earned_month_mi.index(mi)]
                if mods_on else None)
        rec = dict(
            mi=mi, month=pd.Timestamp(mi_year(mi), mi_month(mi), 1),
            label=_label(mi), year=mi_year(mi), plan_band=mi_year(mi) == p,
            written=written, earned=earned,
            wmod=wmod, emod=emod,
            price_w=(written * wmod / m_ind) if mods_on else None,
            price_e=(earned * emod / m_ind) if mods_on else None,
            runway=written / earned - 1.0,
        )
        for k in _LEG_KEYS:
            rec[k] = legs[i].get(k)
        rec["w"] = legs[i]["w"]
        if combo.net_mode and mi >= month_index(p, 1):
            x = combo.net_sel_p if mi < month_index(p + 1, 1) \
                else combo.net_x1
            d = rec["delivered"]
            rec["delta_vs_net"] = None if d is None else d - x
        else:
            rec["delta_vs_net"] = None
        rows.append(rec)

    df = pd.DataFrame(rows)
    # pandas coerces a mixed None/float column to float64 (None -> NaN);
    # the charts' is-None gates and the "—" hover text depend on real
    # Nones, so the nullable columns stay object dtype
    df["delta_vs_net"] = pd.Series([r["delta_vs_net"] for r in rows],
                                   dtype=object)
    # YoY of the frame's own earned series — only months with a year-ago
    # base inside the window (the last 24)
    for col, src in (("yoy_earned_rate", "earned"),
                     ("yoy_earned_price", "price_e")):
        vals = [None] * 36
        for i in range(12, 36):
            a, b = df[src].iloc[i], df[src].iloc[i - 12]
            if a is not None and b is not None and not (
                    isinstance(b, float) and b == 0.0):
                vals[i] = a / b - 1.0
        df[col] = pd.Series(vals, dtype=object)
    return df


def ledger_frame(eng, combo, res):
    """The carryover ledger — per rate action, the workbook's own cells:
    share of each year's earned mass from the action's date forward (the
    effective month day-blended), the log-space contribution
    ln(1+pct)·(share_p1 − share_p), the combined product, the disclosed
    timing residual, and the total LINKED to ``res.yoy_earned_p1``."""
    import datetime as dt

    p = res.plan_year
    den = {}
    mass = {}
    for yr in (p, p + 1):
        mass[yr] = {k: eng.w(k) * eng.ec(k, yr) for k in eng.cohorts}
        den[yr] = sum(mass[yr].values())

    def share(d: dt.date, yr: int) -> float:
        mo = d.year * 12 + d.month - 1               # engine month_index
        eom = (dt.date(d.year + (d.month == 12), d.month % 12 + 1, 1)
               - dt.timedelta(days=1))
        frac = (eom - d + dt.timedelta(days=1)).days / eom.day
        tot = sum(v for k, v in mass[yr].items() if k > mo)
        tot += frac * mass[yr].get(mo, 0.0)
        return tot / den[yr]

    rows, logsum = [], 0.0
    for j, rc in enumerate(sorted(combo.rate_changes,
                                  key=lambda r: r.effective), start=1):
        sp = share(rc.effective, p)
        sp1 = share(rc.effective, p + 1)
        lg = math.log(1.0 + rc.effective_pct) * (sp1 - sp)
        logsum += lg
        rows.append(dict(seq=j, eff=rc.effective, pct=rc.effective_pct,
                         share_p=sp, share_p1=sp1,
                         contribution=math.exp(lg) - 1.0, log=lg))
    combined = math.exp(logsum) - 1.0
    totals = dict(
        combined=combined,
        residual=(1.0 + res.yoy_earned_p1) / math.exp(logsum) - 1.0,
        total=res.yoy_earned_p1,                     # linked, never derived
        net_mode=bool(combo.net_mode))
    return pd.DataFrame(rows), totals


# ------------------------------------------------------------------ charts

def _band(p):
    hv = ensure_hv()
    return hv.VSpan(pd.Timestamp(p, 1, 1), pd.Timestamp(p + 1, 1, 1)).opts(
        color=STEEL_LIGHT, alpha=0.35)


def index_chart(df, p):
    """Written vs earned rate level over the full 36 months, plan year
    banded — the Flow Dashboard's opening chart, one year wider than the
    Earning tab's 24-month lag view."""
    hv = ensure_hv()
    d = df.assign(w_txt=df["written"].map(lambda v: f"{v:.4f}"),
                  e_txt=df["earned"].map(lambda v: f"{v:.4f}"))
    written = hv.Curve(d, "month", ["written", "label", "w_txt"]).relabel(
        "Written (in force)").opts(
        color=NAVY, line_width=2.5, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("written", "@w_txt")])
    earned = hv.Curve(d, "month", ["earned", "label", "e_txt"]).relabel(
        "Earned level").opts(
        color=STEEL, line_width=2, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("earned", "@e_txt")])
    return (_band(p) * earned * written).opts(
        responsive=True, height=320, xlabel="", ylabel="rate index",
        legend_position="top_left",
        title=(f"Written vs earned level, Jan {p - 1} – Dec {p + 1} "
               "(plan year banded)"),
        fontsize={"title": "9pt"})


def yoy_bars(df, p, net_mode=False):
    """Monthly YoY earned rate vs earned price — when they detach, the
    schedule mods are doing work the rate filings are not."""
    hv = ensure_hv()
    recs = []
    for _i, r in df.iterrows():
        if r["yoy_earned_rate"] is None:
            continue
        recs.append((r["label"], "rate", r["yoy_earned_rate"]))
        if r["yoy_earned_price"] is not None:
            recs.append((r["label"], "price", r["yoy_earned_price"]))
    if not recs:
        return None
    bars = hv.Bars(recs, kdims=["label", "kind"], vdims=["yoy"])
    title = "Monthly YoY earned rate vs price"
    if net_mode:
        title = "Monthly YoY earned level (net path combines rate + price)"
    return bars.opts(
        responsive=True, height=300, xrotation=90, xlabel="",
        ylabel="YoY earned chg",
        cmap={"rate": STEEL, "price": NAVY}, color="kind",
        show_legend=True, legend_position="top_left",
        title=title, fontsize={"title": "9pt"}, tools=["hover"])


def runway_line(df, p):
    hv = ensure_hv()
    d = df.assign(r_txt=df["runway"].map(lambda v: f"{v:+.2%}"))
    line = hv.Curve(d, "month", ["runway", "label", "r_txt"]).opts(
        color=NAVY, line_width=2.2, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("runway", "@r_txt")])
    zero = hv.HLine(0.0).opts(color="#666666", line_dash="dotted",
                              line_width=1.2)
    return (_band(p) * zero * line).opts(
        responsive=True, height=280, xlabel="",
        ylabel="written / earned − 1",
        title=("Unearned runway — rate locked on the books but not yet "
               "earned (converges to zero; jumps when new rate lands)"),
        fontsize={"title": "9pt"})


def mod_price_chart(df, combo, p):
    """Written vs earned schedule-mod path + the combined price index.
    None when the combo is mod-off or net (the caller shows a note)."""
    if df["wmod"].iloc[0] is None:
        return None
    hv = ensure_hv()
    d = df.assign(wm=df["wmod"].astype(float), em=df["emod"].astype(float),
                  pw=df["price_w"].astype(float),
                  pe=df["price_e"].astype(float))
    wmod = hv.Curve(d, "month", ["wm", "label"]).relabel("Written mod").opts(
        color=NAVY, line_width=2, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("written mod", "@wm{0.0000}")])
    emod = hv.Curve(d, "month", ["em", "label"]).relabel("Earned mod").opts(
        color=STEEL, line_width=2)
    pw = hv.Curve(d, "month", ["pw", "label"]).relabel(
        "Price (written)").opts(color=NAVY, line_width=1.6,
                                line_dash="dashed")
    pe = hv.Curve(d, "month", ["pe", "label"]).relabel(
        "Price (earned)").opts(color=STEEL, line_width=1.6,
                               line_dash="dashed")
    return (_band(p) * emod * wmod * pe * pw).opts(
        responsive=True, height=300, xlabel="",
        ylabel="mod / combined price",
        legend_position="top_left",
        title=("Schedule-mod path and combined price index "
               "(price = rate × mod / M_ind)"),
        fontsize={"title": "9pt"})


def delivery_chart(df, combo, p):
    """The program on renewals: delivered net (navy) vs the rate leg
    (steel) vs LOCKED delivered dashed — what survives if no planned
    action lands (D77). Hover carries the split and Δ vs net assertion."""
    hv = ensure_hv()
    d = df.loc[df["delivered"].map(lambda v: v is not None)].copy()

    def txt(col, fmt="{:+.2%}"):
        return d[col].map(lambda v: "—" if v is None else fmt.format(v))

    d["del_txt"] = txt("delivered")
    d["rate_txt"] = txt("rate_leg")
    d["lock_txt"] = txt("locked_delivered")
    d["net_txt"] = txt("delta_vs_net")
    hover = [("month", "@label"), ("delivered", "@del_txt"),
             ("rate leg", "@rate_txt"), ("locked delivered", "@lock_txt")]
    if combo.net_mode:
        hover.append(("Δ vs net assertion", "@net_txt"))
    deliv = hv.Curve(d, "month", ["delivered", "label", "del_txt",
                                  "rate_txt", "lock_txt",
                                  "net_txt"]).relabel("Delivered net").opts(
        color=NAVY, line_width=2.5, tools=["hover"], hover_tooltips=hover)
    rate = hv.Curve(d, "month", ["rate_leg", "label"]).relabel(
        "Rate leg").opts(color=STEEL, line_width=2)
    locked = hv.Curve(d, "month", ["locked_delivered", "label"]).relabel(
        "Locked delivered").opts(color="#9a9a9a", line_width=1.8,
                                 line_dash="dashed")
    overlay = _band(p) * rate * locked * deliv
    if combo.net_mode and combo.net_sel_p is not None:
        overlay = overlay * hv.HLine(combo.net_sel_p).opts(
            color=FAIL_RED, line_dash="dotted", line_width=1.4)
    return overlay.opts(
        responsive=True, height=320, xlabel="",
        ylabel="YoY on renewals",
        legend_position="top_left",
        title=("Program delivery on renewals — delivered vs rate leg vs "
               "LOCKED (dashed: no planned action lands)"),
        fontsize={"title": "9pt"})


def ledger_html(ldf, totals, p) -> str:
    """The carryover ledger as a small exhibit table."""
    cc = chart_colors()
    rows = "".join(
        f"<tr><td style='text-align:center'>{r.seq}</td>"
        f"<td>{r.eff}</td>"
        f"<td style='text-align:right'>{r.pct:+.1%}</td>"
        f"<td style='text-align:right'>{r.share_p:.1%}</td>"
        f"<td style='text-align:right'>{r.share_p1:.1%}</td>"
        f"<td style='text-align:right'>{r.contribution:+.2%}</td></tr>"
        for r in ldf.itertuples())
    foot = (
        f"<tr style='border-top:1.5px solid {NAVY}; font-weight:600'>"
        f"<td colspan='5'>Actions combined</td>"
        f"<td style='text-align:right'>{totals['combined']:+.2%}</td></tr>"
        f"<tr><td colspan='5'>Timing / compounding interaction "
        f"(residual)</td>"
        f"<td style='text-align:right'>{totals['residual']:+.2%}</td></tr>"
        f"<tr style='font-weight:700; color:{NAVY}'>"
        f"<td colspan='5'>Total carryover — earned rate chg into "
        f"CY {p + 1}</td>"
        f"<td style='text-align:right'>{totals['total']:+.2%}</td></tr>")
    cap = ("Net selection active: the carryover is set by the net renewal "
           "path — this ledger describes the explicit program "
           "counterfactual." if totals["net_mode"] else
           "Each action earns partly in the plan year and fully in the "
           "next; the difference is carryover.")
    head = "".join(f"<th style='text-align:{a}; padding:2px 8px'>{h}</th>"
                   for h, a in (("#", "center"), ("Effective", "left"),
                                ("Effective %", "right"),
                                (f"Share of CY {p} earned", "right"),
                                (f"Share of CY {p + 1} earned", "right"),
                                ("Contribution", "right")))
    return (
        f"<div style='font-variant-numeric:tabular-nums; color:{cc['actual']}'>"
        f"<div style='font-weight:700; color:{NAVY}; margin:6px 0 2px'>"
        f"Carryover ledger — where the head start into CY {p + 1} comes "
        f"from</div>"
        f"<div style='font-style:italic; color:#595959; font-size:0.85em; "
        f"margin-bottom:4px'>{cap}</div>"
        f"<table style='border-collapse:collapse; font-size:0.9em'>"
        f"<tr style='color:#595959; border-bottom:1px solid #e3e6ea'>"
        f"{head}</tr>{rows}{foot}</table></div>")
