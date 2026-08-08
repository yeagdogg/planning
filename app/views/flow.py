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


# ------------------------------------------------- W3e: book-level flow

def combine_rows(flows, attr: str, sl: slice) -> list:
    """The engine's own EP×w(m) combination (D60), reapplied over CACHED
    per-combo ``ProgramFlowResult``s — and extended to the P+1 half
    (``rows[12:24]``) that ``CombinedFlowResult`` does not carry. Pinned
    to ``engine.combined_flow_by_month`` at 1e-15 on the slices both
    compute; the same arithmetic on the extra slice is arrangement, not
    re-derivation. Delivered is the exact statistic: weighted rate x
    weighted mod need NOT reproduce it under mix, and nothing here
    normalises the gap away. The mod leg weights only combos carrying
    the adjustment (None when none in view do); mod-off combos still
    deliver their rate leg — the honest book statistic."""
    live = [(ep, pf) for ep, pf in flows if ep and ep > 0.0]
    if not live:
        raise ValueError("combine_rows needs at least one combo with EP.")
    n = len(getattr(live[0][1], attr))
    out = []
    for j in range(*sl.indices(n)):
        nr = nm = nd_ = dw = dm = 0.0
        mi = None
        for ep, pf in live:
            prow = getattr(pf, attr)[j]
            mi = prow["mi"]
            epw = ep * prow["w"]
            dw += epw
            nr += epw * prow["rate_leg"]
            nd_ += epw * prow["delivered"]
            if pf.mod_on:
                dm += epw
                nm += epw * prow["mod_leg"]
        out.append(dict(mi=mi, rate_leg=nr / dw,
                        mod_leg=(nm / dm) if dm > 0.0 else None,
                        delivered=nd_ / dw))
    return out


def combine_avgs(flows) -> dict:
    """EP-weighted means of the per-combo AVERAGE ratios — the engine's
    top-level combination rule verbatim, extended to the prior-year rate
    and mod averages the engine result does not publish (the workbook
    gets those by weighting the per-combo PRIOR_PUB columns — same
    arithmetic). Ratios in, deltas out (−1 applied here, once)."""
    ar = am = ad = de = dme = adp = arp = amp = 0.0
    for ep, pf in flows:
        if not ep or ep <= 0.0:
            continue
        de += ep
        ar += ep * pf.avg_rate_ratio
        ad += ep * pf.avg_delivered_ratio
        adp += ep * pf.avg_delivered_ratio_prior
        arp += ep * pf.avg_rate_ratio_prior
        if pf.mod_on:
            dme += ep
            am += ep * pf.avg_mod_ratio
            amp += ep * pf.avg_mod_ratio_prior
    if de <= 0.0:
        return {}
    return dict(
        rate=ar / de - 1.0, delivered=ad / de - 1.0,
        mod=(am / dme - 1.0) if dme > 0.0 else None,
        rate_prior=arp / de - 1.0, delivered_prior=adp / de - 1.0,
        mod_prior=(amp / dme - 1.0) if dme > 0.0 else None,
        ep=de)


def grid_cols(prior: bool = True, p1: bool = False) -> list:
    """Month-column keys: pm01..pm12 (CY P−1), pp01..pp12 (CY P),
    q01..q12 (CY P+1)."""
    cols = []
    if prior:
        cols += [f"pm{j:02d}" for j in range(1, 13)]
    cols += [f"pp{j:02d}" for j in range(1, 13)]
    if p1:
        cols += [f"q{j:02d}" for j in range(1, 13)]
    return cols


def state_grid_frame(view: dict, all_flows: list, leg: str,
                     avg_label: str, p1: bool = False) -> pd.DataFrame:
    """One row per state (each state's combos combined by the D60 rule)
    plus the AVG bottom row over everything in view — the workbook's
    BOOK AVG / LINE AVG / BU AVG line: an average of the state rows,
    never a total. ``leg`` ∈ rate_leg | mod_leg | delivered. With
    ``p1`` the columns run Jan P .. Dec P+1 (the Net Delivery grid);
    otherwise Jan P−1 .. Dec P."""
    def cells(flows):
        rec = {}
        if p1:
            rows = combine_rows(flows, "rows", slice(0, 24))
            for j in range(12):
                rec[f"pp{j + 1:02d}"] = rows[j][leg]
                rec[f"q{j + 1:02d}"] = rows[j + 12][leg]
        else:
            prior = combine_rows(flows, "prior_rows", slice(0, 12))
            plan = combine_rows(flows, "rows", slice(0, 12))
            for j in range(12):
                rec[f"pm{j + 1:02d}"] = prior[j][leg]
                rec[f"pp{j + 1:02d}"] = plan[j][leg]
        return rec

    recs = []
    for st in sorted(view):
        flows = [(r["ep"], r["pf"]) for r in view[st]
                 if r["pf"] is not None]
        if not any(ep and ep > 0.0 for ep, _pf in flows):
            continue
        recs.append({"state": st, **cells(flows)})
    if any(ep and ep > 0.0 for ep, _pf in all_flows):
        recs.append({"state": avg_label, **cells(all_flows)})
    cols = ["state"] + grid_cols(prior=not p1, p1=p1)
    df = pd.DataFrame(recs, columns=cols)
    for c in cols[1:]:                    # None must SURVIVE (never NaN)
        df[c] = pd.Series([r.get(c) for r in recs], dtype=object)
    return df


def year_band_styles(df: pd.DataFrame, bands: list) -> pd.DataFrame:
    """The workbook's per-year color scale: ONE scale domain per year
    band (a single scale across both years would flatten the contrast
    inside each), computed over every body cell in the band — the AVG
    bottom row is excluded from the domain and untinted, like TOTAL
    everywhere else."""
    from app.summary import _lerp

    css = pd.DataFrame("", index=df.index, columns=df.columns)
    body = df.index[:-1] if len(df) else df.index
    for cols in bands:
        vals = [df.loc[i, c] for i in body for c in cols
                if isinstance(df.loc[i, c], (int, float))]
        if len(vals) < 2 or min(vals) == max(vals):
            continue
        lo, hi = min(vals), max(vals)
        vs = sorted(vals)
        med = vs[len(vs) // 2] if len(vs) % 2 else (
            vs[len(vs) // 2 - 1] + vs[len(vs) // 2]) / 2.0
        for i in body:
            for c in cols:
                v = df.loc[i, c]
                if not isinstance(v, (int, float)):
                    continue
                if v <= med:
                    t = 0.0 if med == lo else (v - lo) / (med - lo)
                    color = _lerp("#D6E8D5", "#FFF2CC", t)
                else:
                    t = 0.0 if hi == med else (v - med) / (hi - med)
                    color = _lerp("#FFF2CC", "#F4B8B8", t)
                css.loc[i, c] = f"background-color: {color}"
    if len(df):
        css.iloc[-1, :] = (f"background-color: {STEEL_LIGHT}; "
                           f"font-weight: 600")
    return css


def pf_summary_frame(view: dict, prog: dict,
                     avg_label: str) -> pd.DataFrame:
    """The Program Flow per-state summary: prior-year averages (the year
    currently flowing) beside the plan block — EP, net target (simple
    mean over net combos, the workbook rule), the three w-weighted
    averages, Δ vs net assertion, and the plan-LR gap program-vs-asserted
    in points (from the cached program-basis results; net combos only —
    for everything else the program IS the headline)."""
    def row_for(recs, st):
        flows = [(r["ep"], r["pf"]) for r in recs if r["pf"] is not None]
        a = combine_avgs(flows)
        if not a:
            return None
        nets = [r["netp"] for r in recs if r["netp"] is not None]
        tgt = sum(nets) / len(nets) if nets else None
        gaps = [(r["ep"], prog[(r["lob"], r["key"])][0] - r["lr_p"])
                for r in recs
                if (r["lob"], r["key"]) in prog and r["lr_p"] is not None
                and r["ep"]]
        gap = (sum(ep * g for ep, g in gaps) / sum(ep for ep, _g in gaps)
               * 100.0 if gaps else None)
        return dict(state=st, ep=a["ep"],
                    rate_pm1=a["rate_prior"], mod_pm1=a["mod_prior"],
                    del_pm1=a["delivered_prior"],
                    nettgt=tgt, rate_p=a["rate"], mod_p=a["mod"],
                    del_p=a["delivered"],
                    dnet=None if tgt is None else a["delivered"] - tgt,
                    proggap=gap,
                    netcnt=len(nets), cnt=len(recs))

    recs_all = [r for recs in view.values() for r in recs]
    rows = [r for st in sorted(view)
            if (r := row_for(view[st], st)) is not None]
    total = row_for(recs_all, avg_label)
    if total is not None:
        rows.append(total)
    cols = ["state", "ep", "rate_pm1", "mod_pm1", "del_pm1", "nettgt",
            "rate_p", "mod_p", "del_p", "dnet", "proggap", "netcnt",
            "cnt"]
    df = pd.DataFrame(rows, columns=cols)
    for c in ("mod_pm1", "nettgt", "mod_p", "dnet", "proggap"):
        df[c] = pd.Series([r.get(c) for r in rows], dtype=object)
    return df


def nd_summary_frame(view: dict, prog: dict, avg_label: str,
                     suggest: dict | None = None) -> pd.DataFrame:
    """The Book Net Delivery REPORT: targets vs delivered per state.
    Targets average over NET combos only ("a combo that asserts nothing
    is not a target of zero"); the P+1 target column says whether it was
    entered, carried from P, or mixed; plan LR program basis is the
    EP-weighted program-basis value over net combos, and the gap to the
    asserted headline rides beside it in points. ``suggest`` (W3f) is a
    per-state suggested-change map the page computes ONLY for states
    resolving to a single net combo — read-only prescriptive display,
    never an editing surface, and never aggregated."""
    def row_for(recs, st):
        flows = [(r["ep"], r["pf"]) for r in recs if r["pf"] is not None]
        a = combine_avgs(flows)
        if not a:
            return None
        nets = [r for r in recs if r["netp"] is not None]
        tgt = (sum(r["netp"] for r in nets) / len(nets)) if nets else None
        t1v = [(r["netp1"] if r["netp1"] is not None else r["netp"])
               for r in nets]
        tgt1 = sum(t1v) / len(t1v) if t1v else None
        src = {True: "entered", False: "carried"}
        kinds = {src[r["netp1"] is not None] for r in nets}
        set1 = ("—" if not nets else
                "mixed" if len(kinds) > 1 else next(iter(kinds)))
        pg = [(r["ep"], prog[(r["lob"], r["key"])], r["lr_p"])
              for r in nets if (r["lob"], r["key"]) in prog and r["ep"]]
        wep = sum(ep for ep, _p, _h in pg)
        progb = (sum(ep * p[0] for ep, p, _h in pg) / wep if pg else None)
        pgap = (sum(ep * (p[0] - h) for ep, p, h in pg) / wep * 100.0
                if pg else None)
        return dict(state=st, ep=a["ep"], netcnt=len(nets),
                    tgt=tgt, tgt1=tgt1, set1=set1,
                    delivered=a["delivered"],
                    gap=None if tgt is None else
                    (a["delivered"] - tgt) * 100.0,
                    progbasis=progb, proggap=pgap,
                    suggest=(suggest or {}).get(st))

    recs_all = [r for recs in view.values() for r in recs]
    rows = [r for st in sorted(view)
            if (r := row_for(view[st], st)) is not None]
    total = row_for(recs_all, avg_label)
    if total is not None:
        total["suggest"] = None            # never aggregated, by rule
        rows.append(total)
    cols = ["state", "ep", "netcnt", "tgt", "tgt1", "set1", "delivered",
            "gap", "progbasis", "proggap", "suggest"]
    df = pd.DataFrame(rows, columns=cols)
    for c in ("tgt", "tgt1", "gap", "progbasis", "proggap", "suggest"):
        df[c] = pd.Series([r.get(c) for r in rows], dtype=object)
    return df


def required_pricing_frame(view: dict, p: int) -> pd.DataFrame:
    """The required pricing leg, (1+target)/(1+rate leg) − 1 — a ratio of
    weighted means is NOT the weighted mean of the ratios, so a row
    populates ONLY where the filters resolve the state to exactly one
    NET combo; every other row is dashed rather than approximated (the
    workbook's own rule, sheets_book)."""
    recs = []
    for st in sorted(view):
        rows = view[st]
        rec = {"state": st}
        if (len(rows) == 1 and rows[0]["netp"] is not None
                and rows[0]["pf"] is not None):
            pf, x = rows[0]["pf"], rows[0]["netp"]
            for j in range(12):
                rl = pf.rows[j]["rate_leg"]
                rec[f"pp{j + 1:02d}"] = (1.0 + x) / (1.0 + rl) - 1.0
        else:
            for j in range(12):
                rec[f"pp{j + 1:02d}"] = None
        recs.append(rec)
    cols = ["state"] + [f"pp{j:02d}" for j in range(1, 13)]
    df = pd.DataFrame(recs, columns=cols)
    for c in cols[1:]:
        df[c] = pd.Series([r.get(c) for r in recs], dtype=object)
    return df
