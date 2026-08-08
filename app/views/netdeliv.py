"""The Net Delivery microscope (W3f) — the solve surface, one combo.

The workbook's split, kept: the Flow page REPORTS delivery, the Summary
EDITS the targets, and this — the Combo page's Delivery tab — PRESCRIBES:
given the net selection the combo asserts, what filing (and what mod
action) actually closes the gap. Everything is the engine's own closed
forms — ``suggest_net_rate``, ``net_delivery_by_month`` (+ the D102 P+1
carryover twin), ``required_m1``, ``required_mod_step`` — arranged,
never re-derived; feasibility flags and warnings surface verbatim.

Defaults follow the Excel tab: the effective date adopts the FIRST
planned filing in the plan year (flagged MULTI-PLANNED when there are
several), else April 1; the change in force adopts that filing's
effective percent, else the suggested change; a typed override wins.
"""
from __future__ import annotations

import datetime as dt
import math

import pandas as pd

from src import engine
from src.engine import mi_month, mi_year

from app.glue.theme import FAIL_RED, NAVY, STEEL
from . import ensure_hv


def _label(mi: int) -> str:
    return f"{pd.Timestamp(mi_year(mi), mi_month(mi), 1):%b %Y}"


def microscope(p: int, combo, eff_date: dt.date | None = None,
               filed_override: float | None = None) -> dict:
    """One combo under the microscope at a resolved (date, change).

    Returns the 12-row decomposition frame (with the log-scaled shares
    that sum exactly to the target), the P+1 carryover twin and which
    target it used, the suggested change, the required M_1' level (or
    the reason it reads n/a), the required mod STEP (D75), and every
    engine warning verbatim. Raises ValueError with the engine's own
    words when the date cannot work (outside the plan year, no exposure
    after it) — the page shows the words, never a blank tab.
    """
    planned = engine.planned_change_in_plan_year(p, combo)
    if eff_date is None:
        eff_date = planned[0] if planned else dt.date(p, 4, 1)
    suggested, _comp = engine.suggest_net_rate(p, combo, eff_date)
    ach = 1.0
    if planned and planned[2]:
        ach = planned[1] / planned[2] if planned[2] else 1.0
    if filed_override is not None:
        r, source = filed_override, "override"
    elif planned and eff_date == planned[0]:
        r, source = planned[1], f"planned filing {planned[0]}"
    else:
        r, source = suggested, "suggested"

    x = combo.net_sel_p
    rows = engine.net_delivery_by_month(p, combo, eff_date, r)
    ln1x = math.log(1.0 + x) if x is not None and x > -1.0 and x != 0.0 \
        else None

    def shares(rr):
        if ln1x is None or rr["price_leg_required"] is None:
            return None, None
        sr = x * math.log(1.0 + rr["rate_leg"]) / ln1x
        sp = x * math.log(1.0 + rr["price_leg_required"]) / ln1x
        return sr, sp

    recs = []
    for rr in rows:
        sr, sp = shares(rr)
        recs.append(dict(
            label=_label(rr["mi"]), w=rr["w"], rate_leg=rr["rate_leg"],
            price_req=rr["price_leg_required"],
            m_required=rr["m_required"], m_proj=rr["m_proj"],
            gap=(None if rr["m_required"] is None
                 else (rr["m_required"] - rr["m_proj"]) * 100.0),
            delivered=rr["delivered_at_proj"],
            share_rate=sr, share_price=sp))
    df = pd.DataFrame(recs)
    for c in ("price_req", "m_required", "gap", "share_rate",
              "share_price"):
        df[c] = pd.Series([r.get(c) for r in recs], dtype=object)

    p1 = engine.net_delivery_by_month_p1(p, combo, eff_date, r)
    p1_recs = [dict(label=_label(rr["mi"]), w=rr["w"],
                    rate_leg=rr["rate_leg"],
                    price_req=rr["price_leg_required"],
                    m_required=rr["m_required"], m_base=rr["m_base"])
               for rr in p1]
    p1_df = pd.DataFrame(p1_recs)
    for c in ("price_req", "m_required", "m_base"):
        p1_df[c] = pd.Series([r.get(c) for r in p1_recs], dtype=object)
    p1_src = ("entered" if combo.net_sel_p1 is not None
              else "carried from P")

    m1p, m1p_reason = None, None
    if not combo.mod_adjustment_enabled:
        m1p_reason = "n/a — the mod adjustment is off (rate-only)"
    elif combo.mod_changes:
        m1p_reason = "n/a — the Mod Log pins the path (act via a step)"
    else:
        try:
            m1p = engine.required_m1(p, combo, eff_date, r)
        except ValueError as e:
            m1p_reason = str(e)

    step, step_reason = None, None
    if not combo.mod_adjustment_enabled:
        step_reason = ("no mod ask — delivery carries no mod leg, the "
                       "target is interpreted rate-only")
    else:
        try:
            step = engine.required_mod_step(p, combo, eff_date, r,
                                            achievement=ach)
        except ValueError as e:
            step_reason = str(e)

    warnings = list(getattr(_comp, "warnings", ()) or ())
    if step is not None:
        warnings += list(step.warnings or ())
    return dict(eff_date=eff_date, r=r, source=source, ach=ach,
                filed_equiv=r / ach if ach else r,
                suggested=suggested, x=x, x1=p1[0]["x"] if p1 else None,
                rows=df, p1=p1_df, p1_src=p1_src,
                m1_prime=m1p, m1_reason=m1p_reason,
                step=step, step_reason=step_reason,
                multi_planned=planned[3] if planned else 0,
                warnings=warnings)


def share_chart(df: pd.DataFrame, x: float, p: int):
    """How each month reaches the target: rate share + pricing share,
    log-scaled so the two stack EXACTLY to x, with the delivered-at-
    projected-mods line and the target line over the top. None when the
    shares are undefined (mod-off / zero target)."""
    if df["share_rate"].iloc[0] is None:
        return None
    hv = ensure_hv()
    recs = []
    for _i, r in df.iterrows():
        recs.append((r["label"], "rate share", r["share_rate"]))
        recs.append((r["label"], "pricing share", r["share_price"]))
    bars = hv.Bars(recs, kdims=["label", "kind"], vdims=["share"]).opts(
        stacked=True, cmap={"rate share": STEEL, "pricing share": NAVY},
        color="kind", xrotation=90, show_legend=True,
        legend_position="top_left", tools=["hover"])
    d = df.assign(del_txt=df["delivered"].map(lambda v: f"{v:+.2%}"))
    deliv = hv.Curve(d, "label", ["delivered", "del_txt"]).relabel(
        "Delivered @ proj mods").opts(
        color=NAVY, line_width=2.4, tools=["hover"],
        hover_tooltips=[("month", "@label"),
                        ("delivered", "@del_txt")])
    target = hv.HLine(x).opts(color=FAIL_RED, line_dash="dashed",
                              line_width=1.6)
    return (bars * deliv * target).opts(
        responsive=True, height=320, xlabel="",
        ylabel="share of target / delivered",
        title=(f"How the month gets to the target — rate + pricing "
               f"shares (log-scaled to sum to {x:+.1%})"),
        fontsize={"title": "9pt"})
