"""Earning geometry — the parallelogram and the earn-in lag (P3).

Every number here is the ENGINE's: parallelogram cells come from
``MonthlyEngine.e`` / ``w`` / ``index_in_force``, and the lag chart reads
the monthly series on ``EngineResult`` that the workbook's oracle harness
ties cell-by-cell. The frames are pure pandas so the numbers test on both
interpreters; the chart builders load holoviews lazily.
"""
from __future__ import annotations

import pandas as pd

from src.engine import mi_month, mi_year, month_index

from app.glue.theme import GREY_LINE, NAVY, STEEL, STEEL_LIGHT
from . import ensure_hv


def _ym(mi: int) -> str:
    """Lexically sortable month label ('2027-04') — categorical axes then
    order themselves chronologically for free."""
    return f"{mi_year(mi)}-{mi_month(mi):02d}"


def _mix(a: str, b: str, t: float) -> str:
    va = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    vb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}"
                         for x, y in zip(va, vb))


# the workbook's D112 scale: white (earns nothing) -> steel (peak earning)
WHITE_TO_STEEL = [_mix("#FFFFFF", STEEL, i / 31) for i in range(32)]


def parallelogram_frame(eng, plan_year: int) -> pd.DataFrame:
    """Cohort × month earned-share cells over Jan P .. Dec P+1.

    Rows are the cohorts that earn ANYTHING in the window; zero cells are
    left absent — the blank space is what draws the parallelogram's slope.
    """
    m0 = month_index(plan_year, 1)
    months = list(range(m0, m0 + 24))
    recs = []
    for k in eng.cohorts:
        if k + eng.t < m0 or k > months[-1]:
            continue
        idx = eng.index_in_force(k)
        wk = eng.w(k)
        for m in months:
            share = eng.e(k, m)
            if share:
                recs.append(dict(
                    cohort=_ym(k), month=_ym(m), share=share, weight=wk,
                    index=idx, share_txt=f"{share:.1%}",
                    index_txt=f"{idx:.4f}"))
    return pd.DataFrame(recs)


def parallelogram(eng, plan_year: int):
    """The earning parallelogram: written cohorts (rows) releasing premium
    into calendar months (columns), shaded by the share released."""
    hv = ensure_hv()

    df = parallelogram_frame(eng, plan_year)
    n_rows = df["cohort"].nunique()
    hm = hv.HeatMap(df, kdims=["month", "cohort"],
                    vdims=["share", "weight", "index",
                           "share_txt", "index_txt"])
    return hm.opts(
        cmap=WHITE_TO_STEEL, colorbar=False, responsive=True,
        height=min(640, max(320, 15 * n_rows + 110)), invert_yaxis=True,
        xrotation=90, xlabel="earning month", ylabel="written cohort",
        line_color=GREY_LINE, line_width=0.5,
        title=(f"Earning parallelogram — share of each written cohort "
               f"earned by month (term {eng.t} mo)"),
        fontsize={"title": "9pt"}, tools=["hover"],
        hover_tooltips=[("written", "@cohort"), ("earns in", "@month"),
                        ("share of cohort", "@share_txt"),
                        ("index in force", "@index_txt")])


def earn_lag_frame(res) -> pd.DataFrame:
    """24 months Jan P .. Dec P+1: the index the month's NEW cohort carries
    (written, in force) vs the earned level of the whole book — read
    straight off the EngineResult series."""
    m0 = month_index(res.plan_year, 1)
    rows = []
    for i in range(24):
        mi = m0 + i
        ts = pd.Timestamp(mi_year(mi), mi_month(mi), 1)
        rows.append(dict(
            month=ts, label=f"{ts:%b %Y}",
            written=res.inforce_index[res.cohort_mi.index(mi)],
            earned=res.earned_index_m[res.earned_month_mi.index(mi)]))
    df = pd.DataFrame(rows)
    df["written_txt"] = df["written"].map(lambda v: f"{v:.4f}")
    df["earned_txt"] = df["earned"].map(
        lambda v: "—" if pd.isna(v) else f"{v:.4f}")
    return df


def earn_lag(res):
    """Written level vs earned level — the lag the whole methodology is
    about. The area is what the book has actually banked; the navy line is
    what the newest cohort is being written at."""
    hv = ensure_hv()

    df = earn_lag_frame(res)
    p = res.plan_year
    area = hv.Area(df, "month", "earned").opts(
        fill_color=STEEL_LIGHT, fill_alpha=0.75, line_width=0,
        show_legend=False)
    earned = hv.Curve(df, "month", ["earned", "label", "earned_txt"]).relabel(
        "Earned level").opts(
        color=STEEL, line_width=2, tools=["hover"],
        hover_tooltips=[("month", "@label"), ("earned level", "@earned_txt")])
    written = hv.Curve(df, "month",
                       ["written", "label", "written_txt"]).relabel(
        "Written level (in force)").opts(
        color=NAVY, line_width=2.5, tools=["hover"],
        hover_tooltips=[("month", "@label"),
                        ("written level", "@written_txt")])
    return (area * earned * written).opts(
        responsive=True, height=320, ylabel="rate index", xlabel="",
        legend_position="top_left",
        title=(f"Written vs earned level, Jan {p} – Dec {p + 1} "
               "(the earn-in lag)"),
        fontsize={"title": "9pt"})
