"""The bridge chain — ONE definition, drawn two ways (P3).

``chain()`` is the single source of the bridge steps: the Combo card's
tables and the waterfall figures both read it, so the two can never
disagree about what the chain IS. ``waterfall_bars()`` turns a chain into
contiguous bar segments — pure data, tested on both interpreters —
and ``waterfall_figure()`` renders them with the invisible-riser pattern
(each step floats between its before/after running LR; only the endpoints
anchor at zero), the same trick the workbook's Excel charts use.
"""
from __future__ import annotations

from app.glue.format import fmt_pct


def chain(scn, key: str, res):
    """(start, steps_p, steps_p1, net) for one combo's bridge.

    Steps are (label, factor) with factors straight off the EngineResult;
    the start is the projected LR at current rate level. Labels carry no
    '× ' dressing — each surface adds its own. Labels within one chain must
    stay distinct (a categorical axis needs unique factors).
    """
    row = scn.row(key) or {}
    a_other = row.get("a_other")
    a_other = float(a_other) if isinstance(a_other, (int, float)) else 1.0
    trend = row.get("trend")
    trend = (float(trend) if isinstance(trend, (int, float))
             else scn.trend_default)
    net = row.get("netp") is not None or row.get("netp1") is not None
    steps_p = [("Rate earn-in (A_rate)", res.a_rate_p),
               ("Mod drift (A_mod)", res.a_mod_p),
               ("Other", a_other)]
    steps_p1 = [(f"Net trend ({fmt_pct(trend, signed=True)})", 1.0 + trend),
                ("Rate earn-in (A_rate, P+1)", res.a_rate_p1),
                ("Mod drift (A_mod, P+1)", res.a_mod_p1),
                ("Other", a_other)]
    return res.lr_current, steps_p, steps_p1, net


def waterfall_bars(start_label: str, start: float, steps: list,
                   total_label: str) -> list:
    """Chain -> contiguous waterfall segments.

    Each dict: cat / lo / hi (the bar's span), role (start|up|down|total),
    factor, pts (contribution in LR points), run (running LR after the
    step). Steps span [before, after]; endpoints anchor at zero.
    """
    bars = [dict(cat=start_label, lo=0.0, hi=start, role="start",
                 factor=None, pts=None, run=start)]
    run = start
    for label, f in steps:
        before, run = run, run * f
        bars.append(dict(cat=label, lo=min(before, run), hi=max(before, run),
                         role="up" if run >= before else "down",
                         factor=f, pts=(run - before) * 100.0, run=run))
    bars.append(dict(cat=total_label, lo=0.0, hi=run, role="total",
                     factor=None, pts=None, run=run))
    return bars


def waterfall_figure(title: str, bars: list):
    """Render ``waterfall_bars`` output as a raw-Bokeh categorical figure.

    Colors are the workbench waterfall trio (up steel, down navy, total
    green) with a slate start bar; hover carries the factor, the point
    contribution, and the running LR. Background transparent so the figure
    sits on either template theme.
    """
    from bokeh.models import ColumnDataSource, HoverTool, NumeralTickFormatter
    from bokeh.plotting import figure

    from app.glue.theme import chart_colors

    c = chart_colors()
    color_by = {"start": c["bar"], "up": c["up"], "down": c["down"],
                "total": c["total"]}
    src = ColumnDataSource(dict(
        cat=[b["cat"] for b in bars],
        lo=[b["lo"] for b in bars],
        hi=[b["hi"] for b in bars],
        color=[color_by[b["role"]] for b in bars],
        move=["" if b["factor"] is None else f"× {b['factor']:.4f}"
              for b in bars],
        pts=["" if b["pts"] is None else f"{b['pts']:+.1f} pts"
             for b in bars],
        run=[f"{b['run']:.1%}" for b in bars],
    ))
    fig = figure(x_range=[b["cat"] for b in bars], height=320,
                 sizing_mode="stretch_width", title=title,
                 toolbar_location=None, tools="")
    fig.vbar(x="cat", top="hi", bottom="lo", width=0.62, source=src,
             fill_color="color", line_color="color", fill_alpha=0.9)
    fig.add_tools(HoverTool(tooltips=[("step", "@cat"), ("move", "@move"),
                                      ("contribution", "@pts"),
                                      ("running LR", "@run")]))
    ink = c["actual"]
    fig.yaxis.formatter = NumeralTickFormatter(format="0%")
    fig.y_range.start = 0
    fig.xaxis.major_label_orientation = 0.7
    fig.xgrid.grid_line_color = None
    fig.background_fill_alpha = 0
    fig.border_fill_alpha = 0
    fig.title.text_color = ink
    for ax in (fig.xaxis, fig.yaxis):
        ax.major_label_text_color = ink
        ax.axis_line_color = c["ref"]
        ax.major_tick_line_color = c["ref"]
    return fig
