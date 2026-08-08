"""State Summary — the flagship exhibit, live (W2b).

The workbook's State Summary (D38/D90/D101/D104), mirrored from its own
SS_COLS declaration and aggregated the Book's way: one row per state
across every line in view, adjusted-plan-EP weighted, with the visible
bridge chain on single-combo rows and the exact weighted value + Mix on
mixed ones. Line and BU filters only — an exhibit keyed by a dimension
never filters on itself (the workbook doctrine).

W2c adds the levers (net targets, planned filings, mod steps) and the
undo stack; this wave the table is read-only.
"""
from __future__ import annotations

import panel as pn

from app import compute, summary
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.exhibit import banner, echo_html, echo_pane, exhibit_header, \
    note_list
from app.glue.format import TAB_SPECS

_BANNER = ("THE BRIDGE IN ONE LINE      CY plan LR  =  Projected LR "
           "(current level)  ×  rate earn-in  ×  mod drift  ×  other adj"
           "      — and those are literally the four columns to its left")

_NOTES = [
    "Weighted view: every metric is adjusted-plan-EP weighted across the "
    "lines and business units in view. A state with no EP shows — (the "
    "Book dialect; the per-LOB Excel exhibit falls back to a simple "
    "average — a deliberate, tested divergence).",
    "Plan LR is the product of the four factors to its left on rows that "
    "resolve to a single combo; where a row combines combos the exact "
    "EP-weighted engine value is shown and Mix carries the difference "
    "(factor averages do not compound across a mixed book). The TOTAL row "
    "never uses the chain.",
    "Rate change history is chronological and shows only where the "
    "filters resolve to a single combo: T = taken, P = planned, * = NOT "
    "counted in the indication; percents are filed % × achievement. "
    "— means the view mixes combos; blank means no such filing. The "
    "# taken / # planned counts work in every view.",
    "The following year is indicative — it multiplies the same Projected "
    "LR and Other adj by (1 + net trend).",
    "Net sel: the asserted net selection in force (— = explicit program); "
    "net combos carry the combined factor in A_rate with A_mod = 1.000. "
    "Program basis values the same rows on the logged program instead.",
    "Target LR weights over the combos that CARRY a target — the Excel "
    "exhibit dilutes it with the full EP of blank rows (a flaw its own "
    "source documents); the app fixes it and says so here.",
    "TOTAL rides directly under the headers (pinned). Columns marked "
    "under 'App extensions' exist only in the app, not in the Excel "
    "exhibit. ⚠ counts combos the engine rejected — they carry no weight "
    "anywhere in the row.",
]


def _formatters():
    from bokeh.models.widgets.tables import NumberFormatter
    return {k: NumberFormatter(format=TAB_SPECS[kind], nan_format="—",
                               null_format="—")
            for k, kind in summary.SS_KINDS.items()}


def build(session):
    bag = WatcherBag()

    # ---- sidebar -----------------------------------------------------------
    line_f = pn.widgets.MultiChoice(label="Lines", options=[],
                                    placeholder="all lines",
                                    sizing_mode="stretch_width")
    bu_f = pn.widgets.MultiChoice(label="Business units", options=[],
                                  placeholder="all BUs",
                                  sizing_mode="stretch_width")
    show_hist = pn.widgets.Checkbox(
        name="Show rate-change chronology", value=False)
    show_inputs = pn.widgets.Checkbox(
        name="Show bridge inputs (mods, engine levels)", value=True)

    def _bu_opts():
        vals = set()
        for scn in session.page.book.values():
            for k in scn.combo_keys():
                vals.add(k.split("|", 1)[0])
        return sorted(vals)

    sync_options(line_f, lambda: list(session.page.book),
                 session.ctx.param.data_rev, bag=bag)
    sync_options(bu_f, _bu_opts, session.ctx.param.data_rev, bag=bag)

    # ---- main --------------------------------------------------------------
    header = exhibit_header(
        "State Summary — every line, one row per state",
        "Adjusted-EP weighted across the lines and business units in view; "
        "the aggregation is live — an edit anywhere moves it. Rate-change "
        "history shows only where the filters resolve to a single combo — "
        "the workbook's honesty rule, one dimension deeper.")
    echo = echo_pane()

    p0 = next((s.plan_year for s in session.page.book.values()), 0)
    titles = {**{c.key: c.header for c in summary.SS_COLS},
              **summary.live_headers(p0), **dict(summary.APP_EXT)}
    hidden = summary.hidden_sets()

    import pandas as pd
    table = pn.widgets.Tabulator(
        pd.DataFrame(columns=summary.ALL_KEYS),
        show_index=False,
        titles=titles,
        groups=summary.ss_groups_map(p0),
        formatters=_formatters(),
        editors={k: None for k in summary.ALL_KEYS},   # W2b: all locked
        frozen_columns=["state"],                      # by NAME (the scar)
        frozen_rows=[-1],                              # TOTAL pinned (top)
        text_align={k: "right" for k in summary.SS_KINDS},
        hidden_columns=["_index"] + hidden["hist"],
        layout="fit_data_table", height=560,
        configuration={"clipboard": "copy"},           # NO "columns" key —
        # configuration["columns"] cannot coexist with `groups` (Panel
        # raises); the W2c editors gate server-side instead
        sizing_mode="stretch_width")

    holder = {"meta": None, "frame": None}

    def _hidden(*_events):
        cols = ["_index"]
        if not show_hist.value:
            cols += hidden["hist"]
        if not show_inputs.value:
            cols += hidden["inputs"]
        table.hidden_columns = cols

    show_hist.param.watch(_hidden, "value")
    show_inputs.param.watch(_hidden, "value")

    def _render(*_events):
        if not session.page.book:
            table.value = table.value.iloc[0:0]
            echo.object = echo_html(
                "No lines loaded — Book page → Load every line, paste "
                "masters, or load a scenario.")
            return
        df, meta = summary.state_frame(session, lines=line_f.value or None,
                                       bus=bu_f.value or None)
        holder["meta"], holder["frame"] = meta, df
        p = meta.plan_year
        table.titles = {**{c.key: c.header for c in summary.SS_COLS},
                        **summary.live_headers(p), **dict(summary.APP_EXT)}
        if table.groups != summary.ss_groups_map(p):
            # the page builds before any book exists (plan year 0) — the
            # year-bearing group captions refresh on the first real render
            table.groups = summary.ss_groups_map(p)
        table.value = df
        table.style = df.style.apply(summary.ss_styles, axis=None)
        lines_txt = (", ".join(line_f.value) if line_f.value
                     else "every line")
        bus_txt = (", ".join(bu_f.value) if bu_f.value
                   else "every business unit")
        echo.object = echo_html(
            f"Showing: {lines_txt}  |  {bus_txt}  |  "
            f"{sum(meta.cnt.values())} combos across {len(meta.states)} "
            f"states")

    rail_rev = debounce(session.bus.param.rev, delay_ms=300, bag=bag)
    bag.watch(rail_rev, _render, "rev")
    bag.watch(session.ctx, _render, "data_rev")
    for f in (line_f, bu_f):
        bag.watch(f, _render, "value")
    _render()

    sidebar = pn.Column(
        pn.pane.Markdown("**Filters** — empty means everything"),
        line_f, bu_f,
        pn.layout.Divider(),
        pn.pane.Markdown("**Columns** (the workbook's collapse outline)"),
        show_hist, show_inputs,
        sizing_mode="stretch_width")
    main = pn.Column(header, echo, table, banner(_BANNER),
                     note_list(_NOTES),
                     sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "table": table, "holder": holder,
            "filters": {"line": line_f, "bu": bu_f},
            "toggles": {"hist": show_hist, "inputs": show_inputs},
            "on_show": _render}
