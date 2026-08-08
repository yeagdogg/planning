"""State Summary — the flagship exhibit, live and editable (W2b/W2c).

The workbook's State Summary (D38/D90/D101/D104), mirrored from its own
SS_COLS declaration and aggregated the Book's way: one row per state
across every line in view, adjusted-plan-EP weighted, with the visible
bridge chain on single-combo rows and the exact weighted value + Mix on
mixed ones. Line and BU filters only — an exhibit keyed by a dimension
never filters on itself (the workbook doctrine).

W2c makes it the app's what-if surface: net targets fan out to every
combo in the row's view, planned filings and mod steps edit in place on
single-combo rows (taken = historical fact, locked), every gesture lands
on the undo stack, and the recompute patches only the cells that moved
(the workbench diff-then-patch idiom — scroll and filters survive). All
gating is SERVER-side: this table carries Panel `groups`, which cannot
coexist with configuration["columns"], so rejects revert the cell via
patch and say why.
"""
from __future__ import annotations

import re

import panel as pn

from app import compute, summary
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.exhibit import banner, echo_html, echo_pane, exhibit_header, \
    note_list
from app.glue.format import TAB_SPECS, fmt_pct, md_safe
from app.glue.theme import GREY_LINE, NAVY, TABLE_CSS
from app.paste import coerce
from app.undo import UndoStack, entry

_BANNER = ("THE BRIDGE IN ONE LINE      CY plan LR  =  Projected LR "
           "(current level)  ×  rate earn-in  ×  mod drift  ×  other adj"
           "      — the four factors ride in the bridge band; the answer "
           "stays pinned at the right edge, always in view")

# The two notes that change BEHAVIOR stay inline under the table (W3c);
# everything else folds into the fine-print card.
_NOTE_LEVERS = (
    "LEVERS (blue, STEEL-edged): net selections fan out to EVERY combo in "
    "the row's view — blank CLEARS net mode, a methodology flip the flash "
    "names; filings/mod/achievement edit the log on single-combo views; "
    "taken filings are locked. ° columns exist only in the app. Every "
    "gesture lands on ↶ Undo.")
_NOTE_CHAIN = (
    "Plan LR is the four-factor chain on single-combo rows; mixed rows "
    "show the exact EP-weighted engine value with Mix carrying the "
    "difference. The TOTAL row never uses the chain.")

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
    "LEVERS (blue-tinted cells): Net rate sel / +1 apply to EVERY combo "
    "in the row's current view (blank clears net mode — a methodology "
    "flip, and the flash says so); planned filings' % and date, the mod "
    "step, and achievement edit the underlying log row on single-combo "
    "views. Taken filings are historical fact and locked. Percents edit "
    "as fractions or with a % sign; every gesture lands on ↶ Undo.",
]

# One-sentence hovers for the dense columns (the notes, distilled).
_TOOLTIPS = {
    "planlr": "Product of the four bridge factors on single-combo rows; "
              "the exact EP-weighted engine value on mixed rows (Mix "
              "carries the difference).",
    "planlr1": "Indicative: the same Projected LR and Other adj x "
               "(1 + net trend).",
    "mix": "Weighted plan LR minus the four-factor chain, in points — "
           "factor averages do not compound across a mixed book.",
    "target": "EP-weighted over the combos that CARRY a target (the Excel "
              "exhibit dilutes with blank rows; the app fixes it).",
    "progbasis": "The same rows valued on the LOGGED rate program instead "
                 "of the net assertion (D65).",
    "proggap": "Program basis minus asserted plan LR, in points.",
    "netsel": "Asserted net selection — applies to EVERY combo in the "
              "row's view; blank clears net mode.",
    "netsel1": "Net selection for the following year (blank carries P's).",
    "ach_next": "Achievement on the next planned filing (blank = 100%).",
    "crl": "Cumulative rate level in the indication.",
    "ecy": "Calendar-year earned rate level.",
    "mbar": "Earned schedule mod for the year.",
}

_LEVER_NUM = ("netsel", "netsel1")
_LEVER_TXT = tuple(f"chg{j}_{p}" for j in (1, 2, 3, 4)
                   for p in ("date", "pct")) + \
    ("modstep_date", "modstep_pct", "ach_next")
LEVERS = _LEVER_NUM + _LEVER_TXT


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
    editors = {k: None for k in summary.ALL_KEYS}      # computed = locked
    editors.update({k: {"type": "number", "step": 0.001}
                    for k in _LEVER_NUM})
    editors.update({k: {"type": "input"} for k in _LEVER_TXT})
    table = pn.widgets.Tabulator(
        pd.DataFrame(columns=summary.DISPLAY_ORDER),
        show_index=False,
        titles=titles,
        groups=summary.display_groups_map(p0),
        formatters=_formatters(),
        editors=editors,
        # dict form freezes by NAME on a SIDE; planlr/planlr1 are the
        # answer and ride the right edge (they are deliberately ungrouped
        # — Tabulator ignores frozen inside a column group). Any value
        # other than left/right silently VANISHES the column — pinned by
        # a test.
        frozen_columns={"state": "left",
                        "planlr": "right", "planlr1": "right"},
        # frozen_rows is set per roster in _apply_frame: a negative index
        # here would resolve ONCE against this empty frame and the TOTAL
        # pin would silently never work (W3a — found broken since W2b)
        text_align={k: "right" for k in summary.SS_KINDS},
        hidden_columns=["_index"] + hidden["hist"],
        layout="fit_data_table",
        # grow to the roster, capped: 21 states + TOTAL must sit on screen
        # without an inner scrollbar (a fixed 560 showed 16 of 22 rows)
        max_height=840,
        selectable=False,              # selection is never read here, and
        # the fast theme's selected-row repaint whited out the lever cells
        stylesheets=[TABLE_CSS],
        header_tooltips=_TOOLTIPS,
        configuration={"clipboard": "copy",            # NO "columns" key —
                       # configuration["columns"] cannot coexist with
                       # `groups` (Panel raises); columnDefaults is legal
                       # beside groups and lets the long headers wrap
                       # instead of forcing 200px columns
                       "columnDefaults": {"headerWordWrap": True}},
        sizing_mode="stretch_width")

    flash = pn.pane.Markdown("", sizing_mode="stretch_width")
    holder = {"meta": None, "frame": None}
    undo = UndoStack()

    # ---- the last-edit impact strip (W3c) ----------------------------------
    # "hard to see what my change does to the plan loss ratios": one
    # snapshot per accepted gesture, 1:1 with the undo stack — the answer
    # cells for the edited state + TOTAL, before vs after the recompute.
    impact = pn.pane.HTML("", sizing_mode="stretch_width")
    strip_stack: list = []

    def _impact_before(st, desc):
        """Read the pre-recompute frame still sitting in table.value: the
        lever cell is already updated there, but planlr/planlr1/TOTAL move
        only when the engine has spoken — exactly the 'before'."""
        df = table.value
        try:
            row = df.loc[df["state"] == st].iloc[0]
            tot = df.iloc[-1]
        except (IndexError, KeyError):
            return
        strip_stack.append({
            "desc": desc, "state": st, "p": None,
            "before": (row["planlr"], row["planlr1"],
                       tot["planlr"], tot["planlr1"]),
            "after": None})
        del strip_stack[:-20]                    # the undo stack's own cap

    def _impact_fill(df, p):
        for snap in strip_stack:
            if snap["after"] is not None:
                continue
            sel = df.loc[df["state"] == snap["state"]]
            if not len(sel) or not len(df) \
                    or df.iloc[-1]["state"] != "TOTAL":
                snap["after"] = "stale"          # roster moved under it
                continue
            row, tot = sel.iloc[0], df.iloc[-1]
            snap["after"] = (row["planlr"], row["planlr1"],
                             tot["planlr"], tot["planlr1"])
            snap["p"] = p
        top = strip_stack[-1] if strip_stack else None
        impact.object = ("" if top is None or top["after"] in (None, "stale")
                         else _impact_html(top))

    def _impact_html(snap):
        import html as _html
        import math

        def f(v):
            bad = v is None or (isinstance(v, float) and math.isnan(v))
            return "—" if bad else fmt_pct(v)

        def pair(b, a):
            ok = (isinstance(b, (int, float)) and isinstance(a, (int, float))
                  and not (math.isnan(b) or math.isnan(a)))
            if ok:
                d = (a - b) * 100.0
                arrow = "▼" if d < 0 else ("▲" if d > 0 else "＝")
                return f"{f(b)} → {f(a)} ({arrow} {abs(d):.1f})"
            return f"{f(b)} → {f(a)}"

        p = snap["p"]
        b, a = snap["before"], snap["after"]
        parts = [
            f"<b>Last change</b> — {_html.escape(str(snap['desc']))}",
            f"<b>{_html.escape(str(snap['state']))}</b> "
            f"CY{p} {pair(b[0], a[0])} · CY{p + 1} {pair(b[1], a[1])}",
            f"<b>TOTAL</b> CY{p} {pair(b[2], a[2])} · "
            f"CY{p + 1} {pair(b[3], a[3])}",
        ]
        return (f"<div style='border:1px solid {GREY_LINE}; "
                "border-radius:8px; "
                "padding:6px 12px; font-variant-numeric:tabular-nums; "
                f"font-size:0.95em; color:{NAVY};'>"
                + "&nbsp; │ &nbsp;".join(parts) + "</div>")

    def _hidden(*_events):
        cols = ["_index"]
        if not show_hist.value:
            cols += hidden["hist"]
        if not show_inputs.value:
            cols += hidden["inputs"]
        table.hidden_columns = cols

    show_hist.param.watch(_hidden, "value")
    show_inputs.param.watch(_hidden, "value")

    def _style_fn(d):
        return summary.ss_styles(d, levers=LEVERS)

    def _apply_frame(df):
        """The workbench diff-then-patch idiom: when the roster is
        unchanged, patch only the cells that moved — scroll, filters and
        the open editor state survive; the Styler recomputes on patch.
        The frame arrives in ALL_KEYS order (the D90 mirror); the DISPLAY
        band order is applied here, before both the value and the Styler,
        so styles stay aligned by construction."""
        import pandas as pd
        df = df[summary.DISPLAY_ORDER]
        old = table.value
        if (old is None or len(old) != len(df)
                or list(old["state"]) != list(df["state"])):
            table.value = df
            table.style = df.style.apply(_style_fn, axis=None)
            # the pin is POSITIONAL and Panel resolves negative indices only
            # once, against whatever length the frame had at set time — so
            # re-pin TOTAL by its real index on every roster change (W3a)
            table.frozen_rows = [len(df) - 1] if len(df) else []
            return
        patches = {}
        for c in df.columns:
            if c == "state":
                continue
            ch = []
            for i, (a, b) in enumerate(zip(old[c], df[c])):
                try:
                    same = (a == b) or (pd.isna(a) and pd.isna(b))
                except (TypeError, ValueError):
                    same = False
                if not same:
                    ch.append((i, b))
            if ch:
                patches[c] = ch
        if patches:
            table.patch(patches)

    def _render(*_events):
        if not session.page.book:
            table.value = table.value.iloc[0:0]
            table.frozen_rows = []
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
        if table.groups != summary.display_groups_map(p):
            # the page builds before any book exists (plan year 0) — the
            # year-bearing group captions refresh on the first real render
            table.groups = summary.display_groups_map(p)
        _apply_frame(df)
        _impact_fill(df, p)
        lines_txt = (", ".join(line_f.value) if line_f.value
                     else "every line")
        bus_txt = (", ".join(bu_f.value) if bu_f.value
                   else "every business unit")
        echo.object = echo_html(
            f"Showing: {lines_txt}  |  {bus_txt}  |  "
            f"{sum(meta.cnt.values())} combos across {len(meta.states)} "
            f"states")

    # ---- the edit router (W2c): every gate is server-side ------------------
    def _reject(event, reason):
        table.patch({event.column: [(event.row, event.old)]})
        flash.object = f"⚠ **{md_safe(reason)}** — nothing was changed."

    def _frac(raw):
        """Editor input -> fraction or None (blank). Accepts 0.05, 5%,
        +5.0% — the paste vocabulary."""
        if raw is None or raw == "":
            return None
        if isinstance(raw, (int, float)):
            return float(raw)
        return coerce("pct", str(raw))

    def _route_edit(event):
        meta = holder["meta"]
        df = table.value
        if meta is None or df is None or not (0 <= event.row < len(df)):
            return
        st = str(df.iloc[event.row]["state"])
        col = event.column
        if col not in LEVERS:
            return _reject(event, f"'{col}' is calculated, not entered")
        if st == "TOTAL":
            return _reject(event, "the TOTAL row is an aggregate — edit a "
                                  "state row")
        if meta.problems.get(st):
            return _reject(event, f"{st} carries combos the engine "
                                  "rejected — fix them on Inputs first")

        # -- net targets: fan out to every combo in the row's view ----------
        if col in _LEVER_NUM:
            field_ = "netp" if col == "netsel" else "netp1"
            try:
                new = _frac(event.value)
            except ValueError as e:
                return _reject(event, str(e))
            combos = meta.view.get(st, [])
            entries, names = [], []
            for lob, _scn, key, _bu, row, _res in combos:
                old = row.get(field_)
                if old == new:
                    continue
                entries.append(entry(lob, "lr_rows", row, {field_: old},
                                     after={field_: new}))
                row[field_] = new
                names.append(key)
            if not entries:
                flash.object = "*No change — every combo already there.*"
                return
            label = ("Net rate P" if field_ == "netp" else "Net rate P+1")
            what = ("cleared (net mode OFF)" if new is None
                    else f"→ {new:+.1%}")
            desc = f"{label} {what} on {st} ({len(entries)} combo(s))"
            _impact_before(st, desc)
            undo.push(desc, entries)
            _undo_sync()
            flip = (" — **this flips those combos out of net mode**"
                    if new is None else "")
            flash.object = (f"✓ **{label} {md_safe(what)}** applied to "
                            f"{len(entries)} combo(s): "
                            f"{md_safe(', '.join(names[:6]))}"
                            f"{' …' if len(names) > 6 else ''}{flip} — ↶ "
                            f"Undo is armed.")
            session.bus.bump()
            return

        # -- everything below edits ONE underlying log row -------------------
        if meta.cnt.get(st, 0) != 1:
            return _reject(event, "log edits need the filters to resolve "
                                  "to a single combo — pick a line and BU")
        lob = meta.single[st][0]

        m = re.match(r"chg(\d)_(date|pct)$", col)
        if m:
            j, part = int(m.group(1)), m.group(2)
            r = meta.slot_map.get((st, j))
            if r is None:
                return _reject(event, "that slot holds no filing — add "
                                      "one on Inputs")
            if r.get("status") != "planned":
                return _reject(event, "taken filings are historical fact "
                                      "— only planned rows are editable")
            if part == "date":
                try:
                    d = coerce("date", event.value)
                except ValueError as e:
                    return _reject(event, str(e))
                if d is None:
                    return _reject(event, "a planned filing needs a date "
                                          "(delete rows on Inputs)")
                desc = f"Chg {j} date {r['eff']} → {d} ({st})"
                _impact_before(st, desc)
                undo.push(desc,
                          [entry(lob, "rate_rows", r, {"eff": r["eff"]},
                                 after={"eff": d})])
                r["eff"] = d
                flash.object = f"✓ **Chg {j} effective → {d}** — ↶ armed."
            else:
                try:
                    v = _frac(event.value)
                except ValueError as e:
                    return _reject(event, str(e))
                if v is None:
                    return _reject(event, "blank would delete the filing "
                                          "— do that on Inputs")
                desc = f"Chg {j} filed % → {v:+.1%} ({st})"
                _impact_before(st, desc)
                undo.push(desc,
                          [entry(lob, "rate_rows", r,
                                 {"filed": r.get("filed")},
                                 after={"filed": v})])
                r["filed"] = v
                flash.object = (f"✓ **Chg {j} Filed % → {v:+.1%}** — the "
                                f"slot shows filed × achievement — ↶ armed.")
            _undo_sync()
            session.bus.bump()
            return

        if col in ("modstep_date", "modstep_pct"):
            got = meta.modstep_map.get(st)
            if got is None:
                return _reject(event, "no planned mod action for this "
                                      "combo — add one on Inputs")
            lob2, r = got
            if col == "modstep_date":
                try:
                    d = coerce("date", event.value)
                except ValueError as e:
                    return _reject(event, str(e))
                if d is None:
                    return _reject(event, "a mod step needs a date")
                desc = f"Mod step date {r['eff']} → {d} ({st})"
                _impact_before(st, desc)
                undo.push(desc,
                          [entry(lob2, "mod_rows", r, {"eff": r["eff"]},
                                 after={"eff": d})])
                r["eff"] = d
            else:
                try:
                    v = _frac(event.value)
                except ValueError as e:
                    return _reject(event, str(e))
                if v is None:
                    return _reject(event, "blank would delete the action "
                                          "— do that on Inputs")
                desc = f"Mod step → {v:+.1%} ({st})"
                _impact_before(st, desc)
                undo.push(desc,
                          [entry(lob2, "mod_rows", r,
                                 {"chg": r.get("chg")},
                                 after={"chg": v})])
                r["chg"] = v
            _undo_sync()
            flash.object = "✓ **Mod step updated** — ↶ Undo is armed."
            session.bus.bump()
            return

        if col == "ach_next":
            got = meta.achnext_map.get(st)
            if got is None:
                return _reject(event, "no planned filing in the plan year "
                                      "for this combo")
            lob2, r = got
            try:
                v = _frac(event.value)
            except ValueError as e:
                return _reject(event, str(e))
            desc = (f"Achievement → "
                    f"{'blank (=100%)' if v is None else f'{v:.0%}'} "
                    f"({st})")
            _impact_before(st, desc)
            undo.push(desc,
                      [entry(lob2, "rate_rows", r,
                             {"achievement": r.get("achievement")},
                             after={"achievement": v})])
            r["achievement"] = v
            _undo_sync()
            flash.object = ("✓ **Achievement updated** (blank = 100%) — ↶ "
                            "Undo is armed.")
            session.bus.bump()
            return

    table.on_edit(_route_edit)

    # ---- undo (inline with the flash — the whole edit loop in one place)
    undo_btn = pn.widgets.Button(label="↶ Undo", disabled=True, width=280)

    def _undo_sync():
        d = undo.peek()
        undo_btn.label = f"↶ Undo: {d}" if d else "↶ Undo"
        undo_btn.disabled = d is None

    def _do_undo(_event):
        n, skipped = undo.pop_apply(session)          # bumps the bus itself
        if strip_stack:
            strip_stack.pop()           # 1:1 with the undo stack — the
            # render that follows repaints the strip from the new top,
            # whose numbers ARE the current frame again
        msg = f"restored {n} row(s)"
        if skipped:
            msg += " — " + "; ".join(skipped[:3])
        flash.object = (("⚠ " if skipped else "↶ ") + f"**{md_safe(msg)}**")
        _undo_sync()

    undo_btn.on_click(_do_undo)

    def _on_data_rev(*_events):
        undo.clear()                    # locators dangle across a load
        strip_stack.clear()
        impact.object = ""
        _undo_sync()

    bag.watch(session.ctx, _on_data_rev, "data_rev")

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
        pn.layout.Divider(),
        pn.pane.Markdown(
            "*Undo covers Summary edits since the last load — scenario "
            "files are the durable history.*"),
        sizing_mode="stretch_width")
    fine_print = pn.Card(
        note_list(_NOTES),
        title="The fine print — aggregation, divergences, edit rules",
        collapsed=True, sizing_mode="stretch_width")
    main = pn.Column(header, echo, banner(_BANNER),
                     pn.Row(undo_btn, flash,
                            sizing_mode="stretch_width"),
                     impact, table,
                     note_list([_NOTE_LEVERS, _NOTE_CHAIN]),
                     fine_print,
                     sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "table": table, "holder": holder, "flash": flash,
            "filters": {"line": line_f, "bu": bu_f},
            "toggles": {"hist": show_hist, "inputs": show_inputs},
            "edit_router": _route_edit, "undo": undo,
            "undo_btn": undo_btn,
            "impact": {"pane": impact, "stack": strip_stack},
            "fine_print": fine_print,
            "on_show": _render}
