"""Inputs — where the book STARTS (P2, promoted W3g).

The masters live here now: three paste boxes spanning EVERY line (the
per-line columns with a leading Line column), working on an empty book —
``masters.apply_master`` takes the session and can create lines, which
makes this page the true cold start John described ("this app will
always be the starting point to paste our info into"). Below them, the
four per-line grids edit the ACTIVE line only, and the context bar says
which line that is — two paste surfaces with near-identical column
orders must be unmistakable.

Flow (the workbench contract): grid edit -> ``value`` watcher
(authoritative write-back) -> scenario rows replaced -> bus.bump() ->
every derived surface recomputes off ONE signal. Reseeds (scenario load,
paste apply) run through the same watcher and are swallowed by
rows_equal, so echo loops are structurally impossible.
"""
from __future__ import annotations

import panel as pn

from app import masters, validate as validate_mod
from app.glue.bindings import WatcherBag, debounce
from app.glue.exhibit import echo_html, exhibit_header, html_pane, \
    note_list
from app.glue.format import md_safe
from app.grids import (LR_SCHEMA, ML_SCHEMA, RL_SCHEMA, SE_SCHEMA,
                       grid_to_season, make_grid, rows_equal, rows_to_df,
                       season_to_grid)
from app.paste import apply_block

_EMPTY_RAIL = ("*No book yet — paste the masters above, or Book page → "
               "Load every line.*")


def build(session):
    bag = WatcherBag()
    seeding = {"on": False}

    def scn():
        return session.page.config

    # ---- one table section -------------------------------------------------
    tables = []          # (label, schema, grid, get_rows) for reseeds
    paste_parts = {}     # label -> (textarea, apply button, flash pane)

    def section(label, schema, get_rows, set_rows, note):
        grid = make_grid(schema, get_rows() if scn() is not None else [])
        flash = pn.pane.Markdown("", sizing_mode="stretch_width")

        def _write_back(event):
            s = scn()
            if s is None or seeding["on"]:
                return
            from app.grids import df_to_rows
            try:
                rows = df_to_rows(schema, event.new)
            except ValueError as e:
                flash.object = f"⚠ **{md_safe(e)}** — fix the cell; nothing was saved."
                return
            flash.object = ""
            if rows_equal(rows, get_rows()):
                return                      # reseed echo / no-op edit
            set_rows(s, rows)
            session.ctx.combos = s.combo_keys()
            session.bus.bump()

        grid.param.watch(_write_back, "value")

        ta = pn.widgets.TextAreaInput(
            placeholder="Click here, then Ctrl+V the block copied from the "
                        "workbook (same column order; a copied header row "
                        "is skipped automatically) …",
            # Panel's default max_length is 5000 and the browser clamps a
            # paste SILENTLY at it (W2a — a fleet-sized block is ~60KB)
            max_length=2_000_000,
            height=84, sizing_mode="stretch_width")
        apply_btn = pn.widgets.Button(
            label=f"Apply paste — replaces every {label} row",
            button_type="warning")

        def _apply(_event):
            s = scn()
            # value syncs on BLUR; value_input syncs per keystroke — read
            # the live one so paste-then-click cannot race to an empty read
            text = (ta.value_input or ta.value or "").strip()
            if s is None:
                flash.object = ("⚠ **No line active** — paste the MASTERS "
                                "above (they work on an empty book), or "
                                "Book page → Load every line.")
                return
            if not text:
                flash.object = ("⚠ **Nothing to paste** — click into the "
                                "box and Ctrl+V the copied block, then "
                                "Apply.")
                return
            notes: list = []
            try:
                rows = apply_block(schema, text, notes=notes)
            except ValueError as e:
                flash.object = f"⚠ **{md_safe(e)}** — nothing was applied."
                return
            if not rows:
                flash.object = "⚠ **The pasted text held no rows.**"
                return
            grid.value = rows_to_df(schema, rows)   # -> same write-back path
            ta.value = ""
            ta.value_input = ""
            extra = f" ({'; '.join(notes)})" if notes else ""
            flash.object = (f"✓ Applied **{len(rows)} row(s)** to {label}"
                            f"{extra} — blank/incomplete rows drop out per "
                            f"the populated-row rule.")

        apply_btn.on_click(_apply)
        paste_card = pn.Card(
            ta, apply_btn,
            pn.pane.Markdown(
                "*Ctrl+V pastes into the box above — pasting directly into "
                "the grid is not supported (the grid does support Ctrl+C "
                "copy-out).*"),
            title="📋 Paste from Excel — THIS line only",
            collapsed=False, sizing_mode="stretch_width")
        tables.append((label, schema, grid, get_rows))
        paste_parts[label] = (ta, apply_btn, flash)
        body = pn.Column(pn.pane.Markdown(note), paste_card, flash, grid,
                         sizing_mode="stretch_width")
        return body

    def _lr_get():
        return [] if scn() is None else scn().lr_rows

    def _rl_get():
        return [] if scn() is None else scn().rate_rows

    def _ml_get():
        return [] if scn() is None else scn().mod_rows

    def _se_get():
        return [] if scn() is None else season_to_grid(scn().season_rows)

    strip = pn.Tabs(
        ("tbl_LR", section(
            "tbl_LR", LR_SCHEMA, _lr_get,
            lambda s, rows: setattr(s, "lr_rows", rows),
            "One row per BU × state. Percent columns hold FRACTIONS "
            "(0.103 = 10.3%), exactly what Excel's cells hold.")),
        ("Rate Log", section(
            "Rate Log", RL_SCHEMA, _rl_get,
            lambda s, rows: setattr(s, "rate_rows", rows),
            "One row per rate change. Planned rows earn Filed % × "
            "Achievement.")),
        ("Mod Log", section(
            "Mod Log", ML_SCHEMA, _ml_get,
            lambda s, rows: setattr(s, "mod_rows", rows),
            "Stepped schedule-mod actions (D70). With any row present for a "
            "combo, its drift path hands off to this log on 1/1 of the plan "
            "year.")),
        ("Seasonality", section(
            "Seasonality", SE_SCHEMA, _se_get,
            lambda s, rows: setattr(s, "season_rows", grid_to_season(rows)),
            "Relative written weights by month, per state. Blank row = "
            "uniform writings.")),
        dynamic=False,                       # dynamic=True makes deaf cards
        sizing_mode="stretch_width")

    # ---- the MASTERS: every line at once (moved from Book, W3g) ------------
    master_parts = {}

    def _master_section(label):
        ta = pn.widgets.TextAreaInput(
            placeholder=(f"Ctrl+V the {label} MASTER here — the per-line "
                         "columns with a leading Line column. A copied "
                         "header row is skipped automatically."),
            # Panel's default max_length is 5000; the tbl_LR master alone
            # is ~60KB and the browser truncates silently (W2a)
            max_length=2_000_000,
            height=84, sizing_mode="stretch_width")
        apply_btn = pn.widgets.Button(
            label=f"Apply master — replaces {label} for every pasted line",
            button_type="warning")
        flash = pn.pane.Markdown("", sizing_mode="stretch_width")

        def _apply(_event):
            # value syncs on BLUR; value_input per keystroke (the P2 scar)
            text = (ta.value_input or ta.value or "").strip()
            if not text:
                flash.object = ("⚠ **Nothing to paste** — click into the "
                                "box, Ctrl+V the master block, then Apply.")
                return
            notes: list = []
            try:
                by, created = masters.apply_master(session, label, text,
                                                   notes=notes)
            except ValueError as e:
                flash.object = f"⚠ **{md_safe(e)}** — nothing was applied."
                return
            if not by:
                flash.object = "⚠ **The pasted text held no populated rows.**"
                return
            ta.value = ""
            ta.value_input = ""
            parts = ", ".join(f"{md_safe(ln)} ({len(rows)})"
                              for ln, rows in by.items())
            newly = ("" if not created else
                     " — new line(s): " + ", ".join(md_safe(c)
                                                    for c in created))
            extra = f" ({'; '.join(notes)})" if notes else ""
            flash.object = (f"✓ Applied **{label}** to {len(by)} line(s): "
                            f"{parts}{newly}{extra}. Lines not in the paste "
                            f"kept theirs.")

        apply_btn.on_click(_apply)
        master_parts[label] = (ta, apply_btn, flash)
        return pn.Column(ta, apply_btn, flash, sizing_mode="stretch_width")

    master_card = pn.Card(
        pn.pane.Markdown(
            "*One table per master, spanning every line — column order is "
            "the per-line paste block with a leading **Line** column "
            "(a BU can write several lines, so the line is never "
            "inferred). Lines present in a paste get that table replaced; "
            "absent lines keep theirs.*"),
        pn.Tabs(*[(lbl, _master_section(lbl)) for lbl in masters.MASTERS],
                dynamic=False, sizing_mode="stretch_width"),
        title="MASTERS — every line at once (leading Line column)",
        # open on a cold start; folded when a book already exists
        collapsed=bool(session.page.book), sizing_mode="stretch_width")

    # the visual firewall between the two paste surfaces
    context_bar = html_pane()

    def _context(*_events):
        s = scn()
        if s is None:
            context_bar.object = echo_html(
                "No line active — the masters above work on an empty book.")
        else:
            context_bar.object = echo_html(
                f"Editing: {s.lob} — the four tables below touch THIS "
                f"line only. Switch lines on the Book page.")

    # ---- scenario-level toggles (sidebar) ----------------------------------
    term = pn.widgets.IntInput(label="Term (months)", start=1, end=12,
                               value=12, sizing_mode="stretch_width")
    trend_d = pn.widgets.FloatInput(label="Net trend default", step=0.001,
                                    value=0.0, sizing_mode="stretch_width")
    season_on = pn.widgets.Checkbox(label="Seasonality ON")
    mod_master = pn.widgets.Checkbox(label="Mod adjustment master ON",
                                     value=True)

    def _toggle(attr, conv=lambda v: v):
        def _cb(event):
            s = scn()
            if s is None or seeding["on"]:
                return
            if getattr(s, attr) == conv(event.new):
                return
            setattr(s, attr, conv(event.new))
            session.bus.bump()
        return _cb

    term.param.watch(_toggle("term_months", int), "value")
    trend_d.param.watch(_toggle("trend_default", float), "value")
    season_on.param.watch(_toggle("season_on", bool), "value")
    mod_master.param.watch(_toggle("mod_master", bool), "value")

    # ---- validation rail (the Checks rules, live) --------------------------
    rail = pn.pane.Markdown(_EMPTY_RAIL, sizing_mode="stretch_width")

    def _revalidate(*_events):
        s = scn()
        if s is None:
            rail.object = _EMPTY_RAIL
            return
        issues = validate_mod.validate(s)
        fails = [m for sev, m in issues if sev == "FAIL"]
        warns = [m for sev, m in issues if sev == "WARN"]
        head = (f"**Checks: {len(fails)} FAIL · {len(warns)} WARN**"
                if issues else "**Checks: clean** ✓")
        lines = [f"- 🟥 {md_safe(m)}" for m in fails[:8]]
        lines += [f"- 🟡 {md_safe(m)}" for m in warns[:8]]
        if len(fails) > 8 or len(warns) > 8:
            lines.append(f"- … and {max(len(fails) - 8, 0) + max(len(warns) - 8, 0)} more")
        rail.object = head + "\n" + "\n".join(lines)

    # ---- reseed on scenario replacement ------------------------------------
    def _reseed(*_events):
        s = scn()
        seeding["on"] = True
        try:
            for _label, schema, grid, get_rows in tables:
                grid.value = rows_to_df(schema, get_rows() if s else [])
            if s is not None:
                term.value = int(s.term_months)
                trend_d.value = float(s.trend_default)
                season_on.value = bool(s.season_on)
                mod_master.value = bool(s.mod_master)
        finally:
            seeding["on"] = False
        _context()
        _revalidate()

    bag.watch(session.ctx, _reseed, "data_rev")
    rail_rev = debounce(session.bus.param.rev, delay_ms=300, bag=bag)
    bag.watch(rail_rev, _revalidate, "rev")
    _reseed()

    sidebar = pn.Column(
        pn.pane.Markdown("**Scenario settings**"),
        term, trend_d, season_on, mod_master,
        pn.layout.Divider(),
        rail,
        sizing_mode="stretch_width")
    page_notes = note_list([
        "Blank and 0 are DIFFERENT values: a blank Net rate P means no "
        "net selection; 0% means net mode ON at a zero target. Blanks "
        "render blank (W2a).",
        "Percent columns hold FRACTIONS (0.103 = 10.3%) — pasting "
        "'10.3%' or '0.103' both land as 0.103.",
        "Paste replaces the whole table; blank/incomplete rows drop out "
        "per the populated-row rule (BU + State, plus a date for the "
        "logs).",
    ])
    header = exhibit_header(
        "Inputs — start the book here",
        "The masters seed EVERY line at once; the four tables below edit "
        "the ACTIVE line. Percent columns hold fractions, exactly what "
        "Excel's cells hold.")
    return {"main": pn.Column(header, master_card, context_bar, strip,
                              page_notes, sizing_mode="stretch_both"),
            "sidebar": sidebar, "bag": bag,
            "tables": {label: grid for label, _s, grid, _g in tables},
            "paste": paste_parts, "rail": rail,
            "masters": master_parts, "master_card": master_card,
            "context": context_bar,
            # re-seed after the visibility flip: Tabulator drops its virtual
            # rows while hidden and needs a value write to redraw (P4)
            "on_show": _reseed}
