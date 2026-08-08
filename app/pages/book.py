"""The Book — every line at once (P4): the hub's landing view.

One row per line × BU × state across every LOADED line, aggregated by the
Book workbook's own rule (sheets_book ``wtd()``): every mean is
adjusted-plan-EP weighted over what is in view, '—' when nothing carries
weight. Weighted factors deliberately do NOT compound to the weighted LR —
the gap is mix, and the caption says so rather than normalising it away
(D103).

Click a combo row to open it on the Combo page: the line activates (Inputs
and Combo follow it) and the selector adopts the combo. Edits on Inputs
move these numbers through the same debounced bus as every other surface;
``compute.book_results`` recomputes only the ACTIVE line per edit, so the
book stays live at keystroke cadence with six lines loaded.
"""
from __future__ import annotations

import pandas as pd
import panel as pn

from app import compute, importers, masters
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.engineio import run_async
from app.glue.format import (DASH, fmt_count, fmt_dollar, fmt_pct, md_safe,
                             tab_formatters)
from app.glue.theme import FAIL_RED, NAVY

_CARD_CSS = f"""
<style>
.plw-cards {{ display: flex; flex-wrap: wrap; gap: 10px;
  font-family: var(--body-font, Segoe UI), sans-serif; }}
.plw-cards .c {{ border: 1px solid #e3e6ea; border-radius: 8px;
  padding: 8px 16px; min-width: 128px; }}
.plw-cards .v {{ font-size: 1.35em; font-weight: 700; color: {NAVY};
  font-variant-numeric: tabular-nums; }}
.plw-cards .l {{ color: #777; font-size: 0.8em; }}
.plw-cards .bad {{ color: {FAIL_RED}; }}
</style>
"""

_FRAME_COLS = ["lob", "bu", "state", "ep", "lr_p", "lr_p1", "a_rate",
               "a_mod", "crl", "ecy", "net", "error"]
_TITLES = {"lob": "Line", "bu": "BU", "state": "State",
           "ep": "Plan EP (000s)", "lr_p": "Plan LR", "lr_p1": "Plan LR +1",
           "a_rate": "A_rate", "a_mod": "A_mod", "crl": "CRL",
           "ecy": "Earned level", "net": "Net?", "error": "Problem",
           "on": "", "combos": "Combos", "errors": "Problems"}


def _cards_html(w, plan_year) -> str:
    def card(label, value, bad=False):
        cls = "v bad" if bad else "v"
        return (f"<div class='c'><div class='{cls}'>{value}</div>"
                f"<div class='l'>{label}</div></div>")

    cards = [
        card("lines in view", fmt_count(w["lines"])),
        card("combos", fmt_count(w["combos"])),
        card("adjusted plan EP", fmt_dollar(w["ep"])),
        card(f"CY {plan_year} plan LR", fmt_pct(w["lr_p"])),
        card(f"CY {plan_year + 1} plan LR", fmt_pct(w["lr_p1"])),
        card("net-target combos", fmt_count(w["net"])),
    ]
    if w["errors"]:
        cards.append(card("rows the engine rejected", fmt_count(w["errors"]),
                          bad=True))
    return _CARD_CSS + "<div class='plw-cards'>" + "".join(cards) + "</div>"


_EMPTY_HTML = (_CARD_CSS + "<div class='plw-cards'><div class='c'>"
               "<div class='v'>—</div><div class='l'>No lines loaded — use "
               "<b>Load every line</b>, or open one workbook on the Combo "
               "page.</div></div></div>")


def _display(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.drop(columns=["key"]).copy()
    df["net"] = df["net"].map(lambda v: "net" if v else "")
    return df


def build(session):
    bag = WatcherBag()

    # ---- sidebar: the one-click on-ramp + filters --------------------------
    load_btn = pn.widgets.Button(label="Load every line from output/",
                                 button_type="primary",
                                 sizing_mode="stretch_width")
    status = pn.pane.Markdown("", sizing_mode="stretch_width")

    line_f = pn.widgets.MultiChoice(label="Lines", options=[],
                                    placeholder="all lines",
                                    sizing_mode="stretch_width")
    bu_f = pn.widgets.MultiChoice(label="Business units", options=[],
                                  placeholder="all BUs",
                                  sizing_mode="stretch_width")
    state_f = pn.widgets.MultiChoice(label="States", options=[],
                                     placeholder="all states",
                                     sizing_mode="stretch_width")

    def _load_all(_event):
        choices = importers.fleet_choices()
        if not choices:
            status.object = "⚠ **No line workbooks found** in output/."
            return
        load_btn.loading = True

        def _work():
            return [importers.from_workbook(p) for p in choices.values()]

        def _done(loaded):
            load_btn.loading = False
            for scn in loaded:
                session.page.book[scn.lob] = scn
            if loaded:
                session.activate(loaded[0].lob)
            status.object = ("**" + str(len(loaded)) + " lines loaded** — "
                             + ", ".join(md_safe(s.lob) for s in loaded))

        def _err(e):
            load_btn.loading = False
            status.object = f"⚠ **Load failed** — {md_safe(e)}"

        run_async(_work, _done, _err)

    load_btn.on_click(_load_all)

    def _bu_opts():
        vals = set()
        for scn in session.page.book.values():
            for k in scn.combo_keys():
                vals.add(k.split("|", 1)[0])
        return sorted(vals)

    def _state_opts():
        vals = set()
        for scn in session.page.book.values():
            for k in scn.combo_keys():
                vals.add(k.split("|", 1)[1])
        return sorted(vals)

    sync_options(line_f, lambda: list(session.page.book),
                 session.ctx.param.data_rev, bag=bag)
    sync_options(bu_f, _bu_opts, session.ctx.param.data_rev, bag=bag)
    sync_options(state_f, _state_opts, session.ctx.param.data_rev, bag=bag)

    # ---- main: cards, per-line roll-up, all combos -------------------------
    cards = pn.pane.HTML(_EMPTY_HTML, sizing_mode="stretch_width")

    line_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["on", "lob", "combos", "ep", "lr_p", "lr_p1",
                              "a_rate", "a_mod", "errors"]),
        disabled=True, show_index=False, titles=_TITLES,
        formatters=tab_formatters(ep="dollar", lr_p="pct", lr_p1="pct",
                                  a_rate="idx", a_mod="idx", combos="count",
                                  errors="count"),
        text_align={c: "right" for c in ("combos", "ep", "lr_p", "lr_p1",
                                         "a_rate", "a_mod", "errors")},
        layout="fit_data_table", height=230, sizing_mode="stretch_width",
        hidden_columns=["_index"],       # rendered despite show_index=False
        configuration={"clipboard": "copy"})

    table = pn.widgets.Tabulator(
        pd.DataFrame(columns=_FRAME_COLS),
        disabled=True, show_index=False, titles=_TITLES,
        formatters=tab_formatters(ep="dollar", lr_p="pct", lr_p1="pct",
                                  a_rate="idx", a_mod="idx", crl="idx",
                                  ecy="idx"),
        text_align={c: "right" for c in ("ep", "lr_p", "lr_p1", "a_rate",
                                         "a_mod", "crl", "ecy")},
        frozen_columns=["lob", "bu", "state"],       # by NAME (the scar)
        layout="fit_data_table", height=520, sizing_mode="stretch_width",
        hidden_columns=["_index"],       # rendered despite show_index=False
        configuration={"clipboard": "copy"})

    mix_md = pn.pane.Markdown(
        "*Every mean is adjusted-plan-EP weighted over what's in view — the "
        "Book workbook's own rule. Weighted factors do **not** multiply to "
        "the weighted plan LR; the difference is mix, disclosed rather than "
        "normalised away (D103). Click a combo row to open it on the Combo "
        "page.*", sizing_mode="stretch_width")

    # ---- master paste (P5): the paste-first on-ramp ------------------------
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
        title="📋 Master paste — start or update the whole book from "
              "three tables",
        collapsed=True, sizing_mode="stretch_width")

    def _view(frame: pd.DataFrame) -> pd.DataFrame:
        view = frame
        if line_f.value:
            view = view[view["lob"].isin(line_f.value)]
        if bu_f.value:
            view = view[view["bu"].isin(bu_f.value)]
        if state_f.value:
            view = view[view["state"].isin(state_f.value)]
        return view

    def _render(*_events):
        if not session.page.book:
            cards.object = _EMPTY_HTML
            line_tbl.value = line_tbl.value.iloc[0:0]
            table.value = table.value.iloc[0:0]
            return
        view = _view(compute.book_frame(session))
        any_scn = next(iter(session.page.book.values()))
        cards.object = _cards_html(compute.book_weighted(view),
                                   any_scn.plan_year)
        roll = compute.line_rollup(view)
        active = getattr(session.page.config, "lob", None)
        roll.insert(0, "on",
                    roll["lob"].map(lambda l: "▶" if l == active else ""))
        line_tbl.value = roll
        table.value = _display(view)

    def _open_row(event):
        """Combo click-through: activate the line, focus the combo, go."""
        try:
            rec = table.value.iloc[int(event.row)]
        except Exception:                               # noqa: BLE001
            return
        session.ctx.focus = f"{rec['bu']}|{rec['state']}"
        session.activate(rec["lob"])
        if callable(session.goto):
            session.goto("combo")

    def _line_click(event):
        """Line click just refocuses the active line (stay on the book)."""
        try:
            lob = line_tbl.value.iloc[int(event.row)]["lob"]
        except Exception:                               # noqa: BLE001
            return
        session.activate(lob)

    table.on_click(_open_row)
    line_tbl.on_click(_line_click)

    rail_rev = debounce(session.bus.param.rev, delay_ms=300, bag=bag)
    bag.watch(rail_rev, _render, "rev")
    bag.watch(session.ctx, _render, "data_rev")
    for f in (line_f, bu_f, state_f):
        bag.watch(f, _render, "value")
    _render()

    sidebar = pn.Column(
        pn.pane.Markdown("**The whole book**"),
        load_btn, status,
        pn.layout.Divider(),
        pn.pane.Markdown("**Filters** — empty means everything"),
        line_f, bu_f, state_f,
        sizing_mode="stretch_width")
    main = pn.Column(cards, master_card,
                     pn.pane.Markdown("### By line"), line_tbl,
                     pn.pane.Markdown("### Every combo"), mix_md, table,
                     sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "table": table, "line_tbl": line_tbl, "cards": cards,
            "filters": {"line": line_f, "bu": bu_f, "state": state_f},
            "open_row": _open_row, "load_btn": load_btn, "status": status,
            "masters": master_parts,
            # Tabulator's virtual renderer tears its rows down while the
            # page is hidden and never redraws on the visibility flip —
            # main._switch calls this after re-showing (found live in P4)
            "on_show": _render}
