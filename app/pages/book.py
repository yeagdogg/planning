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

from app import compute, importers
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.engineio import run_async
from app.glue.exhibit import html_pane as ex_html_pane, note_list
from app.glue.format import (DASH, fmt_count, fmt_dollar, fmt_pct, md_safe,
                             tab_formatters)
from app.glue.theme import FAIL_RED, NAVY, TABLE_CSS

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
    return "<div class='plw-cards'>" + "".join(cards) + "</div>"


_EMPTY_HTML = ("<div class='plw-cards'><div class='c'>"
               "<div class='v'>—</div><div class='l'>No lines loaded — paste the "
               "masters on the <b>Inputs</b> page, or use "
               "<b>Load every line</b> here.</div></div></div>")


def _display(frame: pd.DataFrame) -> pd.DataFrame:
    df = frame.drop(columns=["key"]).copy()
    df["net"] = df["net"].map(lambda v: "net" if v else "")
    return df


def _book_styles(df: pd.DataFrame) -> pd.DataFrame:
    """The W2d chrome: the summary module's pure style helpers on the
    all-combo table — plan-LR color scales + the EP bar (D104's look)."""
    from app.summary import ep_bar, three_color_scale
    css = pd.DataFrame("", index=df.index, columns=df.columns)
    for col_ in ("lr_p", "lr_p1"):
        for i, s in zip(df.index, three_color_scale(df[col_])):
            if s:
                css.loc[i, col_] = s
    for i, s in zip(df.index, ep_bar(df["ep"])):
        if s:
            css.loc[i, "ep"] = s
    return css


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
    cards = ex_html_pane()
    cards.object = _EMPTY_HTML

    line_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["on", "lob", "combos", "ep", "lr_p", "lr_p1",
                              "a_rate", "a_mod", "errors"]),
        disabled=True, show_index=False, titles=_TITLES,
        formatters=tab_formatters(ep="dollar", lr_p="pct", lr_p1="pct",
                                  a_rate="idx", a_mod="idx", combos="count",
                                  errors="count"),
        text_align={c: "right" for c in ("combos", "ep", "lr_p", "lr_p1",
                                         "a_rate", "a_mod", "errors")},
        layout="fit_data_table", max_height=300,      # grow to the roster
        sizing_mode="stretch_width",
        selectable=False,                # row click activates; never selects
        stylesheets=[TABLE_CSS],
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
        layout="fit_data_table", max_height=640,     # grow to the page
        sizing_mode="stretch_width",
        selectable=False,      # click-through rides cellClick — survives
        # 378 real rows tripped Panel's silent auto-pagination at 20/page,
        # and it never reset when filtered down — make it a decision
        pagination="local", page_size=100,
        stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],       # rendered despite show_index=False
        configuration={"clipboard": "copy",
                       "columnDefaults": {"headerWordWrap": True}})

    mix_md = pn.pane.Markdown(
        "*Every mean is adjusted-plan-EP weighted over what's in view — the "
        "Book workbook's own rule. Weighted factors do **not** multiply to "
        "the weighted plan LR; the difference is mix, disclosed rather than "
        "normalised away (D103). Click a combo row to open it on the Combo "
        "page.*", sizing_mode="stretch_width")

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
        disp = _display(view)
        table.value = disp
        table.style = disp.style.apply(_book_styles, axis=None)

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
    book_notes = note_list([
        "Color scale spans the combos in view (green low → red high plan "
        "LR); the EP bar is proportional to the largest combo in view.",
        "Rows the engine rejected carry no weight anywhere and show their "
        "reason in the Problem column.",
    ])
    main = pn.Column(cards,
                     pn.pane.Markdown("### By line"), line_tbl,
                     pn.pane.Markdown("### Every combo"), mix_md, table,
                     book_notes,
                     sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "table": table, "line_tbl": line_tbl, "cards": cards,
            "filters": {"line": line_f, "bu": bu_f, "state": state_f},
            "open_row": _open_row, "load_btn": load_btn, "status": status,
            # Tabulator's virtual renderer tears its rows down while the
            # page is hidden and never redraws on the visibility flip —
            # main._switch calls this after re-showing (found live in P4)
            "on_show": _render}
