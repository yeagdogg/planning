"""Scenarios — the hub's memory and its output chute (P5).

Save the whole book to one yaml in scenarios/ (each save appends a
changelog entry diffed against the file it replaces — the audit trail
travels WITH the scenario; git carries the version control). Load
replaces the in-memory book behind a two-step confirm (pn.Modal is
unusable — the workbench scar — so the button itself arms and confirms).
Diff shows file → current without touching either.

Generate builds real fleet workbooks from the book through the SAME
generator the release harness verifies, tagged _APP beside the fleet
(never over it), with optional Excel recalc via the system python's COM
session.
"""
from __future__ import annotations

from pathlib import Path

import panel as pn

from app import exporters, importers, scenarios_io
from app.glue.bindings import WatcherBag, debounce, gate_hidden
from app.glue.engineio import run_async
from app.glue.format import fmt_dollar, md_safe

_DIFF_CAP = 40


def build(session):
    bag = WatcherBag()
    cfg = importers.app_config()

    # ---- sidebar: save -----------------------------------------------------
    name_in = pn.widgets.TextInput(label="Scenario name",
                                   value=f"plan-{cfg.plan_year}",
                                   sizing_mode="stretch_width")
    note_in = pn.widgets.TextInput(label="Change note (goes in the log)",
                                   placeholder="what changed, in a phrase",
                                   sizing_mode="stretch_width")
    save_btn = pn.widgets.Button(label="Save book to scenarios/",
                                 button_type="primary",
                                 sizing_mode="stretch_width")
    save_flash = pn.pane.Markdown("", sizing_mode="stretch_width")

    # ---- sidebar: load / inspect -------------------------------------------
    file_sel = pn.widgets.Select(label="Scenario file",
                                 options=scenarios_io.scenario_files(),
                                 sizing_mode="stretch_width")
    load_btn = pn.widgets.Button(label="Load — replaces the current book",
                                 sizing_mode="stretch_width")
    diff_btn = pn.widgets.Button(label="Diff file → current book",
                                 sizing_mode="stretch_width")
    load_flash = pn.pane.Markdown("", sizing_mode="stretch_width")

    # ---- main: summary, generate, log/diff ---------------------------------
    summary = pn.pane.Markdown("", sizing_mode="stretch_width")
    lines_ms = pn.widgets.MultiChoice(label="Lines to generate (empty = all)",
                                      options=[],
                                      sizing_mode="stretch_width")
    dialect = pn.widgets.Select(
        label="Formula dialect",
        options={"config default": None, "classic": "classic",
                 "modern": "modern"},
        sizing_mode="stretch_width")
    recalc_cb = pn.widgets.Checkbox(
        name="Recalc in Excel after building (needs Excel; ~1 min for six)")
    gen_btn = pn.widgets.Button(label="Generate workbooks (_APP tag)",
                                button_type="success")
    book_btn = pn.widgets.Button(
        label="Generate fleet + summary Book → output/app/ (Excel, ~2 min)",
        button_type="success")
    gen_flash = pn.pane.Markdown("", sizing_mode="stretch_width")
    log_md = pn.pane.Markdown("", sizing_mode="stretch_width")

    def _refresh_files(keep=None):
        file_sel.options = scenarios_io.scenario_files()
        if keep and keep in file_sel.options.values():
            file_sel.value = keep

    def _summary(*_events):
        book = session.page.book
        if not book:
            summary.object = ("*No lines loaded — Book page → **Load every "
                              "line**, paste masters there, or load a "
                              "scenario file here.*")
            lines_ms.options = []
            return
        total = sum(len(s.combo_keys()) for s in book.values())
        ep = sum(s.ep(k) for s in book.values() for k in s.combo_keys())
        parts = ", ".join(f"{md_safe(l)} ({len(s.combo_keys())})"
                          for l, s in book.items())
        summary.object = (f"**Current book: {len(book)} line(s) · {total} "
                          f"combos · plan EP {fmt_dollar(ep)}** — {parts}")
        lines_ms.options = list(book)
        lines_ms.value = [v for v in lines_ms.value if v in book]

    # ---- save --------------------------------------------------------------
    def _save(_event):
        book = session.page.book
        if not book:
            save_flash.object = "⚠ **Nothing to save** — the book is empty."
            return
        name = (name_in.value or "").strip() or f"plan-{cfg.plan_year}"
        path = scenarios_io.SCEN_DIR / scenarios_io.safe_name(name)
        try:
            changes = scenarios_io.save_book(path, book, name,
                                             note=note_in.value or "")
        except Exception as e:                          # noqa: BLE001
            save_flash.object = f"⚠ **Save failed** — {md_safe(e)}"
            return
        note_in.value = ""
        _refresh_files(keep=str(path))
        save_flash.object = (f"✓ Saved **{md_safe(path.name)}** — "
                             f"{len(changes)} change(s) logged.")
        _show_log()

    save_btn.on_click(_save)

    # ---- load (two-step confirm: the pn.Modal scar) ------------------------
    armed = {"on": False}

    def _disarm():
        armed["on"] = False
        load_btn.label = "Load — replaces the current book"
        load_btn.button_type = "default"

    def _load(_event):
        path = file_sel.value
        if not path:
            load_flash.object = "⚠ **No scenario file selected.**"
            return
        if not armed["on"]:
            armed["on"] = True
            load_btn.label = "Confirm — unsaved edits will be lost"
            load_btn.button_type = "danger"
            return
        _disarm()
        try:
            book, doc = scenarios_io.load_book(path)
        except Exception as e:                          # noqa: BLE001
            load_flash.object = f"⚠ **Load failed** — {md_safe(e)}"
            return
        if not book:
            load_flash.object = "⚠ **The file holds no lines.**"
            return
        session.page.book.clear()
        session.page.book.update(book)
        session.activate(next(iter(book)))
        name_in.value = str(doc.get("name") or name_in.value)
        load_flash.object = (f"✓ Loaded **{md_safe(doc.get('name') or path)}"
                             f"** — {len(book)} line(s): "
                             + ", ".join(md_safe(l) for l in book))

    load_btn.on_click(_load)
    file_sel.param.watch(lambda _e: (_disarm(), _show_log()), "value")

    # ---- diff --------------------------------------------------------------
    def _diff(_event):
        path = file_sel.value
        if not path:
            load_flash.object = "⚠ **No scenario file selected.**"
            return
        try:
            doc = scenarios_io.read_doc(path)
            cur = scenarios_io.book_to_doc(session.page.book, "current")
            changes = scenarios_io.diff_books(doc.get("lines") or {},
                                              cur["lines"])
        except Exception as e:                          # noqa: BLE001
            load_flash.object = f"⚠ **Diff failed** — {md_safe(e)}"
            return
        shown = [f"- {md_safe(c)}" for c in changes[:_DIFF_CAP]]
        if len(changes) > _DIFF_CAP:
            shown.append(f"- … and {len(changes) - _DIFF_CAP} more")
        log_md.object = ("### Diff — file → current book\n"
                         + ("\n".join(shown) if shown
                            else "*No differences.*"))

    diff_btn.on_click(_diff)

    # ---- changelog viewer --------------------------------------------------
    def _show_log(*_events):
        path = file_sel.value
        if not path:
            log_md.object = ""
            return
        try:
            doc = scenarios_io.read_doc(path)
        except Exception as e:                          # noqa: BLE001
            log_md.object = f"⚠ could not read file: {md_safe(e)}"
            return
        entries = list(doc.get("changelog") or [])[-8:][::-1]
        if not entries:
            log_md.object = "*The file carries no changelog yet.*"
            return
        parts = []
        for e in entries:
            head = f"**{md_safe(e.get('at', '?'))}**"
            if e.get("note"):
                head += f" — {md_safe(e['note'])}"
            body = "".join(f"\n- {md_safe(c)}"
                           for c in (e.get("changes") or [])[:12])
            more = len(e.get("changes") or []) - 12
            if more > 0:
                body += f"\n- … and {more} more"
            parts.append(head + body)
        log_md.object = ("### Changelog — " + md_safe(Path(path).name)
                         + "\n\n" + "\n\n".join(parts))

    # ---- generate ----------------------------------------------------------
    def _generate(_event):
        book = session.page.book
        if not book:
            gen_flash.object = "⚠ **Nothing to generate** — the book is empty."
            return
        sel = list(lines_ms.value) or list(book)
        mode = dialect.value
        do_recalc = bool(recalc_cb.value)
        gen_btn.loading = True

        def _work():
            paths = [exporters.export_line(book[lob], formula_mode=mode)
                     for lob in sel]
            if do_recalc:
                exporters.recalc_files(paths)
                tail = "values recalculated in Excel."
            else:
                tail = "values populate on first open in Excel."
            return paths, tail

        def _done(res):
            gen_btn.loading = False
            paths, tail = res
            files = "\n".join(f"- `{md_safe(p.name)}`" for p in paths)
            gen_flash.object = (f"✓ **Built {len(paths)} workbook(s)** — "
                                f"{tail}\n{files}")

        def _err(e):
            gen_btn.loading = False
            gen_flash.object = f"⚠ **Generate failed** — {md_safe(e)}"

        run_async(_work, _done, _err)

    gen_btn.on_click(_generate)

    def _generate_book(_event):
        book = dict(session.page.book)
        if not book:
            gen_flash.object = "⚠ **Nothing to generate** — the book is empty."
            return
        book_btn.loading = True

        def _work():
            return exporters.export_fleet_and_book(book, formula_mode=dialect.value)

        def _done(res):
            book_btn.loading = False
            paths, book_file = res
            gen_flash.object = (
                f"✓ **Fleet + Book built and recalculated** in "
                f"`output/app/` — {len(paths)} line file(s) + "
                f"`{md_safe(book_file.name)}`. The verified fleet in "
                f"output/ was not touched.")

        def _err(e):
            book_btn.loading = False
            gen_flash.object = f"⚠ **Fleet + Book failed** — {md_safe(e)}"

        run_async(_work, _done, _err)

    book_btn.on_click(_generate_book)

    # ---- wiring ------------------------------------------------------------
    # hidden pages skip the refresh; _on_show recomputes it (W3h)
    summary_g = gate_hidden(_summary)
    rail_rev = debounce(session.bus.param.rev, delay_ms=300, bag=bag)
    bag.watch(rail_rev, summary_g, "rev")
    bag.watch(session.ctx, summary_g, "data_rev")
    _summary()
    _show_log()

    def _on_show():
        _refresh_files(keep=file_sel.value)
        _summary()
        _disarm()

    sidebar = pn.Column(
        pn.pane.Markdown("**Save the book**"),
        name_in, note_in, save_btn, save_flash,
        pn.layout.Divider(),
        pn.pane.Markdown("**Load / inspect**"),
        file_sel, load_btn, diff_btn, load_flash,
        sizing_mode="stretch_width")
    main = pn.Column(
        summary,
        pn.pane.Markdown("### Generate workbooks\n*Real fleet workbooks "
                         "from the current book, through the verified "
                         "generator — tagged `_APP`, never overwriting the "
                         "fleet's own files.*"),
        lines_ms, dialect, recalc_cb, gen_btn, book_btn, gen_flash,
        pn.layout.Divider(),
        log_md,
        sizing_mode="stretch_both")
    summary_g.attach(main)
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "on_show": _on_show,
            "widgets": {"name": name_in, "note": note_in, "save": save_btn,
                        "file": file_sel, "load": load_btn, "diff": diff_btn,
                        "lines": lines_ms, "dialect": dialect,
                        "recalc": recalc_cb, "generate": gen_btn,
                        "book": book_btn},
            "flashes": {"save": save_flash, "load": load_flash,
                        "gen": gen_flash, "log": log_md},
            "summary": summary}
