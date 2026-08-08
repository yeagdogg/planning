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

from app import (compute, exporters, importers, scenarios_io, targetbook,
                 targets_io)
from app.glue.bindings import WatcherBag, debounce, gate_hidden
from app.glue.engineio import run_async
from app.glue.format import fmt_dollar, md_safe
from app.undo import UndoStack

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

    # ---- field targets (W4a): the small file that goes OUT to the field ----
    targets_btn = pn.widgets.Button(
        label="Generate field targets files → output/targets/",
        button_type="success")
    targets_flash = pn.pane.Markdown("", sizing_mode="stretch_width")

    # ---- collecting them back (W4b) ----------------------------------------
    folder_in = pn.widgets.TextInput(
        label="Also scan this folder (optional)",
        placeholder=r"e.g. \\share\plan\returned targets",
        sizing_mode="stretch_width")
    scan_btn = pn.widgets.Button(label="Scan for returned files")
    coll_sel = pn.widgets.Select(label="Returned file", options={},
                                 sizing_mode="stretch_width")
    rows_ms = pn.widgets.MultiChoice(label="Rows to apply", options=[],
                                     sizing_mode="stretch_width")
    apply_btn = pn.widgets.Button(label="Apply selected targets",
                                  button_type="success", disabled=True)
    coll_undo = pn.widgets.Button(label="↶ Undo", disabled=True)
    coll_flash = pn.pane.Markdown("", sizing_mode="stretch_width")
    coll_preview = pn.pane.Markdown("", sizing_mode="stretch_width")
    coll_undo_stack = UndoStack()
    coll_held: dict = {"tf": None, "changes": []}

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

    def _generate_targets(_event):
        """One small workbook per line: net targets in, plan LRs live out.
        The constants are engine output baked at build time, so the file
        needs no add-in and no trip back here to compute."""
        book = session.page.book
        if not book:
            targets_flash.object = ("⚠ **Nothing to generate** — the book "
                                    "is empty.")
            return
        sel = list(lines_ms.value) or list(book)
        do_recalc = bool(recalc_cb.value)
        prog = compute.program_results(session)
        targets_btn.loading = True

        def _work():
            paths = targetbook.build_all(
                {lob: book[lob] for lob in sel}, prog_by_lob=prog)
            if do_recalc:
                exporters.recalc_files(paths)
                tail = "values recalculated in Excel."
            else:
                tail = "values populate on first open in Excel."
            return paths, tail

        def _done(res):
            targets_btn.loading = False
            paths, tail = res
            files = "\n".join(f"- `{md_safe(p.name)}`" for p in paths)
            targets_flash.object = (
                f"✓ **Built {len(paths)} targets file(s)** — {tail} Only the "
                f"net-target cells are editable; hand them back and use "
                f"Collect below.\n{files}")

        def _err(e):
            targets_btn.loading = False
            targets_flash.object = f"⚠ **Targets build failed** — {md_safe(e)}"

        run_async(_work, _done, _err)

    targets_btn.on_click(_generate_targets)

    # ---- collect the returned files ----------------------------------------
    def _fmt_target(v):
        return "—" if v is None else f"{v:+.1%}"

    def _scan(*_events):
        found = targets_io.scan_targets(folder_in.value or None)
        opts, labels = {}, []
        for p in found:
            try:
                tf = targets_io.read_targets(p)
            except targets_io.TargetsReadError:
                continue
            scn = session.page.book.get(tf.lob)
            tail = ""
            if scn is None:
                tail = " · line NOT loaded"
            elif targets_io.is_stale(tf, scn):
                tail = " · STALE inputs"
            labels.append(f"{p.name} — {tf.lob} {tf.plan_year} · "
                          f"generated {tf.generated}{tail}")
            opts[labels[-1]] = str(p)
        coll_sel.options = opts
        coll_flash.object = (
            f"Found **{len(opts)}** returned file(s)." if opts else
            "*No returned targets files found — generate some above, or "
            "point at the folder they came back to.*")
        _preview()

    scan_btn.on_click(_scan)

    def _preview(*_events):
        path = coll_sel.value
        coll_held["tf"], coll_held["changes"] = None, []
        rows_ms.options, rows_ms.value = [], []
        apply_btn.disabled = True
        if not path:
            coll_preview.object = ""
            return
        try:
            tf = targets_io.read_targets(path)
        except targets_io.TargetsReadError as e:
            coll_preview.object = f"⚠ **{md_safe(e)}**"
            return
        scn = session.page.book.get(tf.lob)
        if scn is None:
            coll_preview.object = (
                f"⚠ **{md_safe(tf.lob)} is not loaded** — load that line "
                f"before applying its targets.")
            return
        if tf.plan_year != scn.plan_year:
            coll_preview.object = (
                f"⚠ **Plan year mismatch** — the file is {tf.plan_year}, the "
                f"book's {md_safe(tf.lob)} is {scn.plan_year}. Nothing can "
                f"be applied.")
            return
        changes, skipped = targets_io.diff_targets(tf, scn)
        coll_held["tf"], coll_held["changes"] = tf, changes
        rows_ms.options = [c.key for c in changes]
        rows_ms.value = [c.key for c in changes]
        apply_btn.disabled = not changes

        head = []
        if targets_io.is_stale(tf, scn):
            head.append(
                "⚠ **Stale inputs** — this file was built before the current "
                "dates/mods/loss ratios, so the loss ratios shown IN it were "
                "computed from older inputs. The targets are still what the "
                "field asked for; applying them recomputes from truth.")
        if tf.problems:
            head.append("**Cells skipped:**\n"
                        + "\n".join(f"- {md_safe(m)}" for m in tf.problems))
        if skipped:
            head.append("\n".join(f"- {md_safe(m)}" for m in skipped))
        if not changes:
            head.append("*No target changes — this file matches the book.*")
        else:
            lines = targets_io.diff_sentences(tf, scn, changes)[:_DIFF_CAP]
            head.append(f"**{len(changes)} row(s) changed:**\n"
                        + "\n".join(f"- {md_safe(s)}" for s in lines))
            warns = [c for c in changes if c.warn]
            if warns:
                head.append("\n".join(
                    f"- ⚠ {md_safe(c.key)}: {md_safe(c.warn)}"
                    for c in warns))
        coll_preview.object = "\n\n".join(head)

    coll_sel.param.watch(_preview, "value")

    def _apply_targets(_event):
        tf = coll_held["tf"]
        if tf is None:
            return
        scn = session.page.book.get(tf.lob)
        applied, msgs = targets_io.apply_targets(
            session, tf, list(rows_ms.value), undo=coll_undo_stack)
        tail = ("" if not msgs else " — "
                + "; ".join(md_safe(m) for m in msgs[:3]))
        coll_flash.object = (
            f"✓ **Applied {applied} row(s)** from `{md_safe(tf.path.name)}` "
            f"to {md_safe(tf.lob)}{tail}" if applied
            else f"⚠ **Nothing applied**{tail}")
        _undo_sync()
        _preview()                       # the file now matches the book

    apply_btn.on_click(_apply_targets)

    def _undo_sync():
        d = coll_undo_stack.peek()
        coll_undo.label = f"↶ Undo: {d}" if d else "↶ Undo"
        coll_undo.disabled = d is None

    def _coll_undo(_event):
        n, skipped = coll_undo_stack.pop_apply(session)
        msg = f"restored {n} row(s)"
        if skipped:
            msg += " — " + "; ".join(skipped[:2])
        coll_flash.object = ("↶ " if not skipped else "⚠ ") + \
            f"**{md_safe(msg)}**"
        _undo_sync()
        _preview()

    coll_undo.on_click(_coll_undo)

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
        pn.pane.Markdown(
            "### Field targets\n*The small file that goes OUT: one workbook "
            "per line, one tab per business unit, and exactly one editable "
            "column — the net rate target. Type a target in Excel and the "
            "plan loss ratios move, because the engine's own constants are "
            "baked in. Dates are baked too: regenerate to move one.*"),
        targets_btn, targets_flash,
        pn.pane.Markdown(
            "**Collect them back** — *read the returned files, see exactly "
            "what each person changed, apply the rows you approve. Blank "
            "means 'no net selection' (the logged program), not zero.*"),
        pn.Row(folder_in, scan_btn, sizing_mode="stretch_width"),
        coll_sel, coll_preview, rows_ms,
        pn.Row(apply_btn, coll_undo, sizing_mode="stretch_width"),
        coll_flash,
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
                        "book": book_btn, "targets": targets_btn,
                        "scan": scan_btn, "collect": coll_sel,
                        "rows": rows_ms, "apply": apply_btn,
                        "coll_undo": coll_undo, "folder": folder_in},
            "collect": {"scan": _scan, "preview": _preview,
                        "apply": _apply_targets, "held": coll_held,
                        "undo": coll_undo_stack},
            "flashes": {"save": save_flash, "load": load_flash,
                        "gen": gen_flash, "log": log_md,
                        "targets": targets_flash, "collect": coll_flash,
                        "preview": coll_preview},
            "summary": summary}
