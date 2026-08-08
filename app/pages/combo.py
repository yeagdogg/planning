"""One combo under the microscope — the live bridge (P1) and its deep-dive
visuals (P3).

The bridge card renders the same chain the workbook's Bridge tab shows —
Projected LR x A_rate x A_mod x A_other = CY plan LR, with per-step point
contributions — from a fresh engine result, in place, on every input change
(debounced). Below it, four tabs draw the same numbers: the chain as
waterfalls, the earning parallelogram + earn-in lag, the monthly LR walk
(``engine.lr_flow_by_month``, D93) and the ec(k,P) delivery band. One
definition of the chain (`views.bridge.chain`) feeds card and charts alike.
"""
from __future__ import annotations

import panel as pn

from src import engine

from app import compute, importers
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.engineio import run_async
from app.glue.exhibit import chip, chips_html, html_pane as ex_html_pane
from app.glue.format import DASH, TAB_SPECS, fmt_dollar, fmt_idx, fmt_pct, \
    md_safe
from app.glue.theme import FAIL_RED, NAVY, PASS_GREEN, STEEL, TABLE_CSS
from app.paste import coerce
from app.views import bridge as v_bridge
from app.views import delivery as v_delivery
from app.views import earning as v_earning
from app.views import flow as v_flow
from app.views import netdeliv as v_nd
from app.views import onlevel as v_onlevel

def _chain_rows(start, steps, total_label):
    """(label, factor, running) rows for one bridge chain plus its total."""
    rows, run = [], start
    for label, factor in steps:
        before = run
        run = run * factor
        rows.append((label, factor, (run - before) * 100.0))
    return rows, run


def _bridge_html(key: str, scn, res) -> str:
    if res is None:
        return ("<div class='plw-bridge'><h3>No result</h3>"
                "<div class='sub'>Open a workbook and pick a combo.</div></div>")
    if compute.is_error(res):
        return (f"<div class='plw-bridge'><h3>{md_safe(key)}</h3>"
                f"<div class='plw-err'>engine rejected this row: "
                f"{md_safe(res[1])}</div></div>")
    p = scn.plan_year
    start, steps_p, steps_p1, net = v_bridge.chain(scn, key, res)
    rows_p, tot_p = _chain_rows(start, steps_p, "")
    rows_p1, tot_p1 = _chain_rows(start, steps_p1, "")

    badge = ("<span class='plw-badge net'>NET RATE SELECTED</span>" if net
             else "")

    def table(title, rows, total, total_label):
        body = "".join(
            f"<tr><td>× {md_safe(lbl)}</td><td>{fmt_idx(f)}</td>"
            f"<td class='pts'>{pts:+.1f} pts</td></tr>"
            for lbl, f, pts in rows)
        return (f"<table><tr><th>{title}</th><th>factor</th><th></th></tr>"
                f"<tr><td>Projected LR (current rate level)</td>"
                f"<td>{fmt_pct(res.lr_current)}</td><td></td></tr>"
                f"{body}"
                f"<tr class='total'><td>{total_label}</td>"
                f"<td>{fmt_pct(total)}</td><td></td></tr></table>")

    ep = scn.ep(key)
    sub = (f"{md_safe(scn.lob)} · plan year {p} · term {scn.term_months} mo · "
           f"EP {fmt_dollar(ep) if ep else DASH} · "
           f"CRL_ind {fmt_idx(res.crl_ind)} · "
           f"E_CY({p}) {fmt_idx(res.e_cy[p])}")
    return (f"<div class='plw-bridge'><h3>{md_safe(key)}{badge}"
            f"</h3><div class='sub'>{sub}</div>"
            + table(f"CY {p}", rows_p, tot_p, f"CY {p} plan loss ratio")
            + "<div style='height:12px'></div>"
            + table(f"CY {p + 1}", rows_p1, tot_p1,
                    f"CY {p + 1} plan loss ratio")
            + "</div>")


def build(session):
    bag = WatcherBag()

    # ---- sidebar: open a workbook + pick a combo ---------------------------
    choices = importers.workbook_choices()
    pick = pn.widgets.Select(label="Workbook", options=choices,
                             sizing_mode="stretch_width")
    open_btn = pn.widgets.Button(label="Open workbook", button_type="primary",
                                 sizing_mode="stretch_width")
    status = pn.pane.Markdown("", sizing_mode="stretch_width")

    def _open(_event):
        path = pick.value
        if not path:
            return
        open_btn.loading = True

        def _done(scn):
            open_btn.loading = False
            session.replace_config(scn)
            status.object = (f"**{md_safe(scn.name)}** — "
                             f"{len(scn.combo_keys())} combos, "
                             f"term {scn.term_months} mo")
            pn.state.notifications.success(
                f"Loaded {len(scn.combo_keys())} combos from "
                f"{md_safe(scn.name)}", duration=4000)

        def _err(e):
            open_btn.loading = False
            status.object = f"**Load failed** — {md_safe(e)}"

        run_async(lambda: importers.from_workbook(path), _done, _err)

    open_btn.on_click(_open)

    combo_sel = pn.widgets.Select(label="BU | State", options=[],
                                  sizing_mode="stretch_width")
    sync_options(combo_sel, lambda: list(session.ctx.combos),
                 session.ctx.param.data_rev, bag=bag)

    def _adopt_focus(*_events):
        """One-shot adoption of a Book click-through (P4): registered AFTER
        sync_options, so the options are already fresh when this runs."""
        want = getattr(session.ctx, "focus", None)
        if want and want in (combo_sel.options or []):
            session.ctx.focus = None
            combo_sel.value = want

    bag.watch(session.ctx, _adopt_focus, "data_rev")
    # ...and once at build: a FIRST-EVER click-through lazy-builds this page
    # AFTER the Book already set focus and bumped data_rev, so the watcher
    # alone would miss it (found live, then pinned by the lazy-order test)
    _adopt_focus()

    # ---- main: the live bridge card ---------------------------------------
    card = ex_html_pane()
    card.object = _bridge_html("", session.page.config, None)

    # ---- deep-dive visuals (P3 + the W2d on-level rectangle) ---------------
    chips_pane = ex_html_pane()
    ol_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    ol_note = pn.pane.Markdown("", sizing_mode="stretch_width")
    # W3b interaction: the P-1 strip halves horizontal resolution, so it
    # is a choice; net combos can flip between the asserted net path and
    # the logged program — the two geometries ARE the program-vs-asserted
    # story, and the note quantifies the gap in points.
    ol_window = pn.widgets.Checkbox(
        name="Show CY P−1 (36-month window)", value=True)
    ol_basis = pn.widgets.RadioButtonGroup(
        options=["Net path (asserted)", "Logged program (D65)"],
        value="Net path (asserted)", visible=False)
    wf_p = pn.pane.Bokeh(sizing_mode="stretch_width")
    wf_p1 = pn.pane.Bokeh(sizing_mode="stretch_width")
    lag_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    para_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    race_md = pn.pane.Markdown("", sizing_mode="stretch_width")
    walk_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    band_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    vis_note = pn.pane.Markdown("", sizing_mode="stretch_width")

    # ---- W3d: the Flow tab (per-combo Flow Dashboard) ----------------------
    fl_chips = ex_html_pane()
    fl_idx = pn.pane.HoloViews(sizing_mode="stretch_width")
    fl_yoy = pn.pane.HoloViews(sizing_mode="stretch_width")
    fl_run = pn.pane.HoloViews(sizing_mode="stretch_width")
    fl_mod = pn.pane.HoloViews(sizing_mode="stretch_width")
    fl_deliv = pn.pane.HoloViews(sizing_mode="stretch_width")
    fl_ledger = pn.pane.HTML("", sizing_mode="stretch_width")
    fl_note = pn.pane.Markdown("", sizing_mode="stretch_width")

    # ---- W3d: the walk, deeper (chips + table + reconciliation) ------------
    import pandas as pd
    from bokeh.models.widgets.tables import NumberFormatter

    walk_chips = ex_html_pane()
    _WALK_COLS = ["label", "trend_factor", "a_rate", "a_mod", "lr",
                  "vs_target", "den", "weight"]
    _WALK_KINDS = {"trend_factor": "idx", "a_rate": "idx", "a_mod": "idx",
                   "lr": "pct", "vs_target": "pts", "den": "idx",
                   "weight": "idx"}
    walk_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=_WALK_COLS),
        disabled=True, show_index=False,
        titles={"label": "Month", "trend_factor": "Trend factor",
                "a_rate": "Rate earn-in", "a_mod": "Mod earn-in",
                "lr": "Plan LR", "vs_target": "vs target (pts)",
                "den": "Earned exposure", "weight": "Premium weight"},
        formatters={k: NumberFormatter(format=TAB_SPECS[kind],
                                       nan_format="—", null_format="—")
                    for k, kind in _WALK_KINDS.items()},
        text_align={k: "right" for k in _WALK_KINDS},
        layout="fit_data_table", max_height=760,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")
    recon_pane = pn.pane.HTML("", sizing_mode="stretch_width")

    def _walk_styles(d):
        css = pd.DataFrame("", index=d.index, columns=d.columns)
        css.iloc[12:, :] = "background-color: #EDF1F6"   # P+1 tinted
        return css

    # ---- W3f: the Net Delivery microscope (net combos' solve surface) ------
    nd_eff = pn.widgets.DatePicker(name="Filing effective date")
    nd_ovr = pn.widgets.TextInput(
        name="Filed % override",
        placeholder="blank = planned filing / suggested")
    nd_chips = ex_html_pane()
    nd_share = pn.pane.HoloViews(sizing_mode="stretch_width")
    _MIC_KINDS = {"w": "idx", "rate_leg": "pct_signed",
                  "price_req": "pct_signed", "m_required": "mod",
                  "m_proj": "mod", "gap": "pts", "delivered": "pct_signed"}
    _MIC_TITLES = {"label": "Month", "w": "Written weight",
                   "rate_leg": "YoY rate leg",
                   "price_req": "Required pricing leg",
                   "m_required": "Required written mod",
                   "m_proj": "Projected written mod",
                   "gap": "Gap (pts)", "delivered": "Delivered @ proj",
                   "m_base": "Mod base (plan-yr)"}
    micro_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["label"] + list(_MIC_KINDS)),
        disabled=True, show_index=False, titles=_MIC_TITLES,
        formatters={k: NumberFormatter(format=TAB_SPECS[kind],
                                       nan_format="—", null_format="—")
                    for k, kind in _MIC_KINDS.items()},
        text_align={k: "right" for k in _MIC_KINDS},
        layout="fit_data_table", max_height=420,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index", "share_rate", "share_price"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")
    _P1_KINDS = {"w": "idx", "rate_leg": "pct_signed",
                 "price_req": "pct_signed", "m_required": "mod",
                 "m_base": "mod"}
    p1_cap = pn.pane.Markdown("", sizing_mode="stretch_width")
    p1_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["label"] + list(_P1_KINDS)),
        disabled=True, show_index=False, titles=_MIC_TITLES,
        formatters={k: NumberFormatter(format=TAB_SPECS[kind],
                                       nan_format="—", null_format="—")
                    for k, kind in _P1_KINDS.items()},
        text_align={k: "right" for k in _P1_KINDS},
        layout="fit_data_table", max_height=420,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")
    nd_note = pn.pane.Markdown("", sizing_mode="stretch_width")
    nd_hint = pn.pane.Markdown("", sizing_mode="stretch_width")
    nd_col = pn.Column(pn.Row(nd_eff, nd_ovr), nd_chips, nd_share,
                       micro_tbl, p1_cap, p1_tbl, nd_note,
                       visible=False, sizing_mode="stretch_width")

    tabs = pn.Tabs(
        # the whole reason this app exists, first (W2d, user direction)
        ("On-level", pn.Column(pn.Row(ol_window, ol_basis),
                               ol_pane, ol_note,
                               sizing_mode="stretch_width")),
        ("Waterfall", pn.Row(wf_p, wf_p1, sizing_mode="stretch_width")),
        ("Earning", pn.Column(lag_pane, para_pane,
                              sizing_mode="stretch_width")),
        ("Flow", pn.Column(fl_chips, fl_idx, fl_yoy, fl_run, fl_mod,
                           fl_deliv, fl_ledger, fl_note,
                           sizing_mode="stretch_width")),
        ("Monthly walk", pn.Column(race_md, walk_chips, walk_pane,
                                   walk_tbl, recon_pane,
                                   sizing_mode="stretch_width")),
        ("Delivery", pn.Column(band_pane, nd_hint, nd_col,
                               sizing_mode="stretch_width")),
        dynamic=False,                       # dynamic=True makes deaf cards
        sizing_mode="stretch_width")

    def _clear_visuals(msg: str = ""):
        for pane in (ol_pane, wf_p, wf_p1, lag_pane, para_pane, walk_pane,
                     band_pane, fl_idx, fl_yoy, fl_run, fl_mod, fl_deliv):
            pane.object = None
        race_md.object = ""
        ol_note.object = ""
        ol_basis.visible = False
        chips_pane.object = ""
        for h in (fl_chips, fl_ledger, walk_chips, recon_pane):
            h.object = ""
        fl_note.object = ""
        walk_tbl.value = walk_tbl.value.iloc[0:0]
        nd_col.visible = False
        nd_hint.object = ""
        nd_chips.object = ""
        nd_note.object = ""
        nd_share.object = None
        micro_tbl.value = micro_tbl.value.iloc[0:0]
        p1_tbl.value = p1_tbl.value.iloc[0:0]
        vis_note.object = msg

    def _render_visuals(key, scn, res):
        row = scn.row(key)
        ci = compute.combo_inputs(scn, row)
        eng = engine.MonthlyEngine(scn.plan_year, ci)
        p = scn.plan_year

        # KPI chips: verbatim EngineResult fields + the one disclosed
        # subtraction (vs target, in points)
        target = row.get("target")
        vs = ((res.cy_lr_p - float(target)) * 100.0
              if isinstance(target, (int, float)) else None)
        chips_pane.object = chips_html(
            chip(f"CY {p} plan LR", fmt_pct(res.cy_lr_p)),
            chip("vs target", DASH if vs is None else f"{vs:+.1f} pts",
                 sub=(DASH if target is None
                      else f"target {fmt_pct(target)}"),
                 bad=bool(vs is not None and vs > 0)),
            chip(f"CY {p + 1} plan LR", fmt_pct(res.cy_lr_p1),
                 sub="indicative"),
            chip("earned chg vs indication",
                 fmt_pct(res.earned_chg_vs_ind, signed=True)),
            chip(f"YoY earned {p - 1}→{p}",
                 fmt_pct(res.yoy_earned_p, signed=True)))

        ol_basis.visible = bool(ci.net_mode)
        basis = ("program" if ci.net_mode
                 and ol_basis.value.startswith("Logged") else "auto")
        fr = v_onlevel.onlevel_frame(
            ci, res, window=36 if ol_window.value else 24, basis=basis)
        ol_pane.object = v_onlevel.onlevel(ci, res, frame=fr)
        notes = list(fr.notes)
        if ci.net_mode:
            prog = compute.program_results(session).get(
                scn.lob, {}).get(key)
            if prog is not None:
                gap = (prog[0] - res.cy_lr_p) * 100.0
                notes.insert(0, (
                    f"Program vs asserted: the logged program would plan "
                    f"CY {p} at {fmt_pct(prog[0])} vs the net path's "
                    f"{fmt_pct(res.cy_lr_p)} ({gap:+.1f} pts) — flip the "
                    f"basis above to see both geometries."))
        ol_note.object = ("\n".join(f"*{md_safe(n)}*" for n in notes)
                          if notes else "")

        start, steps_p, steps_p1, _net = v_bridge.chain(scn, key, res)
        wf_p.object = v_bridge.waterfall_figure(
            f"CY {p}: how the plan LR is built",
            v_bridge.waterfall_bars("Projected LR", start, steps_p,
                                    f"CY {p} plan LR"))
        wf_p1.object = v_bridge.waterfall_figure(
            f"CY {p + 1} (indicative)",
            v_bridge.waterfall_bars("Projected LR", start, steps_p1,
                                    f"CY {p + 1} plan LR"))
        lag_pane.object = v_earning.earn_lag(res)
        para_pane.object = v_earning.parallelogram(eng, p)

        # ---- the Flow tab (W3d): the Excel Flow Dashboard, live ------------
        pf = engine.program_flow_by_month(p, ci)
        fdf = v_flow.dashboard_frame(res, pf, ci)
        dec_run = float(fdf.loc[fdf["mi"] == engine.month_index(p, 12),
                                "runway"].iloc[0])
        fl_chips.object = chips_html(
            chip(f"avg YoY delivered, CY {p}",
                 fmt_pct(pf.avg_delivered_ratio - 1.0, signed=True),
                 sub="w-weighted mean of monthly ratios"),
            chip(f"avg YoY delivered, CY {p - 1}",
                 fmt_pct(pf.avg_delivered_ratio_prior - 1.0, signed=True),
                 sub="the year currently flowing"),
            chip(f"carryover into CY {p + 1}",
                 fmt_pct(res.yoy_earned_p1, signed=True),
                 sub="earned rate chg, next year"),
            chip(f"runway at Dec {p}", fmt_pct(dec_run, signed=True),
                 sub="written / earned − 1"))
        fl_idx.object = v_flow.index_chart(fdf, p)
        fl_yoy.object = v_flow.yoy_bars(fdf, p, net_mode=ci.net_mode)
        fl_run.object = v_flow.runway_line(fdf, p)
        mp = v_flow.mod_price_chart(fdf, ci, p)
        fl_mod.object = mp
        fl_mod.visible = mp is not None
        fl_deliv.object = v_flow.delivery_chart(fdf, ci, p)
        ldf, totals = v_flow.ledger_frame(eng, ci, res)
        fl_ledger.object = v_flow.ledger_html(ldf, totals, p)
        if ci.net_mode:
            fl_note.object = ("*Mod path & price index not drawn: the net "
                              "path already combines rate and price — one "
                              "combined series, never two.*")
        elif not ci.mod_adjustment_enabled:
            fl_note.object = ("*Mod path & price index not drawn: the mod "
                              "adjustment is OFF — the program is "
                              "rate-only.*")
        else:
            fl_note.object = ""

        # ---- the walk, deeper (W3d) ----------------------------------------
        flow = engine.lr_flow_by_month(p, ci)
        race_md.object = "**" + v_delivery.race_sentence(flow,
                                                         ci.net_trend) + "**"
        target = row.get("target")
        t_ok = target if isinstance(target, (int, float)) else None
        pts, bad = v_delivery.headroom(flow, ci.net_trend)
        walk_chips.object = chips_html(
            chip("breakeven net trend",
                 DASH if flow.breakeven_trend is None
                 else fmt_pct(flow.breakeven_trend)),
            chip("headroom", DASH if pts is None else f"{pts:+.1f} pts",
                 sub="trend − breakeven; positive = trend wins", bad=bad),
            chip("trend in force", fmt_pct(ci.net_trend, signed=True)))
        walk_pane.object = v_delivery.walk(flow, t_ok)
        wdf = v_delivery.walk_table_frame(flow, t_ok)
        walk_tbl.value = wdf
        walk_tbl.style = wdf.style.apply(_walk_styles, axis=None)
        recon_pane.object = v_delivery.recon_html(flow)
        band_pane.object = v_delivery.band(eng, p)

        # ---- the Net Delivery microscope (W3f) -----------------------------
        nd_col.visible = bool(ci.net_mode)
        nd_hint.object = ("" if ci.net_mode else
                          "*The Net Delivery microscope appears for combos "
                          "with a net selection — this combo runs the "
                          "explicit program (the band above is its "
                          "delivery geometry).*")
        if ci.net_mode:
            try:
                raw = (nd_ovr.value or "").strip()
                ovr = coerce("pct", raw) if raw else None
                mic = v_nd.microscope(p, ci, nd_eff.value, ovr)
            except ValueError as e:
                nd_note.object = f"⚠ **{md_safe(e)}**"
                nd_chips.object = ""
                nd_share.object = None
                micro_tbl.value = micro_tbl.value.iloc[0:0]
                p1_tbl.value = p1_tbl.value.iloc[0:0]
                p1_cap.object = ""
            else:
                step = mic["step"]
                multi = (f" — ⚠ MULTI-PLANNED ({mic['multi_planned']})"
                         if mic["multi_planned"] > 1 else "")
                nd_chips.object = chips_html(
                    chip("solving for", fmt_pct(mic["x"], signed=True),
                         sub=f"the asserted net target, CY {p}"),
                    chip("suggested change",
                         fmt_pct(mic["suggested"], signed=True),
                         sub=f"@ {mic['eff_date']}"),
                    chip("change in force", fmt_pct(mic["r"], signed=True),
                         sub=mic["source"] + multi),
                    chip("filed equivalent",
                         fmt_pct(mic["filed_equiv"], signed=True),
                         sub=f"achievement {mic['ach']:.0%}"),
                    chip("required M_1′",
                         "n/a" if mic["m1_prime"] is None
                         else f"{mic['m1_prime']:.3f}",
                         sub=(mic["m1_reason"] or
                              "year-end written mod LEVEL")),
                    chip("required mod step",
                         DASH if step is None
                         else fmt_pct(step.required_step, signed=True),
                         sub=(mic["step_reason"] if step is None else
                              f"direct {step.directed_equivalent:+.1%} · "
                              f"post-share {step.post_share:.0%}"),
                         bad=bool(step is not None and not step.feasible)))
                nd_share.object = v_nd.share_chart(mic["rows"], mic["x"],
                                                   p)
                micro_tbl.value = mic["rows"]
                p1_cap.object = (
                    f"**CY {p + 1} carryover** — no second filing "
                    f"assumed; target used: "
                    f"{fmt_pct(mic['x1'], signed=True)} "
                    f"({mic['p1_src']})")
                p1_tbl.value = mic["p1"]
                warn = "; ".join(mic["warnings"])
                nd_note.object = f"*⚠ {md_safe(warn)}*" if warn else ""
        vis_note.object = ""

    def _render(*_events):
        scn = session.page.config
        key = combo_sel.value
        if scn is None or not key:
            card.object = _bridge_html("", scn, None)
            _clear_visuals()
            return
        res = compute.result_for(session, key)
        card.object = _bridge_html(key, scn, res)
        if res is None or compute.is_error(res):
            _clear_visuals()
            return
        try:
            _render_visuals(key, scn, res)
        except Exception as e:                          # noqa: BLE001
            # a half-typed row must never blank the page — the card above
            # still explains itself; say why the charts sat this one out
            _clear_visuals(f"⚠ visuals unavailable: {md_safe(e)}")

    # grid keystrokes arrive as bus bumps (P2); coalesce bursts into one
    # recompute+render shortly after the last event
    rail_rev = debounce(session.bus.param.rev, delay_ms=250, bag=bag)
    bag.watch(rail_rev, _render, "rev")
    bag.watch(session.ctx, _render, "data_rev")
    combo_sel.param.watch(_render, "value")
    ol_window.param.watch(_render, "value")
    ol_basis.param.watch(_render, "value")
    nd_eff.param.watch(_render, "value")
    nd_ovr.param.watch(_render, "value")
    _render()

    sidebar = pn.Column(
        pn.pane.Markdown("**Scenario**"),
        pick, open_btn, status,
        pn.layout.Divider(),
        combo_sel,
        sizing_mode="stretch_width")
    main = pn.Column(chips_pane, card, vis_note, tabs,
                     sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "card": card, "combo_sel": combo_sel,
            "panes": {"waterfall_p": wf_p, "waterfall_p1": wf_p1,
                      "earn_lag": lag_pane, "parallelogram": para_pane,
                      "race": race_md, "walk": walk_pane, "band": band_pane,
                      "onlevel": ol_pane, "chips": chips_pane,
                      "note": vis_note,
                      "flow_chips": fl_chips, "flow_index": fl_idx,
                      "flow_yoy": fl_yoy, "flow_runway": fl_run,
                      "flow_mod": fl_mod, "flow_delivery": fl_deliv,
                      "flow_ledger": fl_ledger, "flow_note": fl_note,
                      "walk_chips": walk_chips, "recon": recon_pane},
            "walk_tbl": walk_tbl,
            "nd": {"eff": nd_eff, "override": nd_ovr, "chips": nd_chips,
                   "share": nd_share, "table": micro_tbl, "p1": p1_tbl,
                   "p1_cap": p1_cap, "note": nd_note, "col": nd_col,
                   "hint": nd_hint},
            "ol_window": ol_window, "ol_basis": ol_basis,
            "ol_note": ol_note}
