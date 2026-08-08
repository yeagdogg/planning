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
from app.glue.format import DASH, fmt_dollar, fmt_idx, fmt_pct, md_safe
from app.glue.theme import FAIL_RED, NAVY, PASS_GREEN, STEEL
from app.views import bridge as v_bridge
from app.views import delivery as v_delivery
from app.views import earning as v_earning

_CARD_CSS = f"""
<style>
.plw-bridge {{ font-family: var(--body-font, Segoe UI), sans-serif; }}
.plw-bridge h3 {{ color: {NAVY}; margin: 0 0 2px 0; }}
.plw-bridge .sub {{ color: #666; font-size: 0.85em; margin-bottom: 10px; }}
.plw-bridge table {{ border-collapse: collapse; min-width: 460px; }}
.plw-bridge td, .plw-bridge th {{ padding: 5px 14px; text-align: right;
  border-bottom: 1px solid #e3e6ea; font-variant-numeric: tabular-nums; }}
.plw-bridge td:first-child {{ text-align: left; }}
.plw-bridge tr.total td {{ border-top: 2px solid {NAVY}; font-weight: 700;
  color: {NAVY}; }}
.plw-bridge .pts {{ color: #777; }}
.plw-badge {{ display: inline-block; padding: 1px 8px; border-radius: 9px;
  font-size: 0.78em; color: #fff; background: {STEEL};
  vertical-align: middle; margin-left: 8px; }}
.plw-badge.net {{ background: {PASS_GREEN}; }}
.plw-err {{ color: {FAIL_RED}; font-weight: 600; }}
</style>
"""


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
        return (_CARD_CSS + "<div class='plw-bridge'><h3>No result</h3>"
                "<div class='sub'>Open a workbook and pick a combo.</div></div>")
    if compute.is_error(res):
        return (_CARD_CSS + f"<div class='plw-bridge'><h3>{md_safe(key)}</h3>"
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
    return (_CARD_CSS + f"<div class='plw-bridge'><h3>{md_safe(key)}{badge}"
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

    # ---- main: the live bridge card ---------------------------------------
    card = pn.pane.HTML(_bridge_html("", session.page.config, None),
                        sizing_mode="stretch_width")

    # ---- deep-dive visuals (P3) --------------------------------------------
    wf_p = pn.pane.Bokeh(sizing_mode="stretch_width")
    wf_p1 = pn.pane.Bokeh(sizing_mode="stretch_width")
    lag_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    para_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    race_md = pn.pane.Markdown("", sizing_mode="stretch_width")
    walk_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    band_pane = pn.pane.HoloViews(sizing_mode="stretch_width")
    vis_note = pn.pane.Markdown("", sizing_mode="stretch_width")

    tabs = pn.Tabs(
        ("Waterfall", pn.Row(wf_p, wf_p1, sizing_mode="stretch_width")),
        ("Earning", pn.Column(lag_pane, para_pane,
                              sizing_mode="stretch_width")),
        ("Monthly walk", pn.Column(race_md, walk_pane,
                                   sizing_mode="stretch_width")),
        ("Delivery", pn.Column(band_pane, sizing_mode="stretch_width")),
        dynamic=False,                       # dynamic=True makes deaf cards
        sizing_mode="stretch_width")

    def _clear_visuals(msg: str = ""):
        for pane in (wf_p, wf_p1, lag_pane, para_pane, walk_pane, band_pane):
            pane.object = None
        race_md.object = ""
        vis_note.object = msg

    def _render_visuals(key, scn, res):
        row = scn.row(key)
        ci = compute.combo_inputs(scn, row)
        eng = engine.MonthlyEngine(scn.plan_year, ci)
        p = scn.plan_year
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
        flow = engine.lr_flow_by_month(p, ci)
        race_md.object = "**" + v_delivery.race_sentence(flow,
                                                         ci.net_trend) + "**"
        target = row.get("target")
        walk_pane.object = v_delivery.walk(
            flow, target if isinstance(target, (int, float)) else None)
        band_pane.object = v_delivery.band(eng, p)
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
    _render()

    sidebar = pn.Column(
        pn.pane.Markdown("**Scenario**"),
        pick, open_btn, status,
        pn.layout.Divider(),
        combo_sel,
        sizing_mode="stretch_width")
    main = pn.Column(card, vis_note, tabs, sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "combo_sel": combo_sel,
            "panes": {"waterfall_p": wf_p, "waterfall_p1": wf_p1,
                      "earn_lag": lag_pane, "parallelogram": para_pane,
                      "race": race_md, "walk": walk_pane, "band": band_pane,
                      "note": vis_note}}
