"""One combo under the microscope — P1: scenario loading + the live bridge.

The bridge card renders the same chain the workbook's Bridge tab shows —
Projected LR x A_rate x A_mod x A_other = CY plan LR, with per-step point
contributions — from a fresh engine result, in place, on every input change
(debounced). Waves P3+ add the parallelogram, the monthly walk and the
delivery band around it.
"""
from __future__ import annotations

import panel as pn

from app import compute, importers
from app.glue.bindings import WatcherBag, debounce, sync_options
from app.glue.engineio import run_async
from app.glue.format import DASH, fmt_dollar, fmt_idx, fmt_pct, md_safe
from app.glue.theme import FAIL_RED, NAVY, PASS_GREEN, STEEL

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
    row = scn.row(key) or {}
    a_other = row.get("a_other")
    a_other = float(a_other) if isinstance(a_other, (int, float)) else 1.0
    trend = row.get("trend")
    trend = (float(trend) if isinstance(trend, (int, float))
             else scn.trend_default)
    net = row.get("netp") is not None or row.get("netp1") is not None
    p = scn.plan_year

    steps_p = [("× Rate earn-in (A_rate)", res.a_rate_p),
               ("× Mod drift (A_mod)", res.a_mod_p),
               ("× Other", a_other)]
    rows_p, tot_p = _chain_rows(res.lr_current, steps_p, "")
    steps_p1 = [(f"× Net trend ({fmt_pct(trend, signed=True)})", 1.0 + trend),
                ("× Rate earn-in (A_rate, P+1)", res.a_rate_p1),
                ("× Mod drift (A_mod, P+1)", res.a_mod_p1),
                ("× Other", a_other)]
    rows_p1, tot_p1 = _chain_rows(res.lr_current, steps_p1, "")

    badge = ("<span class='plw-badge net'>NET RATE SELECTED</span>" if net
             else "")

    def table(title, rows, total, total_label):
        body = "".join(
            f"<tr><td>{md_safe(lbl)}</td><td>{fmt_idx(f)}</td>"
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

    def _render(*_events):
        scn = session.page.config
        if scn is None or not combo_sel.value:
            card.object = _bridge_html("", scn, None)
            return
        card.object = _bridge_html(combo_sel.value, scn,
                                   compute.result_for(session,
                                                      combo_sel.value))

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
    main = pn.Column(card, sizing_mode="stretch_both")
    return {"main": main, "sidebar": sidebar, "bag": bag}
