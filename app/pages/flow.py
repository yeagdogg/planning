"""The Flow page — the aggregable half of the flow exhibits (W3e).

Doctrine, the workbook's own split: **Flow reports, Summary edits, Combo
prescribes.** This page never takes input — it shows what the logged
program delivers, state by state and month by month, against whatever
the Summary's levers currently assert. The live coupling is free: an
edit anywhere bumps the session bus, ``compute.flow_results`` recomputes
the one line that moved, and every grid here re-renders.

Aggregation rules, encoded not hoped for: combination is the engine's
own EP×w(m) rule (``combine_rows`` is pinned to
``engine.combined_flow_by_month`` at 1e-15); delivered is the exact
statistic and weighted rate × weighted mod need not reproduce it (mix is
disclosed, never normalised); the mod leg dashes when nothing in view
carries the adjustment; net targets average over NET combos only; the
required pricing leg populates ONLY where a state resolves to exactly
one net combo (a ratio of weighted means is not the weighted mean of the
ratios); and there is NO book-level LR walk — earning patterns never
aggregate across terms. Lines on a different plan year are excluded and
named, never combined.
"""
from __future__ import annotations

import pandas as pd
import panel as pn

from app import compute
from app.glue.bindings import WatcherBag, debounce, gate_hidden, \
    sync_options
from app.glue.exhibit import chip, chips_html, echo_html, echo_pane, \
    exhibit_header, note_list
from app.glue.format import DASH, TAB_SPECS, fmt_count, fmt_dollar, \
    fmt_pct, md_safe
from app.glue.theme import TABLE_CSS
from app.views import flow as v_flow

_MONTH_ABBR = ["Jan", "Feb", "Mar", "Apr", "May", "Jun",
               "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]

_NOTES = [
    "Averages are w-weighted means of the monthly YoY ratios (the Net "
    "Delivery delivered-average convention) — NOT the E_CY aggregate "
    "convention calendar-year earnings use.",
    "Delivered is the exact statistic: weighted rate × weighted mod need "
    "not reproduce it under mix — the difference is disclosed, never "
    "normalised away (D60).",
    "The bottom row is an average of the state rows, weighted by adjusted "
    "plan EP × written weight — an average, not a total.",
    "The mod leg shows — where no combo in view carries the mod "
    "adjustment; those combos still deliver their rate leg.",
    "Net targets average over NET combos only — a combo that asserts "
    "nothing is not a target of zero.",
    "The required pricing leg populates only where a state resolves to "
    "exactly ONE net combo: a ratio of weighted means is not the weighted "
    "mean of the ratios, so it is dashed rather than approximated.",
    "There is no book-level LR walk, deliberately: earning patterns never "
    "aggregate across lines with different terms. Written-basis legs do "
    "— which is what makes this page valid across the whole book.",
    "Click a state row that resolves to a single combo to open it on the "
    "Combo page.",
]


def _fmt(kinds: dict) -> dict:
    from bokeh.models.widgets.tables import NumberFormatter
    return {k: NumberFormatter(format=TAB_SPECS[kind], nan_format="—",
                               null_format="—")
            for k, kind in kinds.items()}


def build(session):
    bag = WatcherBag()

    # ---- sidebar -----------------------------------------------------------
    line_f = pn.widgets.MultiChoice(label="Lines", options=[],
                                    placeholder="all lines",
                                    sizing_mode="stretch_width")
    bu_f = pn.widgets.MultiChoice(label="Business units", options=[],
                                  placeholder="all BUs",
                                  sizing_mode="stretch_width")
    show_prior = pn.widgets.Checkbox(
        name="Show prior year (CY P−1 columns)", value=False)

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
        "Flow — what the program delivers, month by month",
        "The descriptive half: the YoY legs the logged program writes "
        "into the book, state by state, and delivery against the net "
        "targets the Summary asserts. Edits anywhere move this page "
        "live; it never takes input itself.")
    echo = echo_pane()
    chips_pane = pn.pane.HTML("", sizing_mode="stretch_width")

    leg_radio = pn.widgets.RadioButtonGroup(
        options=["delivered", "rate", "mod"], value="delivered")
    nd_radio = pn.widgets.RadioButtonGroup(
        options=["delivered", "rate"], value="delivered")

    _SUM_KINDS = {"ep": "dollar", "rate_pm1": "pct_signed",
                  "mod_pm1": "pct_signed", "del_pm1": "pct_signed",
                  "nettgt": "pct_signed", "rate_p": "pct_signed",
                  "mod_p": "pct_signed", "del_p": "pct_signed",
                  "dnet": "pct_signed", "proggap": "pts",
                  "netcnt": "count", "cnt": "count"}
    sum_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["state"] + list(_SUM_KINDS)),
        disabled=True, show_index=False,
        titles={"state": "State", "ep": "Adj plan EP (000s)",
                "rate_pm1": "Avg rate leg", "mod_pm1": "Avg mod leg",
                "del_pm1": "Avg delivered", "nettgt": "Net target",
                "rate_p": "Avg rate leg", "mod_p": "Avg mod leg",
                "del_p": "Avg delivered", "dnet": "Δ vs net assertion",
                "proggap": "Plan LR gap: program vs asserted (pts)",
                "netcnt": "Net combos", "cnt": "Combos"},
        formatters=_fmt(_SUM_KINDS),
        text_align={k: "right" for k in _SUM_KINDS},
        frozen_columns=["state"],
        layout="fit_data_table", max_height=560,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy",
                       "columnDefaults": {"headerWordWrap": True}},
        sizing_mode="stretch_width")

    # all three year bands (W3h): P−1 behind the sidebar toggle, P, and
    # P+1 — the flow result carries 24 plan-rooted months, and "program
    # flow for 2028" is half the point of a plan-year page
    _GRID_KINDS = {c: "pct_signed"
                   for c in v_flow.grid_cols(prior=True, p1=True)}
    grid_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["state"]
                     + v_flow.grid_cols(prior=True, p1=True)),
        disabled=True, show_index=False,
        formatters=_fmt(_GRID_KINDS),
        text_align={k: "right" for k in _GRID_KINDS},
        frozen_columns=["state"],
        layout="fit_data_table", max_height=840,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")

    _ND_KINDS = {"ep": "dollar", "netcnt": "count", "tgt": "pct_signed",
                 "tgt1": "pct_signed", "delivered": "pct_signed",
                 "gap": "pts", "progbasis": "pct", "proggap": "pts",
                 "suggest": "pct_signed"}
    nd_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["state"] + list(_ND_KINDS) + ["set1"]),
        disabled=True, show_index=False,
        titles={"state": "State", "ep": "Adj plan EP (000s)",
                "netcnt": "Net combos", "tgt": "Net target P",
                "tgt1": "Net target P+1", "set1": "P+1 target set?",
                "delivered": "Avg delivered (as logged)",
                "gap": "Gap to target (pts)",
                "progbasis": "Plan LR — program basis",
                "proggap": "Program vs asserted (pts)",
                "suggest": "Suggested change ° (single net combo, "
                           "default date)"},
        formatters=_fmt(_ND_KINDS),
        text_align={k: "right" for k in _ND_KINDS},
        frozen_columns=["state"],
        layout="fit_data_table", max_height=560,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy",
                       "columnDefaults": {"headerWordWrap": True}},
        sizing_mode="stretch_width")

    nd_grid = pn.widgets.Tabulator(
        pd.DataFrame(columns=["state"] + v_flow.grid_cols(prior=False,
                                                          p1=True)),
        disabled=True, show_index=False,
        formatters=_fmt({c: "pct_signed"
                         for c in v_flow.grid_cols(prior=False, p1=True)}),
        text_align={c: "right"
                    for c in v_flow.grid_cols(prior=False, p1=True)},
        frozen_columns=["state"],
        layout="fit_data_table", max_height=840,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")

    req_tbl = pn.widgets.Tabulator(
        pd.DataFrame(columns=["state"]
                     + [f"pp{j:02d}" for j in range(1, 13)]),
        disabled=True, show_index=False,
        formatters=_fmt({f"pp{j:02d}": "pct_signed"
                         for j in range(1, 13)}),
        text_align={f"pp{j:02d}": "right" for j in range(1, 13)},
        frozen_columns=["state"],
        layout="fit_data_table", max_height=560,
        selectable=False, stylesheets=[TABLE_CSS],
        hidden_columns=["_index"],
        configuration={"clipboard": "copy"},
        sizing_mode="stretch_width")

    flash = pn.pane.Markdown("", sizing_mode="stretch_width")
    holder = {"view": None, "p": None, "avg": None}

    # ---- assembly ----------------------------------------------------------
    def _assemble():
        """The in-view records, state-keyed. Lines on a different plan
        year than the first loaded one are EXCLUDED and named."""
        fr = compute.flow_results(session)
        br = compute.book_results(session)
        view: dict = {}
        p0, excluded = None, []
        for lob, scn in session.page.book.items():
            if line_f.value and lob not in line_f.value:
                continue
            if p0 is None:
                p0 = scn.plan_year
            if scn.plan_year != p0:
                excluded.append(lob)
                continue
            for r in scn.lr_rows:
                bu, st = r.get("bu"), r.get("state")
                if not (bu and st):
                    continue
                if bu_f.value and bu not in bu_f.value:
                    continue
                key = f"{bu}|{st}"
                pf = fr.get(lob, {}).get(key)
                if isinstance(pf, tuple):
                    pf = None                       # engine-rejected row
                res = br.get(lob, {}).get(key)
                lr_p = (None if res is None or compute.is_error(res)
                        else res.cy_lr_p)
                view.setdefault(st, []).append(dict(
                    lob=lob, key=key, bu=bu, ep=r.get("ep"), pf=pf,
                    netp=r.get("netp"), netp1=r.get("netp1"), lr_p=lr_p,
                    scn=scn, row=r))
        return view, p0, excluded

    def _avg_label():
        if bu_f.value:
            return "BU AVG"
        if line_f.value:
            return "LINE AVG"
        return "BOOK AVG"

    def _month_titles(p):
        t = {"state": "State"}
        for j in range(1, 13):
            t[f"pm{j:02d}"] = f"{_MONTH_ABBR[j - 1]} {p - 1}"
            t[f"pp{j:02d}"] = f"{_MONTH_ABBR[j - 1]} {p}"
            t[f"q{j:02d}"] = f"{_MONTH_ABBR[j - 1]} {p + 1}"
        return t

    def _render(*_events):
        if not session.page.book:
            for t in (sum_tbl, grid_tbl, nd_tbl, nd_grid, req_tbl):
                t.value = t.value.iloc[0:0]
            chips_pane.object = ""
            echo.object = echo_html(
                "No lines loaded — Book page → Load every line, paste "
                "masters, or load a scenario.")
            return
        view, p, excluded = _assemble()
        holder["view"], holder["p"] = view, p
        all_flows = [(r["ep"], r["pf"]) for recs in view.values()
                     for r in recs if r["pf"] is not None]
        if not any(ep and ep > 0.0 for ep, _pf in all_flows):
            for t in (sum_tbl, grid_tbl, nd_tbl, nd_grid, req_tbl):
                t.value = t.value.iloc[0:0]
            chips_pane.object = ""
            echo.object = echo_html("Nothing in view carries EP.")
            return
        prog = {(lob, k): v
                for lob, d in compute.program_results(session).items()
                for k, v in d.items()}
        label = _avg_label()

        avg = v_flow.combine_avgs(all_flows)
        holder["avg"] = avg
        netcnt = sum(1 for recs in view.values() for r in recs
                     if r["netp"] is not None)
        gaps = [(r["ep"], prog[(r["lob"], r["key"])][0] - r["lr_p"])
                for recs in view.values() for r in recs
                if (r["lob"], r["key"]) in prog and r["lr_p"] is not None
                and r["ep"]]
        gap = (sum(ep * g for ep, g in gaps) / sum(ep for ep, _g in gaps)
               * 100.0 if gaps else None)
        chips_pane.object = chips_html(
            chip("adjusted plan EP", fmt_dollar(avg["ep"]),
                 sub="in view"),
            chip("net-target combos", fmt_count(netcnt)),
            chip(f"avg YoY delivered, CY {p}",
                 fmt_pct(avg["delivered"], signed=True)),
            chip(f"avg YoY delivered, CY {p - 1}",
                 fmt_pct(avg["delivered_prior"], signed=True),
                 sub="now flowing"),
            chip("program vs asserted",
                 DASH if gap is None else f"{gap:+.1f} pts",
                 sub="EP-weighted over net combos",
                 bad=bool(gap is not None and abs(gap) > 0.5)))

        # -- Program flow tab ------------------------------------------------
        sdf = v_flow.pf_summary_frame(view, prog, label)
        sum_tbl.titles = {**sum_tbl.titles,
                          "rate_pm1": f"Avg rate leg  (CY {p - 1})",
                          "mod_pm1": f"Avg mod leg  (CY {p - 1})",
                          "del_pm1": f"Avg delivered  (CY {p - 1})",
                          "rate_p": f"Avg rate leg  (CY {p})",
                          "mod_p": f"Avg mod leg  (CY {p})",
                          "del_p": f"Avg delivered  (CY {p})"}
        hide = ["_index"] + (["rate_pm1", "mod_pm1", "del_pm1"]
                             if not show_prior.value else [])
        sum_tbl.hidden_columns = hide
        sum_tbl.value = sdf
        sum_tbl.style = sdf.style.apply(
            lambda d: v_flow.year_band_styles(d, []), axis=None)

        leg = {"rate": "rate_leg", "mod": "mod_leg",
               "delivered": "delivered"}[leg_radio.value]
        gdf = v_flow.state_grid_frame(view, all_flows, leg, label,
                                      p1=True, prior=True)
        grid_tbl.titles = _month_titles(p)
        grid_tbl.hidden_columns = ["_index"] + (
            [] if show_prior.value else
            [f"pm{j:02d}" for j in range(1, 13)])
        grid_tbl.value = gdf
        grid_tbl.style = gdf.style.apply(
            lambda d: v_flow.year_band_styles(
                d, [[f"pm{j:02d}" for j in range(1, 13)],
                    [f"pp{j:02d}" for j in range(1, 13)],
                    [f"q{j:02d}" for j in range(1, 13)]]), axis=None)
        grid_tbl.frozen_rows = [len(gdf) - 1] if len(gdf) else []

        # -- Net delivery tab ------------------------------------------------
        # suggested change: single-net-combo states only (W3f) — one
        # closed form per qualifying row, at the Excel default date
        import datetime as dt

        from src import engine
        sugg = {}
        for st, recs in view.items():
            if len(recs) != 1 or recs[0]["netp"] is None:
                continue
            try:
                ci = compute.combo_inputs(recs[0]["scn"], recs[0]["row"])
                planned = engine.planned_change_in_plan_year(p, ci)
                d0 = planned[0] if planned else dt.date(p, 4, 1)
                sugg[st] = engine.suggest_net_rate(p, ci, d0)[0]
            except Exception:                           # noqa: BLE001
                continue                # no exposure after the date, etc.
        ndf = v_flow.nd_summary_frame(view, prog, label, suggest=sugg)
        nd_tbl.value = ndf
        nleg = {"rate": "rate_leg", "delivered": "delivered"}[
            nd_radio.value]
        ngdf = v_flow.state_grid_frame(view, all_flows, nleg, label,
                                       p1=True)
        nd_grid.titles = _month_titles(p)
        nd_grid.value = ngdf
        nd_grid.style = ngdf.style.apply(
            lambda d: v_flow.year_band_styles(
                d, [[f"pp{j:02d}" for j in range(1, 13)],
                    [f"q{j:02d}" for j in range(1, 13)]]), axis=None)
        nd_grid.frozen_rows = [len(ngdf) - 1] if len(ngdf) else []
        rdf = v_flow.required_pricing_frame(view, p)
        req_tbl.titles = _month_titles(p)
        req_tbl.value = rdf

        lines_txt = (", ".join(line_f.value) if line_f.value
                     else "every line")
        bus_txt = (", ".join(bu_f.value) if bu_f.value
                   else "every business unit")
        extra = (f"  |  ⚠ excluded (different plan year): "
                 f"{', '.join(md_safe(x) for x in excluded)}"
                 if excluded else "")
        echo.object = echo_html(
            f"Showing: {lines_txt}  |  {bus_txt}  |  plan year {p}  |  "
            f"{sum(len(v) for v in view.values())} combos across "
            f"{len(view)} states{extra}")

    # ---- click-through -----------------------------------------------------
    def _open_from(tbl):
        def _handler(event):
            try:
                st = str(tbl.value.iloc[int(event.row)]["state"])
            except Exception:                           # noqa: BLE001
                return
            recs = (holder["view"] or {}).get(st, [])
            if len(recs) != 1:
                flash.object = (f"*{md_safe(st)} resolves to "
                                f"{len(recs)} combo(s) — narrow the "
                                f"filters to click through.*")
                return
            session.ctx.focus = recs[0]["key"]
            session.activate(recs[0]["lob"])
            if callable(session.goto):
                session.goto("combo")
        return _handler

    _open_state = _open_from(sum_tbl)
    sum_tbl.on_click(_open_state)
    nd_tbl.on_click(_open_from(nd_tbl))

    # hidden pages skip the render and catch up via on_show (W3h)
    render_g = gate_hidden(_render)
    rail_rev = debounce(session.bus.param.rev, delay_ms=300, bag=bag)
    bag.watch(rail_rev, render_g, "rev")
    bag.watch(session.ctx, render_g, "data_rev")
    for w in (line_f, bu_f, show_prior, leg_radio, nd_radio):
        bag.watch(w, render_g, "value")
    _render()

    sidebar = pn.Column(
        pn.pane.Markdown("**Filters** — empty means everything"),
        line_f, bu_f,
        pn.layout.Divider(),
        show_prior,
        pn.pane.Markdown(
            "*The prior year is the year currently FLOWING — its "
            "year-ago base is P−2.*"),
        sizing_mode="stretch_width")
    main = pn.Column(
        header, echo, chips_pane, flash,
        pn.Tabs(
            ("Program flow", pn.Column(
                sum_tbl,
                pn.Row(pn.pane.Markdown("**Leg:**"), leg_radio),
                grid_tbl, sizing_mode="stretch_width")),
            ("Net delivery (book)", pn.Column(
                nd_tbl,
                pn.Row(pn.pane.Markdown("**Leg:**"), nd_radio),
                nd_grid,
                pn.pane.Markdown("**Required pricing leg** — single-net-"
                                 "combo states only, CY plan months"),
                req_tbl, sizing_mode="stretch_width")),
            dynamic=False, sizing_mode="stretch_width"),
        note_list(_NOTES),
        sizing_mode="stretch_both")
    render_g.attach(main)
    return {"main": main, "sidebar": sidebar, "bag": bag,
            "sum_tbl": sum_tbl, "grid_tbl": grid_tbl, "nd_tbl": nd_tbl,
            "nd_grid": nd_grid, "req_tbl": req_tbl, "holder": holder,
            "chips": chips_pane, "flash": flash,
            "filters": {"line": line_f, "bu": bu_f},
            "leg": leg_radio, "nd_leg": nd_radio,
            "show_prior": show_prior,
            "open_state": _open_state,
            "on_show": _render}
