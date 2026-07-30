"""Sheet builders: Bridge, Portfolio, Scenarios, Solver, Attribution."""

from __future__ import annotations

import datetime as dt

from openpyxl.chart import BarChart, LineChart, Reference
from openpyxl.chart.marker import Marker
from openpyxl.formatting.rule import ColorScaleRule, DataBarRule
from openpyxl.styles import Alignment, Border, Side

from .build_workbook import Ctx, Layout as L
from .xlstyle import (
    ALIGN_C, ALIGN_WRAP, BORDER_THIN, DOWN_BAR, F_HEADER, F_LABEL, F_SMALL_IT, FAIL_RED,
    FILL_NAVY, FILL_PANEL, FMT_DATE, FMT_EP_Z, FMT_GEN, FMT_IDX, FMT_IDX_Z,
    FMT_INT, FMT_MOD, FMT_PCT, FMT_PCT_Z, GREY_DARK, NAVY, STEEL, STEEL_LIGHT,
    TOTAL_BAR, UP_BAR, col, font, formula, header_row, input_cell, label, link, note,
    presentation_setup, print_setup, put, section, set_widths, title,
)

PTS_Z = '+0.00 "pts";-0.00 "pts";""'
PCT_SIGNED_Z = '+0.0%;-0.0%;""'


def _style_chart(chart, title_text, x_title=None, y_title=None, height=8.5, width=16):
    chart.title = title_text
    chart.style = 2
    chart.height = height
    chart.width = width
    if x_title:
        chart.x_axis.title = x_title
    if y_title:
        chart.y_axis.title = y_title
    # openpyxl omits <c:delete> when the attribute is None, and Excel treats
    # the omission as "axis deleted" on resave — every axis must be explicitly
    # visible unless deliberately hidden (DECISIONS.md D36).
    if chart.x_axis.delete is None:
        chart.x_axis.delete = False
    if chart.y_axis.delete is None:
        chart.y_axis.delete = False
    return chart


def _line_color(series, rgb, width_pt=2.25, dashed=False):
    series.graphicalProperties.line.solidFill = rgb
    series.graphicalProperties.line.width = int(width_pt * 12700)
    if dashed:
        series.graphicalProperties.line.dashStyle = "dash"
    series.marker = Marker(symbol="none")
    series.smooth = False


# ---------------------------------------------------------------------------
# Bridge
# ---------------------------------------------------------------------------


def build_bridge(ctx: Ctx):
    ws = ctx.wb["Bridge"]
    p = ctx.p
    title(ws, "B2", f"Bridge — projected LR to calendar-year plan LR ({ctx.lob.name})",
          "Waterfall from the indication's projected loss ratio to the CY plan loss ratio "
          "for the selected BU x state. All factors trace to the Rate Engine, Mod Engine, "
          "and Inputs sheets.")

    # ---- selected inputs block ----
    section(ws, 5, "B", "Selected combo inputs (tbl_LR via INDEX/MATCH on the BU|LOB key)")
    label(ws, "E5", "row in tbl_LR")
    formula(ws, "F5", "=IF(nr_SelOK,MATCH(nr_SelKey,lr_key,0),0)", fmt=FMT_INT)
    ws["F5"].font = font(GREY_DARK, size=9)
    ctx.define("br_selrow", "Bridge", "$F$5", "tbl_LR row of the selected combo (0 = absent)")

    R = "$F$5"
    rows = [
        ("Business unit", "=nr_SelBU", FMT_GEN, None, None, True),
        ("State", "=nr_SelState", FMT_GEN, None, None, True),
        ("Projected LR (as input)", f"=IF({R}=0,0,N(INDEX(lr_lrproj,{R})))", FMT_PCT,
         "nr_LRproj", "Projected loss ratio from the indication, as entered", False),
        ("LR basis", f'=IF({R}=0,"current",INDEX(lr_basis,{R})&"")', FMT_GEN,
         "nr_Basis", "Basis of the input LR: current | proposed", False),
        ("Indication selected change (s)", f"=IF({R}=0,0,N(INDEX(lr_s,{R})))", FMT_PCT,
         "nr_SelS", "Selected/indicated change used for basis normalization", False),
        ("LR at current rate level",
         '=IF(nr_Basis="proposed",nr_LRproj*(1+nr_SelS),nr_LRproj)', FMT_PCT,
         "nr_LRcur", "Basis-normalized projected LR (§3.4.1)", False),
        ("Mod assumed in indication (M_ind)",
         f"=IF({R}=0,1,IF(N(INDEX(lr_mind,{R}))=0,1,INDEX(lr_mind,{R})))", FMT_MOD,
         "nr_MInd", "Average schedule mod assumed in the indication", False),
        ("Current avg written mod (M_0)",
         f"=IF({R}=0,1,IF(N(INDEX(lr_m0,{R}))=0,1,INDEX(lr_m0,{R})))",
         FMT_MOD, "nr_M0", "Current average written mod", False),
        ("Current mod as-of date (close of day)",
         f"=IF({R}=0,DATE(nr_PlanYear-1,10,1),IF(N(INDEX(lr_m0asof,{R}))=0,"
         f"DATE(nr_PlanYear-1,10,1),INDEX(lr_m0asof,{R})))", FMT_DATE,
         "nr_M0Asof", "As-of date of M_0 (anchored close-of-day, D2)", False),
        ("Projected mod, end of plan yr (M_1)",
         f"=IF({R}=0,1,IF(N(INDEX(lr_m1,{R}))=0,1,INDEX(lr_m1,{R})))",
         FMT_MOD, "nr_M1", "Projected average written mod at 12/31/P", False),
        ("Mod ~1 yr before as-of (M_prior; 0 = not given)",
         f"=IF({R}=0,0,N(INDEX(lr_mprior,{R})))", FMT_MOD,
         "nr_MPrior", "Average mod ~12 months before M_0; 0/blank = extrapolate backward", False),
        ("Projected mod, end plan yr+1 (M_2; 0 = use M_1)",
         f"=IF({R}=0,0,N(INDEX(lr_m2,{R})))", FMT_MOD,
         "nr_M2", "Projected written mod at 12/31/(P+1); 0/blank = M_1", False),
        ("Net trend (P+1 view)",
         f"=IF({R}=0,N(nr_TrendDefault),IF(ISBLANK(INDEX(lr_trend,{R})),N(nr_TrendDefault),"
         f"INDEX(lr_trend,{R})))", FMT_PCT,
         "nr_Trend", "Effective P+1 net trend: tbl_LR value, or the Control default when blank",
         False),
        ("Other adjustment factor (A_other)",
         f"=IF({R}=0,1,IF(N(INDEX(lr_aother,{R}))=0,1,INDEX(lr_aother,{R})))", FMT_IDX,
         "nr_AOther", "Manual adjustment factor (label required when <> 1)", False),
        ("Reason for other adjustment", f'=IF({R}=0,"",INDEX(lr_aotherlbl,{R})&"")', FMT_GEN,
         "nr_AOtherLbl", "Reason for A_other", False),
        ("Mod adjustment (combo row)",
         f'=IF({R}=0,"ON",IF(INDEX(lr_modadj,{R})="OFF","OFF","ON"))', FMT_GEN,
         "nr_ModAdjRow", "Per-combo mod adjustment toggle from tbl_LR", False),
        ("Mod adjustment effective",
         '=IF(AND(nr_ModAdjMaster="ON",nr_ModAdjRow="ON"),"ON","OFF")', FMT_GEN,
         "nr_ModAdjEff", "TRUE state of the mod adjustment (master AND combo row)", False),
        ("Net rate selection active?",
         f"=IF({R}=0,FALSE,NOT(ISBLANK(INDEX(lr_netp,{R}))))", FMT_GEN,
         "nr_NetMode", "TRUE when this combo carries a net rate selection (D39)", False),
        ("Net selection x (P)", f"=IF({R}=0,0,N(INDEX(lr_netp,{R})))", FMT_PCT,
         "nr_NetSelP", "YoY combined net price target for cohorts written from 1/1/P", False),
        ("Net selection x (P+1; blank = carry)",
         f"=IF({R}=0,0,IF(ISBLANK(INDEX(lr_netp1,{R})),N(INDEX(lr_netp,{R})),"
         f"N(INDEX(lr_netp1,{R}))))", FMT_PCT,
         "nr_NetSelP1", "Net selection for P+1 cohorts (defaults to the P selection)", False),
    ]
    r = L.BR_IN_FIRST
    for lbl, f, fmt, name, desc, is_link in rows:
        label(ws, f"B{r}", lbl)
        (link if is_link else formula)(ws, f"C{r}", f, fmt=fmt, border=BORDER_THIN)
        if name:
            ctx.define(name, "Bridge", f"$C${r}", desc)
        r += 1
    put(ws, "D19",
        '=IF(AND(nr_AOther<>1,nr_AOtherLbl=""),"WARNING: A_other <> 1 requires a label","")',
        fnt=font(FAIL_RED, size=9, italic=True))

    # ---- waterfall table ----
    section(ws, L.BR_WF_HDR - 1, "B", "The bridge — plan year and the indicative following year")
    header_row(ws, L.BR_WF_HDR, 2,
               ["Step", "Factor", "LR after", "Step chg", "Factor +1", "LR after +1",
                "Step chg +1"],
               widths=[30, 11, 12, 11, 11, 13, 11])
    # live year labels (D44)
    for cc, f in ((3, '="Factor "&nr_PlanYear'), (4, '="LR after ("&nr_PlanYear&")"'),
                  (5, '="Step ("&nr_PlanYear&")"'), (6, '="Factor "&(nr_PlanYear+1)'),
                  (7, '="LR after ("&(nr_PlanYear+1)&")"'),
                  (8, '="Step ("&(nr_PlanYear+1)&")"')):
        ws.cell(row=L.BR_WF_HDR, column=cc).value = f
    wf = L.BR_WF_FIRST  # 27
    steps = [
        ("Projected LR (as input)", None, "=nr_LRproj", None, "=nr_LRproj"),
        ("Convert to current rate level x(1+s)", '=IF(nr_Basis="proposed",1+nr_SelS,1)', None,
         '=IF(nr_Basis="proposed",1+nr_SelS,1)', None),
        (f"Net trend (applies to {p + 1} only)", "=1", None, "=1+nr_Trend", None),
        ('=IF(nr_NetMode,"Rate + price earn-in (net)  A_net","Rate earn-in  A_rate")',
         "=nr_Arate_P", None, "=nr_Arate_P1", None),
        ('=IF(nr_NetMode,"Mod drift (merged into A_net)","Schedule mod drift  A_mod")',
         "=nr_Amod_P", None, "=nr_Amod_P1", None),
        ("Other  A_other", "=nr_AOther", None, "=nr_AOther", None),
    ]
    for i, (lbl, fac_p, lr_p, fac_p1, lr_p1) in enumerate(steps):
        rr = wf + i
        if lbl.startswith("="):
            formula(ws, f"B{rr}", lbl)
            ws[f"B{rr}"].font = F_LABEL
        else:
            label(ws, f"B{rr}", lbl)
        if fac_p:
            formula(ws, f"C{rr}", fac_p, fmt=FMT_IDX, align=ALIGN_C)
            formula(ws, f"D{rr}", f"=$D{rr - 1}*$C{rr}", fmt=FMT_PCT, align=ALIGN_C)
            formula(ws, f"E{rr}", f"=($D{rr}-$D{rr - 1})*100", fmt=PTS_Z, align=ALIGN_C)
            formula(ws, f"F{rr}", fac_p1, fmt=FMT_IDX, align=ALIGN_C)
            formula(ws, f"G{rr}", f"=$G{rr - 1}*$F{rr}", fmt=FMT_PCT, align=ALIGN_C)
            formula(ws, f"H{rr}", f"=($G{rr}-$G{rr - 1})*100", fmt=PTS_Z, align=ALIGN_C)
        else:
            formula(ws, f"D{rr}", lr_p, fmt=FMT_PCT, align=ALIGN_C)
            formula(ws, f"G{rr}", lr_p1, fmt=FMT_PCT, align=ALIGN_C)
    tot = wf + 6  # 33
    put(ws, f"B{tot}", f"CY plan loss ratio", fnt=font(NAVY, bold=True), fill=FILL_PANEL)
    formula(ws, f"D{tot}", f"=$D{tot - 1}", fmt=FMT_PCT, align=ALIGN_C, bold=True,
            fill=FILL_PANEL)
    formula(ws, f"E{tot}", f"=($D{tot}-$D{wf})*100", fmt=PTS_Z, align=ALIGN_C, fill=FILL_PANEL)
    formula(ws, f"G{tot}", f"=$G{tot - 1}", fmt=FMT_PCT, align=ALIGN_C, bold=True,
            fill=FILL_PANEL)
    formula(ws, f"H{tot}", f"=($G{tot}-$G{wf})*100", fmt=PTS_Z, align=ALIGN_C, fill=FILL_PANEL)
    ctx.define("nr_CYLR_P", "Bridge", f"$D${tot}", "CY plan loss ratio for the plan year")
    ctx.define("nr_CYLR_P1", "Bridge", f"$G${tot}",
               "Indicative CY LR for plan year + 1 (assumes no new indication)")
    note(ws, f"B{tot + 2}",
         "The following-year column is indicative: it holds the indication fixed and applies "
         "the net trend input once (default 0.0% — a visible caveat, not a forecast).")

    # ---- communication metrics ----
    section(ws, tot + 4, "B", "Communication metrics")
    for i, (lbl, f, fmt) in enumerate([
        ("CY earned rate chg vs indication level", "=nr_EChgVsInd", PCT_SIGNED_Z),
        (f"Earned rate chg {p - 1} -> {p}", "=nr_YoY_P", PCT_SIGNED_Z),
        (f"Earned rate chg {p} -> {p + 1} (carryover + new actions)", "=nr_YoY_P1", PCT_SIGNED_Z),
    ]):
        label(ws, f"B{tot + 5 + i}", lbl)
        link(ws, f"D{tot + 5 + i}", f, fmt=fmt, align=ALIGN_C)

    # ---- waterfall chart data + chart ----
    cd = L.BR_CHART_DATA  # 55
    put(ws, f"B{cd - 1}", "Waterfall chart data (formulas — do not edit)", fnt=F_SMALL_IT)
    cats = [("Projected LR", f"$D${wf}", None, True),
            ("Basis", f"$D${wf}", f"$D${wf + 1}", False),
            ("Rate earn-in", f"$D${wf + 2}", f"$D${wf + 3}", False),
            ("Mod drift", f"$D${wf + 3}", f"$D${wf + 4}", False),
            ("Other", f"$D${wf + 4}", f"$D${wf + 5}", False),
            (f"CY {p} plan LR", f"$D${tot}", None, True)]
    for j, h in enumerate(["Step", "base", "up", "down", "total"]):
        put(ws, f"{col(2 + j)}{cd}", h, fnt=font(GREY_DARK, size=9))
    for i, (lbl, a, b, is_total) in enumerate(cats):
        rr = cd + 1 + i
        put(ws, f"B{rr}", lbl, fnt=font(GREY_DARK, size=9))
        if is_total:
            formula(ws, f"C{rr}", "=0", fmt=FMT_PCT)
            formula(ws, f"D{rr}", "=0", fmt=FMT_PCT)
            formula(ws, f"E{rr}", "=0", fmt=FMT_PCT)
            formula(ws, f"F{rr}", f"={a}", fmt=FMT_PCT)
        else:
            formula(ws, f"C{rr}", f"=MIN({a},{b})", fmt=FMT_PCT)
            formula(ws, f"D{rr}", f"=MAX(0,{b}-{a})", fmt=FMT_PCT)
            formula(ws, f"E{rr}", f"=MAX(0,{a}-{b})", fmt=FMT_PCT)
            formula(ws, f"F{rr}", "=0", fmt=FMT_PCT)
        for cL in "CDEF":
            ws[f"{cL}{rr}"].font = font(GREY_DARK, size=9)

    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.gapWidth = 60
    data = Reference(ws, min_col=3, min_row=cd, max_col=6, max_row=cd + 6)
    cats_ref = Reference(ws, min_col=2, min_row=cd + 1, max_row=cd + 6)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.series[0].graphicalProperties.noFill = True
    chart.series[1].graphicalProperties.solidFill = UP_BAR
    chart.series[2].graphicalProperties.solidFill = DOWN_BAR
    chart.series[3].graphicalProperties.solidFill = TOTAL_BAR
    chart.legend = None
    chart.y_axis.number_format = "0%"
    _style_chart(chart, "Projected LR -> CY plan loss ratio (stacked-column waterfall)",
                 y_title="Loss ratio", height=9, width=15)
    ws.add_chart(chart, "J5")

    set_widths(ws, {"A": 2, "B": 34, "C": 13, "D": 12, "E": 11, "F": 12, "G": 13, "H": 11})
    presentation_setup(ws, gridlines_off=True, tab_color=NAVY)
    print_setup(ws)


# ---------------------------------------------------------------------------
# Portfolio
# ---------------------------------------------------------------------------


def build_portfolio(ctx: Ctx):
    ws = ctx.wb["Portfolio"]
    p = ctx.p
    title(ws, "A1", f"Portfolio — every BU x state ({ctx.lob.name})",
          "Computed simultaneously from the hidden _calc engine blocks (rows align 1:1 with tbl_LR). "
          "Heatmap: green = favorable vs projected.")
    header_row(ws, L.PF_HDR, 1,
               ["BU", "State", "Key", "Projected LR (as input)", "Basis",
                "Projected LR (current level)", "Plan LR",
                "Chg vs projected (pts)", "Rate earn-in (A_rate)", "Mod drift (A_mod)",
                "Earned rate chg vs indication", "Carryover", "Plan LR +1",
                "Adj plan EP (000s)", "EP weight"],
               widths=[9, 8, 4, 11, 10, 11, 11, 11, 10, 10, 13, 12, 11, 12, 8])
    ws.row_dimensions[L.PF_HDR].height = 42
    # live year labels (D44)
    ws.cell(row=L.PF_HDR, column=7).value = '="CY "&nr_PlanYear&" plan LR"'
    ws.cell(row=L.PF_HDR, column=12).value = '="Carryover into "&(nr_PlanYear+1)'
    ws.cell(row=L.PF_HDR, column=13).value = '="CY "&(nr_PlanYear+1)&" plan LR"'
    for i in range(L.LR_ROWS):
        n = i + 1
        r = L.PF_FIRST + i
        link(ws, f"A{r}", f'=IF(INDEX(lr_key,{n})="","",INDEX(lr_bu,{n}))')
        link(ws, f"B{r}", f'=IF(INDEX(lr_key,{n})="","",INDEX(lr_state,{n}))')
        link(ws, f"C{r}", f"=INDEX(lr_key,{n})")
        ws[f"C{r}"].font = font(GREY_DARK, size=8)
        link(ws, f"D{r}", f"=N(INDEX(lr_lrproj,{n}))", fmt=FMT_PCT_Z, align=ALIGN_C)
        link(ws, f"E{r}", f'=IF($C{r}="","",INDEX(lr_basis,{n})&"")', align=ALIGN_C)
        link(ws, f"F{r}", f"=INDEX(calc_lrcur,{n})", fmt=FMT_PCT_Z, align=ALIGN_C)
        link(ws, f"G{r}", f"=INDEX(calc_cylr_p,{n})", fmt=FMT_PCT_Z, align=ALIGN_C, bold=True)
        link(ws, f"H{r}", f"=(INDEX(calc_cylr_p,{n})-INDEX(calc_lrcur,{n}))*100",
             fmt=PTS_Z, align=ALIGN_C)
        link(ws, f"I{r}", f'=IF($C{r}="",0,INDEX(calc_arate_p,{n}))', fmt=FMT_IDX_Z,
             align=ALIGN_C)
        link(ws, f"J{r}", f'=IF($C{r}="",0,INDEX(calc_amod_p,{n}))', fmt=FMT_IDX_Z,
             align=ALIGN_C)
        link(ws, f"K{r}", f'=IF($C{r}="",0,INDEX(calc_echg,{n}))', fmt=PCT_SIGNED_Z,
             align=ALIGN_C)
        link(ws, f"L{r}", f'=IF($C{r}="",0,INDEX(calc_carry,{n}))', fmt=PCT_SIGNED_Z,
             align=ALIGN_C)
        link(ws, f"M{r}", f"=INDEX(calc_cylr_p1,{n})", fmt=FMT_PCT_Z, align=ALIGN_C)
        link(ws, f"N{r}", f"=N(INDEX(lr_ep,{n}))", fmt=FMT_EP_Z, align=ALIGN_C)
        formula(ws, f"O{r}", f"=IF(SUM($N${L.PF_FIRST}:$N${L.PF_LAST})=0,0,"
                             f"$N{r}/SUM($N${L.PF_FIRST}:$N${L.PF_LAST}))",
                fmt=FMT_PCT_Z, align=ALIGN_C)
        ws[f"O{r}"].font = font(GREY_DARK, size=9)

    f0, f1 = L.PF_FIRST, L.PF_LAST
    wr, sr = L.PF_W_TOTAL, L.PF_S_TOTAL
    put(ws, f"B{wr}", "EP-weighted total *", fnt=font(NAVY, bold=True), fill=FILL_PANEL)
    put(ws, f"B{sr}", "Simple average", fnt=font(GREY_DARK, bold=True))
    epsum = f"SUM($N${f0}:$N${f1})"
    cnt = f'SUMPRODUCT(($C${f0}:$C${f1}<>"")*1)'
    for cL, fmt in (("F", FMT_PCT), ("G", FMT_PCT), ("K", PCT_SIGNED_Z), ("L", PCT_SIGNED_Z),
                    ("M", FMT_PCT)):
        formula(ws, f"{cL}{wr}",
                f'=IF({epsum}=0,"n/a",SUMPRODUCT({cL}${f0}:{cL}${f1},$N${f0}:$N${f1})/{epsum})',
                fmt=fmt, align=ALIGN_C, bold=(cL == "G"), fill=FILL_PANEL)
        formula(ws, f"{cL}{sr}",
                f'=IF({cnt}=0,"n/a",SUM({cL}${f0}:{cL}${f1})/{cnt})',
                fmt=fmt, align=ALIGN_C)
    formula(ws, f"H{wr}", f'=IF(ISNUMBER($G${wr}),($G${wr}-$F${wr})*100,"")', fmt=PTS_Z,
            align=ALIGN_C, fill=FILL_PANEL)
    formula(ws, f"N{wr}", f"={epsum}", fmt=FMT_EP_Z, align=ALIGN_C, fill=FILL_PANEL)
    note(ws, f"B{sr + 2}",
         "* Weighted by plan earned premium over the rows where EP is provided; rows without "
         "EP carry zero weight. The simple average covers every populated row equally.")

    ws.conditional_formatting.add(
        f"G{f0}:G{f1}",
        ColorScaleRule(start_type="min", start_color="D6E8D5",
                       mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                       end_type="max", end_color="F4B8B8"))
    ws.conditional_formatting.add(
        f"H{f0}:H{f1}",
        ColorScaleRule(start_type="min", start_color="D6E8D5",
                       mid_type="num", mid_value=0, mid_color="FFFFFF",
                       end_type="max", end_color="F4B8B8"))
    for cL in ("I", "J"):
        ws.conditional_formatting.add(
            f"{cL}{f0}:{cL}{f1}",
            ColorScaleRule(start_type="num", start_value=0.95, start_color="D6E8D5",
                           mid_type="num", mid_value=1.0, mid_color="FFFFFF",
                           end_type="num", end_value=1.05, end_color="F4B8B8"))
    presentation_setup(ws, gridlines_off=True, freeze=f"D{L.PF_FIRST}", tab_color=NAVY)
    print_setup(ws)


# ---------------------------------------------------------------------------
# State Summary — the flagship leadership exhibit (D38)
# ---------------------------------------------------------------------------

SS_CAPTION_FILL = "EDF1F5"


def build_state_summary(ctx: Ctx):
    """One row per state: adjusted EP, mods, chronological rate-change history,
    engine results, and the CY P / P+1 plan LR — filtered by BU, with an
    'All' view that combines business units on adjusted-EP weights."""
    ws = ctx.wb["State Summary"]
    p = ctx.p
    states = ctx.cfg.states
    title(ws, "A1", f"State Summary — {ctx.lob.name}",
          "Plan-year view with the indicative following year. One row per state; choose a "
          "business unit or view all combined (adjusted-EP weighted). Every figure traces to "
          "the hidden _calc engine blocks and the Inputs tables.")

    # ---- filter chip ----
    label(ws, "A4", "Business unit view", bold=True)
    c = input_cell(ws, "B4", "All")
    c.alignment = ALIGN_C
    c.border = Border(*(Side(style="medium", color=NAVY),) * 4)
    formula(ws, "C4",
            '="Showing: "&IF($B$4="All","all business units combined (adj-EP weighted)",$B$4)')
    ws["C4"].font = font(GREY_DARK, italic=True)
    ctx.define("nr_SumBU", "State Summary", "$B$4",
               "State Summary business-unit filter (All = EP-weighted across BUs)")
    # criteria helper: SUMIFS wildcard when 'All'
    formula(ws, "AA4", '=IF($B$4="All","*",$B$4)')
    ws["AA4"].font = font(GREY_DARK, size=8)
    from openpyxl.worksheet.datavalidation import DataValidation
    dv = DataValidation(type="list", formula1="=lst_bu_all", allow_blank=False,
                        showErrorMessage=True)
    ws.add_data_validation(dv)
    dv.add("B4")

    cap_row, hdr, first = 6, 7, 8
    n_disp = len(states) + 3   # live-roster slots incl. headroom for new states (D42)
    last = first + n_disp - 1
    tot = last + 2
    crit = "$AA$4"

    # ---- two-tier header ----
    groups = [
        (1, 1, ""), (2, 2, "Volume"), (3, 5, "Schedule mods"),
        (6, 14, "Rate change history — chronological (single-BU view)"),
        (15, 19, "Engine results"), (20, 21, '="CY "&nr_PlanYear&" plan"'),
        (22, 25, '="CY "&(nr_PlanYear+1)&" indicative"'), (26, 26, "Net"),
    ]
    for c1, c2, text in groups:
        for cc in range(c1, c2 + 1):
            cell = ws.cell(row=cap_row, column=cc, value=text if cc == c1 else None)
            cell.alignment = Alignment(horizontal="centerContinuous", vertical="center")
            cell.font = font(NAVY, bold=True, size=9)
            cell.fill = PatternFill_safe(SS_CAPTION_FILL)
    headers = ["State", "Adj plan EP (000s)", "Mod in indication", "Current mod",
               "Proj. mod, plan-yr end",
               "Chg 1 date", "Chg 1", "Chg 2 date", "Chg 2", "Chg 3 date", "Chg 3",
               "Chg 4 date", "Chg 4", "# chgs", "Indication rate level",
               "Earned rate level", "Rate earn-in (A_rate)", "Earned mod",
               "Mod drift (A_mod)", "Projected LR (current level)", "Plan LR",
               "Net trend", "Rate earn-in +1", "Mod drift +1", "Plan LR +1", "Net rate sel"]
    widths = [7, 12, 9, 8.5, 9, 9, 7, 9, 7, 9, 7, 9, 7, 6.5, 9.5, 9.5, 9.5, 9, 9.5, 11, 10,
              8.5, 9.5, 9.5, 10, 8.5]
    header_row(ws, hdr, 1, headers, widths=widths)
    ws.row_dimensions[hdr].height = 42
    # year-bearing headers are LIVE so a Control plan-year change relabels
    # the exhibit, not just the math (D44)
    for cc, f in ((16, '=nr_PlanYear&" earned rate level"'),
                  (18, '=nr_PlanYear&" earned mod"'),
                  (21, '="CY "&nr_PlanYear&" plan LR"'),
                  (22, '="Net trend ("&(nr_PlanYear+1)&")"'),
                  (23, '=(nr_PlanYear+1)&" rate earn-in"'),
                  (24, '=(nr_PlanYear+1)&" mod drift"'),
                  (25, '="CY "&(nr_PlanYear+1)&" plan LR"')):
        ws.cell(row=hdr, column=cc).value = f
    group_starts = {3, 6, 15, 20, 22, 26}
    med = Side(style="medium", color=STEEL)

    def wtd(row_, prod, plain, plain_dims=("calc_state", "calc_bu")):
        ps, pb = plain_dims
        return (f'=IF($AA{row_}=0,"",IF($B{row_}>0,'
                f"SUMIFS({prod},calc_state,$A{row_},calc_bu,{crit})/$B{row_},"
                f"SUMIFS({plain},{ps},$A{row_},{pb},{crit})/$AA{row_}))")

    metric_cols = [
        (3, "calc_w_mind", "lr_mind", ("lr_state", "lr_bu"), FMT_MOD),
        (4, "calc_w_m0", "lr_m0", ("lr_state", "lr_bu"), FMT_MOD),
        (5, "calc_w_m1", "lr_m1", ("lr_state", "lr_bu"), FMT_MOD),
        (15, "calc_w_crl", "calc_crl", None, FMT_IDX),
        (16, "calc_w_ecy", "calc_ecy_p", None, FMT_IDX),
        (17, "calc_w_arate", "calc_arate_p", None, FMT_IDX),
        (18, "calc_w_mbar", "calc_mbar_p", None, FMT_MOD),
        (19, "calc_w_amod", "calc_amod_p", None, FMT_IDX),
        (20, "calc_w_lrcur", "calc_lrcur", None, FMT_PCT),
        (21, "calc_w_cylr", "calc_cylr_p", None, FMT_PCT),
        (22, "calc_w_trend", "calc_trend", None, "+0.0%;-0.0%;0.0%"),
        (23, "calc_w_arate1", "calc_arate_p1", None, FMT_IDX),
        (24, "calc_w_amod1", "calc_amod_p1", None, FMT_IDX),
        (25, "calc_w_cylr1", "calc_cylr_p1", None, FMT_PCT),
    ]

    for i in range(n_disp):
        r = first + i
        band = FILL_PANEL if i % 2 else None
        # live roster: the k-th distinct state currently in tbl_LR (D42)
        link(ws, f"A{r}", f"=_lists!$B${3 + i}", align=ALIGN_C, fill=band,
             border=BORDER_THIN, bold=True)
        formula(ws, f"AA{r}", f"=COUNTIFS(calc_state,$A{r},calc_bu,{crit})", fmt=FMT_INT)
        ws[f"AA{r}"].font = font(GREY_DARK, size=8)
        formula(ws, f"AB{r}", f'=IF($B$4="All","",$B$4&"|"&$A{r})')
        ws[f"AB{r}"].font = font(GREY_DARK, size=8)
        formula(ws, f"B{r}",
                f'=IF($AA{r}=0,"",SUMIFS(calc_ep,calc_state,$A{r},calc_bu,{crit}))',
                fmt="#,##0", align=ALIGN_C, fill=band, border=BORDER_THIN)
        for cc, prod, plain, dims, fmt in metric_cols:
            f = wtd(r, prod, plain, dims) if dims else wtd(r, prod, plain)
            formula(ws, ws.cell(row=r, column=cc).coordinate, f, fmt=fmt, align=ALIGN_C,
                    fill=band, border=BORDER_THIN, bold=(cc in (21, 25)))
        for j in range(4):
            cd, cp = 6 + j * 2, 7 + j * 2
            formula(ws, ws.cell(row=r, column=cd).coordinate,
                    f'=IF($AB{r}="","",IF(COUNTIFS(rl_key,$AB{r},rl_seq,{j + 1})=0,"",'
                    f"SUMIFS(rl_eff,rl_key,$AB{r},rl_seq,{j + 1})))",
                    fmt="m/d/yy", align=ALIGN_C, fill=band, border=BORDER_THIN)
            formula(ws, ws.cell(row=r, column=cp).coordinate,
                    f'=IF($AB{r}="","",IF(COUNTIFS(rl_key,$AB{r},rl_seq,{j + 1})=0,"",'
                    f"SUMIFS(rl_reff,rl_key,$AB{r},rl_seq,{j + 1})))",
                    fmt="+0.0%;-0.0%;0.0%", align=ALIGN_C, fill=band, border=BORDER_THIN)
        formula(ws, f"N{r}",
                f'=IF($AA{r}=0,"",IF($AB{r}="",COUNTIFS(rl_state,$A{r},rl_eff,"<>"),'
                f'COUNTIFS(rl_key,$AB{r},rl_eff,"<>")))',
                fmt=FMT_INT, align=ALIGN_C, fill=band, border=BORDER_THIN)
        formula(ws, f"AC{r}", f"=SUMIFS(calc_netmode,calc_state,$A{r},calc_bu,{crit})",
                fmt=FMT_INT)
        ws[f"AC{r}"].font = font(GREY_DARK, size=8)
        formula(ws, f"Z{r}",
                f'=IF($AA{r}=0,"",IF($AC{r}=0,"—",'
                f"SUMIFS(calc_netx,calc_state,$A{r},calc_bu,{crit})/$AC{r}))",
                fmt="+0.0%;-0.0%;0.0%", align=ALIGN_C, fill=band, border=BORDER_THIN)

    # ---- total band ----
    put(ws, f"A{tot}", "TOTAL", fnt=font(NAVY, bold=True), align=ALIGN_C,
        fill=FILL_STEEL_LIGHT_safe(), border=Border(top=med))
    formula(ws, f"B{tot}", f"=SUMIFS(calc_ep,calc_bu,{crit})", fmt="#,##0", align=ALIGN_C,
            bold=True, fill=FILL_STEEL_LIGHT_safe(), border=Border(top=med))
    for cc, prod, plain, dims, fmt in metric_cols:
        formula(ws, ws.cell(row=tot, column=cc).coordinate,
                f'=IF($B${tot}=0,"n/a",SUMIFS({prod},calc_bu,{crit})/$B${tot})',
                fmt=fmt, align=ALIGN_C, bold=(cc in (21, 25)),
                fill=FILL_STEEL_LIGHT_safe(), border=Border(top=med))
    formula(ws, f"N{tot}",
            f'=IF($B$4="All",COUNTIFS(rl_eff,"<>"),COUNTIFS(rl_bu,$B$4,rl_eff,"<>"))',
            fmt=FMT_INT, align=ALIGN_C, fill=FILL_STEEL_LIGHT_safe(), border=Border(top=med))
    formula(ws, f"Z{tot}",
            f'=IF(SUMIFS(calc_netmode,calc_bu,{crit})=0,"—",'
            f"SUMIFS(calc_netx,calc_bu,{crit})/SUMIFS(calc_netmode,calc_bu,{crit}))",
            fmt="+0.0%;-0.0%;0.0%", align=ALIGN_C, fill=FILL_STEEL_LIGHT_safe(),
            border=Border(top=med))
    for cc in list(range(6, 14)):
        put(ws, ws.cell(row=tot, column=cc).coordinate, None,
            fill=FILL_STEEL_LIGHT_safe(), border=Border(top=med))

    # group divider borders down the table
    for cc in group_starts:
        for r in range(cap_row, tot + 1):
            cell = ws.cell(row=r, column=cc)
            b = cell.border
            cell.border = Border(left=med, right=b.right, top=b.top, bottom=b.bottom)

    # conditional formatting: LR heatmaps + EP data bars
    for colL in ("U", "Y"):
        ws.conditional_formatting.add(
            f"{colL}{first}:{colL}{last}",
            ColorScaleRule(start_type="min", start_color="D6E8D5",
                           mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                           end_type="max", end_color="F4B8B8"))
    ws.conditional_formatting.add(
        f"B{first}:B{last}",
        DataBarRule(start_type="num", start_value=0, end_type="max",
                    color="8497B0", showValue=True))

    notes = [
        "Weighted view: every metric is adjusted-plan-EP weighted across the business units in "
        "view; states whose combos carry no EP fall back to a simple average of their combos.",
        "Rate change history lists effective (achievement-adjusted) changes in chronological "
        "order and is shown for single-BU views (business units run different rate programs); "
        "the # chgs column counts changes in every view. More than four changes: see tbl_RateLog.",
        "The following year is indicative: it holds the indication fixed and applies the net "
        "trend once (per-state tbl_LR entries override the Control default).",
        "Net sel: the net rate selection in force (average across the combos in view that "
        "carry one; — = explicit rate program). Net-mode combos show the combined rate x mod "
        "factor in the A_rate column with A_mod = 1 (DECISIONS.md D39).",
    ]
    for i, text in enumerate(notes):
        note(ws, f"A{tot + 2 + i}", text)

    # ------------------------------------------------------------------
    # Documentation panel (D41): this is most readers' first stop, so the
    # calculations are explained in place — plain language for actuaries
    # and product managers, with a LIVE worked example driven by the
    # Control selection (it stays correct after sample data is replaced).
    # ------------------------------------------------------------------
    d = tot + 2 + len(notes) + 2

    def _doc_head(row_, text):
        section(ws, row_, "A", text)
        return row_ + 1

    def _doc_line(row_, term, definition):
        put(ws, f"A{row_}", term, fnt=font(NAVY, bold=True, size=10))
        put(ws, f"E{row_}", definition, fnt=font(GREY_DARK, size=10))
        return row_ + 1

    d = _doc_head(d, "How this exhibit is calculated")
    put(ws, f"A{d}",
        "Every figure below the headers is calculated, never typed. Each state row is the "
        "adjusted-EP-weighted combination of its BU x state calculations for the business "
        "units in view; pick any single combo on the Control sheet to walk the full audit "
        "trail (Rate Engine -> Mod Engine -> Bridge).", fnt=font(GREY_DARK, size=10))
    d += 2

    put(ws, f"A{d}", "THE BRIDGE IN ONE LINE      CY plan LR  =  Projected LR (current "
                     "level)  x  rate earn-in  x  mod drift  x  other adj",
        fnt=font(NAVY, bold=True, size=11), fill=FILL_PANEL)
    for cc in range(2, 26):
        put(ws, ws.cell(row=d, column=cc).coordinate, None, fill=FILL_PANEL)
    d += 1
    put(ws, f"A{d}", "The following year (indicative) additionally applies one year of net "
                     "trend and uses that year's earn-in factors.", fnt=F_SMALL_IT)
    d += 2

    d = _doc_head(d, "Column guide")
    guide = [
        ("Adj plan EP", "Adjusted plan earned premium (000s) — the weight used to combine "
                        "business units and to build the TOTAL row."),
        ("Mod columns  (M_ind / M_0 / M_1)",
         "The average schedule mod assumed in the indication; the current written average; "
         "the projected average at plan-year end. Drift in the average mod is a price change "
         "that earns in exactly like rate."),
        ("Rate change history", "Effective changes (filed % x achievement for planned rows) in "
                                "date order for the business unit in view."),
        ("Indication rate level  (CRL_ind)",
         "The cumulative rate level the indication assumed FULLY EARNED — the product of all "
         "rate changes it considered. This is the 'assumed rate in the indication'."),
        ("Earned rate level  (E_CY)",
         "The rate level the plan calendar year ACTUALLY earns. Policies written last year "
         "carry older rates into its early months, and new actions earn in only gradually "
         "as the book turns over — so the earned level rarely matches the indication's."),
        ("Rate earn-in  (A_rate)",
         "Indication rate level / earned rate level. Above 1.000: the plan year earns less "
         "rate than the indication assumed, pushing the plan LR up. Below 1.000: extra rate "
         "the indication never counted earns in, pushing it down."),
        ("Earned mod  /  Mod drift (A_mod)",
         "The average schedule mod the plan year earns, and M_ind / that value — the "
         "identical earn-in logic applied to the price (mod) lever."),
        ("Projected LR (current level)",
         "The projected LR restated to current rate level (multiplied by (1+s) when the "
         "indication quoted it at the proposed level)."),
        ("Plan LR",
         "Projected LR (current level)  x  rate earn-in  x  mod drift  x  other adj."),
        ("Net trend  /  next-year view",
         "The following-year view holds the indication fixed, applies the net "
         "loss-over-premium trend once, and swaps in that year's earn-in factors. "
         "Indicative only — it assumes no new indication."),
        ("Net rate sel",
         "States planned on a NET RATE SELECTION: a combined rate + price year-over-year "
         "target from 1/1 of the plan year replaces the planned program for business "
         "written from that date. The A_rate column then carries the combined factor "
         "and A_mod shows 1.000."),
    ]
    for term, definition in guide:
        d = _doc_line(d, term, definition)
    d += 1

    section(ws, d, "A", "Worked example — live from the Control selection")
    link(ws, f"E{d}", '="Currently showing:  "&nr_SelKey', bold=True)
    d += 1
    example = [
        ("1.  Projected loss ratio (indication)", "=nr_LRproj", FMT_PCT,
         '="the indication\'s loss pick — already centered on CY "&nr_PlanYear&'
         '" losses, so no additional loss trend applies"'),
        ("2.  x  basis normalization",
         '=IF(nr_Basis="proposed",1+nr_SelS,1)', FMT_IDX,
         '=IF(nr_Basis="proposed","input LR was at proposed level: x (1+s)",'
         '"input LR already at current level: x 1.000")'),
        ("3.  x  rate earn-in  A_rate", "=nr_Arate_P", FMT_IDX,
         '=IF(nr_NetMode,"combined rate + price factor A_net = CRL_ind "&TEXT(nr_CRLind,'
         '"0.0000")&" / net earned "&TEXT(nr_ECY_P,"0.0000"),"CRL_ind "&TEXT(nr_CRLind,'
         '"0.0000")&" / earned "&TEXT(nr_ECY_P,"0.0000")&" — carryover plus partly-earned '
         'actions")'),
        ("4.  x  schedule-mod drift  A_mod", "=nr_Amod_P", FMT_IDX,
         '=IF(nr_NetMode,"1.000 — merged into the net factor above",'
         '"M_ind "&TEXT(nr_MInd,"0.000")&" / earned mod "&TEXT(nr_MEarned_P,"0.000"))'),
        ("5.  x  other adjustment  A_other", "=nr_AOther", FMT_IDX,
         '=IF(nr_AOtherLbl="","no manual adjustment",nr_AOtherLbl)'),
        ("RESULT:  CY plan loss ratio", "=nr_CYLR_P", FMT_PCT,
         '="the value in this state\'s CY "&nr_PlanYear&" plan LR column when its BU is in view"'),
        ("Next-year plan LR (indicative)", "=nr_CYLR_P1", FMT_PCT,
         '="x (1 + net trend "&TEXT(nr_Trend,"0.0%")&") with CY "&(nr_PlanYear+1)&'
         '"\'s earn-in factors"'),
    ]
    for i, (lbl, val_f, fmt, expl_f) in enumerate(example):
        r = d + i
        bold_ = lbl.startswith("RESULT")
        put(ws, f"A{r}", lbl, fnt=font(NAVY, bold=bold_, size=10),
            fill=FILL_PANEL if bold_ else None)
        link(ws, f"E{r}", val_f, fmt=fmt, align=ALIGN_C, bold=bold_,
             border=BORDER_THIN)
        link(ws, f"G{r}", expl_f)
        ws[f"G{r}"].font = font(GREY_DARK, size=10, italic=True)
    d += len(example) + 1
    put(ws, f"A{d}",
        "Full derivations, conventions, and every named range: see the Methodology sheet. "
        "The engines behind every figure are visible on Rate Engine and Mod Engine.",
        fnt=F_SMALL_IT)

    ws.column_dimensions["AA"].width = 6
    ws.column_dimensions["AB"].width = 10
    ws.column_dimensions["AC"].width = 6
    presentation_setup(ws, gridlines_off=True, freeze=f"B{first}", tab_color=NAVY)
    print_setup(ws)


def PatternFill_safe(rgb):
    from openpyxl.styles import PatternFill
    return PatternFill("solid", fgColor=rgb)


def FILL_STEEL_LIGHT_safe():
    from openpyxl.styles import PatternFill
    return PatternFill("solid", fgColor=STEEL_LIGHT)


# ---------------------------------------------------------------------------
# Scenarios
# ---------------------------------------------------------------------------


def build_scenarios(ctx: Ctx):
    ws = ctx.wb["Scenarios"]
    p = ctx.p
    title(ws, "B2", "Scenarios — selected BU x LOB",
          "Up to four what-if variations of the planned rate program and mod path. "
          "Blank lever = Base value. Engines live on the hidden _calc sheet.")
    label(ws, "B3", "Selected:")
    link(ws, "C3", "=nr_SelKey", bold=True)

    section(ws, 5, "B", "Scenario levers (blue = enter; blank = same as Base)")
    header_row(ws, 5, 3, ["S1", "S2", "S3", "S4"], widths=[10, 10, 10, 10])
    levers = [
        ("Extra filed rate on planned rows (+/- pts)", FMT_PCT, "sc_dpts",
         "Scenario lever: added to every planned row's filed %"),
        ("Shift planned effective dates (+/- months)", FMT_INT, "sc_shift",
         "Scenario lever: EDATE month shift applied to planned rows"),
        ("Achievement override (planned rows)", FMT_PCT, "sc_ach",
         "Scenario lever: replaces achievement % on planned rows"),
        ("Shift projected mod path M_1/M_2 (+/-)", FMT_MOD, "sc_dm1",
         "Scenario lever: added to M_1 and M_2 anchors"),
        ("Adjust net rate selection (+/- pts)", FMT_PCT, "sc_dnet",
         "Scenario lever: added to the net selection x for P and P+1 (D39)"),
    ]
    for i, (lbl, fmt, name, desc) in enumerate(levers):
        r = 6 + i
        label(ws, f"B{r}", lbl)
        for cc in range(3, 7):
            input_cell(ws, f"{col(cc)}{r}", None, fmt=fmt, required=False)
        ctx.define(name, "Scenarios", f"$C${r}:$F${r}", desc)

    section(ws, 12, "B", "Results")
    put(ws, "B13", "Metric", fnt=F_HEADER, fill=FILL_NAVY)
    for j, h in enumerate(["Base", "S1", "S2", "S3", "S4"]):
        put(ws, f"{col(3 + j)}13", h, fnt=F_HEADER, fill=FILL_NAVY, align=ALIGN_C)
    ctx.define("sc_res_names", "Scenarios", "$C$13:$G$13", "Scenario result column headers")
    ctx.define("sc_res_cylr", "Scenarios", "$C$14:$G$14", "CY plan LR by scenario")

    metrics = ["Plan LR", "Change vs Base (pts)", "Rate earn-in (A_rate)",
               "Mod drift (A_mod)", "Earned rate level (E_CY)", "Carryover"]
    for i, m in enumerate(metrics):
        label(ws, f"B{14 + i}", m)
    # live year labels (D44)
    formula(ws, "B14", '="CY "&nr_PlanYear&" plan LR"')
    ws["B14"].font = F_LABEL
    formula(ws, "B19", '="Carryover into "&(nr_PlanYear+1)')
    ws["B19"].font = F_LABEL
    # Base column (green links to the live engines)
    for r, f, fmt in ((14, "=nr_CYLR_P", FMT_PCT), (15, "=($C$14-$C$14)*100", PTS_Z),
                      (16, "=nr_Arate_P", FMT_IDX), (17, "=nr_Amod_P", FMT_IDX),
                      (18, "=nr_ECY_P", FMT_IDX), (19, "=nr_YoY_P1", PCT_SIGNED_Z)):
        link(ws, f"C{r}", f, fmt=fmt, align=ALIGN_C, bold=(r == 14))
    for s in range(1, 5):
        cL = col(3 + s)
        res = ctx.lay_dyn["scenario_results"][s - 1]
        link(ws, f"{cL}16", f"=nr_CRLind/{res['e_p']}", fmt=FMT_IDX, align=ALIGN_C)
        link(ws, f"{cL}17",
             f'=IF(nr_NetMode,1,IF(AND(nr_ModAdjMaster="ON",nr_ModAdjRow="ON"),'
             f"nr_MInd/{res['mbar_p']},1))",
             fmt=FMT_IDX, align=ALIGN_C)
        formula(ws, f"{cL}14", f"=nr_LRcur*{cL}16*{cL}17*nr_AOther", fmt=FMT_PCT,
                align=ALIGN_C, bold=True)
        formula(ws, f"{cL}15", f"=({cL}14-$C$14)*100", fmt=PTS_Z, align=ALIGN_C)
        link(ws, f"{cL}18", f"={res['e_p']}", fmt=FMT_IDX, align=ALIGN_C)
        link(ws, f"{cL}19", f"={res['e_p1']}/{res['e_p']}-1", fmt=PCT_SIGNED_Z, align=ALIGN_C)
    note(ws, "B21", "Scenario CY LR = LR_current x A_rate' x A_mod' x A_other. CRL_ind is "
                    "unchanged (scenarios vary planned actions only, never the indication). "
                    "When the selected combo carries a net rate selection, only the D-net "
                    "lever moves results (planned-row levers are superseded, D39).")

    chart = BarChart()
    chart.type = "col"
    chart.gapWidth = 60
    data = Reference(ws, min_col=3, min_row=14, max_col=7, max_row=14)
    cats = Reference(ws, min_col=3, min_row=13, max_col=7, max_row=13)
    chart.add_data(data, from_rows=True, titles_from_data=False)
    chart.set_categories(cats)
    chart.series[0].graphicalProperties.solidFill = STEEL
    chart.legend = None
    chart.y_axis.number_format = "0.0%"
    _style_chart(chart, "Plan-year loss ratio by scenario", y_title="CY plan LR",
                 height=8, width=13)
    ws.add_chart(chart, "B24")

    set_widths(ws, {"A": 2, "B": 40, "C": 11, "D": 11, "E": 11, "F": 11, "G": 11})
    presentation_setup(ws, gridlines_off=True, tab_color=STEEL_LIGHT)
    print_setup(ws)


# ---------------------------------------------------------------------------
# Solver (E1)
# ---------------------------------------------------------------------------


def build_solver(ctx: Ctx):
    ws = ctx.wb["Solver"]
    p = ctx.p
    title(ws, "B2", "Solver — inverse plan (closed form, no iteration)",
          "Because the CY earned index is linear in a single unknown change r, the required "
          "rate solves in closed form.")
    put(ws, "B4",
        "Convention (DECISIONS.md D13): the Solver replaces the ENTIRE planned-rate program "
        "with a single action at the chosen date. Its base index W0 uses taken rows only.",
        fnt=font(FAIL_RED, size=9, italic=True), align=ALIGN_WRAP)
    formula(ws, "B5",
            '=IF(nr_NetMode,"NOTE: the selected combo carries a NET RATE SELECTION - Solver '
            'results describe the explicit-program counterfactual (planned rows + classic '
            'A_mod), not the net path (D39).","")')
    ws["B5"].font = font(FAIL_RED, size=9, italic=True)

    # ---- Mode A ----
    section(ws, 6, "B", "Mode A — required rate for a target CY LR")
    formula(ws, "B7", '="Target CY "&nr_PlanYear&" loss ratio"')
    ws["B7"].font = F_LABEL
    input_cell(ws, "C7", ctx.oracle_m.cy_lr_p, fmt=FMT_PCT)
    label(ws, "B8", "Effective date")
    input_cell(ws, "C8", dt.date(p, 4, 1), fmt=FMT_DATE)
    label(ws, "B9", "Achievement % assumed")
    input_cell(ws, "C9", 1.0, fmt=FMT_PCT)
    note(ws, "D7", "seeded with the sample book's own CY plan LR - the solver should return "
                   "the seeded planned change (+5.0% at 4/1) while sample data is intact")

    helpers = [
        ("abs eff month", "=YEAR($C$8)*12+MONTH($C$8)-1", FMT_INT),
        ("month start", "=DATE(YEAR($C$8),MONTH($C$8),1)", FMT_DATE),
        ("month end", "=EOMONTH($C$8,0)", FMT_DATE),
        ("days in month", "=$G$9-$G$8+1", FMT_INT),
        ("days on/after eff", "=$G$9-$C$8+1", FMT_INT),
        ("eff month in cohort window?",
         "=AND($G$7>=(nr_PlanYear-2)*12,$G$7<=(nr_PlanYear+1)*12+11)", FMT_GEN),
        ("cohort row index (clamped)", "=MIN(48,MAX(1,$G$7-(nr_PlanYear-2)*12+1))", FMT_INT),
        ("first taken-change days (same month)", "=IF($G$12,INDEX(slv_dfirst,$G$13),0)", FMT_INT),
        ("split p (first-change rule, D4/D31)",
         "=IF($G$12,MIN(1,MAX($G$11,$G$14)/$G$10),1)", "0.000"),
        ("W0_pre at eff month", "=IF($G$12,INDEX(slv_wpre,$G$13),1)", FMT_IDX),
        ("W0_eom at eff month", "=IF($G$12,INDEX(slv_weom,$G$13),1)", FMT_IDX),
        ("w x ec at eff month", "=IF($G$12,INDEX(slv_w,$G$13)*INDEX(slv_ecp,$G$13),0)", FMT_IDX),
    ]
    for i, (lbl, f, fmt) in enumerate(helpers):
        r = 7 + i
        put(ws, f"F{r}", lbl, fnt=font(GREY_DARK, size=9))
        formula(ws, f"G{r}", f, fmt=fmt)
        ws[f"G{r}"].font = font(GREY_DARK, size=9)

    outputs = [
        ("Old-rate earned exposure before the date (C_pre)",
         "=SUMPRODUCT((slv_absmi<$G$7)*slv_w*slv_ecp*slv_widx)+$G$18*(1-$G$15)*$G$16", FMT_IDX,
         None),
        ("Earnable exposure on/after the date (C_post)",
         "=SUMPRODUCT((slv_absmi>$G$7)*slv_w*slv_ecp*slv_widx)+$G$18*$G$15*$G$17", FMT_IDX, None),
        ("Total CY earned exposure (D)", "=slv_den", FMT_IDX, None),
        ("Mod drift factor in force (A_mod)", "=nr_Amod_P", FMT_IDX, None),
        ("Rate earn-in factor needed (A_rate)",
         "=IF(OR(nr_LRcur=0,$C$7=\"\"),0,$C$7/(nr_LRcur*$C$14*nr_AOther))",
         FMT_IDX, None),
        ("Earned rate level needed (E_CY)", "=IF($C$15=0,0,nr_CRLind/$C$15)", FMT_IDX, None),
        ("REQUIRED CHANGE r",
         '=IF(OR($C$12=0,$C$16=0),"n/a",($C$16*$C$13-$C$11)/$C$12-1)', "0.000%", "nr_SolverR"),
        ("Filed-rate equivalent (r / achievement)",
         '=IF(OR($C$9=0,NOT(ISNUMBER($C$17))),"n/a",$C$17/$C$9)', "0.000%", None),
        ("Share of CY earned exposure on/after eff", "=$C$12/($C$11+$C$12)", FMT_PCT, None),
    ]
    for i, (lbl, f, fmt, name) in enumerate(outputs):
        r = 11 + i
        label(ws, f"B{r}", lbl, bold=(name == "nr_SolverR"))
        formula(ws, f"C{r}", f, fmt=fmt, border=BORDER_THIN, bold=(name == "nr_SolverR"),
                fill=FILL_PANEL if name == "nr_SolverR" else None)
        if name:
            ctx.define(name, "Solver", f"$C${r}", "Mode A required rate change")
    label(ws, "B21", "Reasonability bound |r|")
    input_cell(ws, "C21", ctx.cfg.solver_max_rate, fmt=FMT_PCT, required=False)
    ctx.define("nr_SolvMaxRate", "Solver", "$C$21", "Solver reasonability bound (config)")
    label(ws, "B22", "Minimum post-eff exposure share")
    input_cell(ws, "C22", ctx.cfg.solver_min_post_share, fmt=FMT_PCT, required=False)
    ctx.define("nr_SolvMinShare", "Solver", "$C$22", "Solver post-share warning threshold (config)")
    formula(ws, "D19",
            '=IF($C$19<nr_SolvMinShare,"WARNING: only "&TEXT($C$19,"0.0%")&" of CY exposure '
            'sits on/after this date - late-year targets are mathematically absurd","OK")',
            fmt=FMT_GEN)
    ws["D19"].font = font(FAIL_RED, size=9, italic=True)
    formula(ws, "D17",
            '=IF(ISNUMBER($C$17),IF(ABS($C$17)>nr_SolvMaxRate,"WARNING: exceeds the '
            'reasonability bound","OK"),"no solvable rate at this date")', fmt=FMT_GEN)
    ws["D17"].font = font(FAIL_RED, size=9, italic=True)

    # ---- Mode B ----
    section(ws, 25, "B",
            "Mode B — timing sensitivity: the FULL-YEAR plan LR under each possible "
            "start month for the same change")
    label(ws, "B26", "Rate change r to time")
    input_cell(ws, "C26", 0.05, fmt=FMT_PCT)
    note(ws, "D26",
         "Each row is a separate what-if, NOT the loss ratio moving through the year: the "
         "same increase made effective later earns in less during the plan year, so the "
         "full-year result lands higher. Mode A's target flags which start months still work.")
    header_row(ws, 28, 2, ["Effective month", "Effective", "C_pre", "C_post", "E_CY(P)",
                           "Full-yr plan LR", "Meets target?"],
               widths=[13, 11, 10, 10, 10, 12, 13])
    slv = ctx.lay_dyn["solver"]
    bf = slv["first"]
    for m in range(1, 13):
        r = 28 + m
        row_m = bf + 24 + (m - 1)          # cohort row of month m of P
        formula(ws, f"B{r}", f'=TEXT(DATE(nr_PlanYear,{m},1),"mmm yyyy")', align=ALIGN_C)
        formula(ws, f"C{r}", f"=DATE(nr_PlanYear,{m},1)", fmt=FMT_DATE, align=ALIGN_C)
        link(ws, f"D{r}", f"='_calc'!$T${row_m - 1}", fmt=FMT_IDX, align=ALIGN_C)
        link(ws, f"E{r}", f"=slv_total-'_calc'!$T${row_m}+'_calc'!$U${row_m}", fmt=FMT_IDX,
             align=ALIGN_C)
        formula(ws, f"F{r}", f"=($D{r}+(1+$C$26)*$E{r})/slv_den", fmt=FMT_IDX, align=ALIGN_C)
        formula(ws, f"G{r}",
                f"=IF(nr_LRcur=0,0,nr_LRcur*(nr_CRLind/$F{r})*$C$14*nr_AOther)",
                fmt=FMT_PCT, align=ALIGN_C)
        formula(ws, f"H{r}", f'=IF($C$7="","-",IF($G{r}<=$C$7,"YES","no"))', align=ALIGN_C)
        formula(ws, f"I{r}", f'=IF($H{r}="YES",{m},0)', fmt=FMT_INT)
        ws[f"I{r}"].font = font(GREY_DARK, size=8)
        formula(ws, f"J{r}", f'=IF($C$7="","",$C$7)', fmt=FMT_PCT)
        ws[f"J{r}"].font = font(GREY_DARK, size=8)
    label(ws, "B42", "Latest start month that still meets the target", bold=True)
    formula(ws, "D42",
            '=IF(MAX($I$29:$I$40)=0,"No start month meets the target",'
            'TEXT(DATE(nr_PlanYear,MAX($I$29:$I$40),1),"mmm yyyy"))', bold=True,
            fill=FILL_PANEL)
    put(ws, "I28", "helper", fnt=font(GREY_DARK, size=8))
    put(ws, "J28", "target", fnt=font(GREY_DARK, size=8))

    lc = LineChart()
    data = Reference(ws, min_col=7, min_row=28, max_row=40)
    lc.add_data(data, titles_from_data=True)
    tgt = Reference(ws, min_col=10, min_row=28, max_row=40)
    lc.add_data(tgt, titles_from_data=True)
    cats = Reference(ws, min_col=2, min_row=29, max_row=40)
    lc.set_categories(cats)
    _line_color(lc.series[0], NAVY)
    _line_color(lc.series[1], FAIL_RED, width_pt=1.25, dashed=True)
    lc.legend = None
    lc.y_axis.number_format = "0.0%"
    _style_chart(lc,
                 "Timing sensitivity: full-year plan LR if the change starts in each "
                 "month (later start = less earned benefit; dashed = target)",
                 y_title="Full-year plan LR", height=8, width=14)
    ws.add_chart(lc, "K28")

    set_widths(ws, {"A": 2, "B": 38, "C": 12, "D": 30, "E": 3, "F": 30, "G": 11})
    presentation_setup(ws, gridlines_off=True, tab_color=STEEL)
    print_setup(ws)


# ---------------------------------------------------------------------------
# Attribution (E2)
# ---------------------------------------------------------------------------


def build_attribution(ctx: Ctx):
    ws = ctx.wb["Attribution"]
    p = ctx.p
    title(ws, "B2", "Attribution — plan vs actual",
          "Sequential multiplicative decomposition of (actual - plan) CY LR. Order used: "
          "1) rate magnitude at planned dates, 2) rate timing at actual dates, 3) mod drift, "
          "4) loss-side residual. The decomposition is order-dependent (see Methodology).")
    label(ws, "B3", "Selected:")
    link(ws, "C3", "=nr_SelKey", bold=True)
    formula(ws, "E3",
            '=IF(nr_NetMode,"NOTE: this combo carries a NET RATE SELECTION - premium-side '
            'steps 1-3 are suspended (factors = 1) because rate and mods are merged in the '
            'net path; plan-vs-actual deviation flows to the residual (D39).","")')
    ws["E3"].font = font(FAIL_RED, size=9, italic=True)

    section(ws, 5, "B", "Actuals — mod path and outcome (blank = as planned)")
    att_inputs = [
        ("Actual M_0", FMT_MOD, "att_m0", "Attribution: actual current written mod (blank = plan)"),
        ("Actual M_0 as-of", FMT_DATE, "att_asof", "Attribution: actual as-of date (blank = plan)"),
        ("Actual M_1 at 12/31/P", FMT_MOD, "att_m1", "Attribution: actual written mod at 12/31/P"),
        ("Actual M_2 (optional)", FMT_MOD, "att_m2", "Attribution: actual written mod at 12/31/(P+1)"),
        (f"ACTUAL CY {p} loss ratio", FMT_PCT, "att_actlr", "Attribution: actual CY LR outcome"),
    ]
    for i, (lbl, fmt, name, desc) in enumerate(att_inputs):
        r = 6 + i
        label(ws, f"B{r}", lbl, bold=(name == "att_actlr"))
        input_cell(ws, f"C{r}", None, fmt=fmt, required=(name == "att_actlr"))
        ctx.define(name, "Attribution", f"$C${r}", desc)
    note(ws, "D6", "The plan M_ind stays fixed - it is an indication property, not an actual (D31).")

    section(ws, 12, "B", "Planned rate actions — enter actuals (blank = as planned)")
    rank_rng = ctx.lay_dyn["att_rank"]
    header_row(ws, 13, 2, ["#", "Planned effective", "Planned eff chg",
                           "ACTUAL achieved chg", "ACTUAL effective"],
               widths=[5, 14, 13, 15, 14])
    for nslot in range(1, 9):
        r = 13 + nslot
        put(ws, f"B{r}", nslot, fnt=font(GREY_DARK, size=9), align=ALIGN_C)
        link(ws, f"C{r}",
             f'=IF(COUNTIF({rank_rng},{nslot})=0,"",INDEX(rl_eff,MATCH({nslot},{rank_rng},0)))',
             fmt=FMT_DATE, align=ALIGN_C)
        link(ws, f"D{r}",
             f'=IF($C{r}="","",INDEX(rl_reff,MATCH({nslot},{rank_rng},0)))',
             fmt=FMT_PCT, align=ALIGN_C)
        input_cell(ws, f"E{r}", None, fmt=FMT_PCT, required=False)
        input_cell(ws, f"F{r}", None, fmt=FMT_DATE, required=False)
    ctx.define("att_actpct", "Attribution", "$E$14:$E$21",
               "Attribution: actual achieved total change per planned row (blank = plan)")
    ctx.define("att_actdate", "Attribution", "$F$14:$F$21",
               "Attribution: actual effective date per planned row (blank = plan)")

    section(ws, 23, "B", "Decomposition")
    header_row(ws, 24, 2, ["Step", "Factor", "LR after", "Step"],
               widths=[30, 10, 10, 11])
    e_mag, e_time = ctx.lay_dyn["att_e_mag"], ctx.lay_dyn["att_e_time"]
    mbar_act = ctx.lay_dyn["att_mbar_act"]
    steps = [
        ("Plan CY LR", None, "=nr_CYLR_P"),
        ("1. Rate magnitude (achieved %, planned dates)",
         f"=IF(nr_NetMode,1,nr_ECY_P/{e_mag})", None),
        ("2. Rate timing (achieved %, actual dates)",
         f"=IF(nr_NetMode,1,{e_mag}/{e_time})", None),
        ("3. Mod drift (actual mod path)",
         f'=IF(nr_NetMode,1,IF(AND(nr_ModAdjMaster="ON",nr_ModAdjRow="ON"),'
         f"nr_MEarned_P/{mbar_act},1))", None),
        ("4. Loss-side residual", "=IF(N(att_actlr)=0,1,att_actlr/$D$28)", None),
        ("Actual CY LR", None, "=$D$29"),
    ]
    for i, (lbl, fac, lr_f) in enumerate(steps):
        r = 25 + i
        label(ws, f"B{r}", lbl, bold=(i in (0, 5)))
        if fac:
            formula(ws, f"C{r}", fac, fmt=FMT_IDX, align=ALIGN_C)
            formula(ws, f"D{r}", f"=$D{r - 1}*$C{r}", fmt=FMT_PCT, align=ALIGN_C)
            formula(ws, f"E{r}", f"=($D{r}-$D{r - 1})*100", fmt=PTS_Z, align=ALIGN_C)
        else:
            formula(ws, f"D{r}", lr_f, fmt=FMT_PCT, align=ALIGN_C, bold=True,
                    fill=FILL_PANEL if i == 5 else None)
    formula(ws, "B32",
            '=IF(N(att_actlr)=0,"Enter the actual CY LR above to complete the residual step.",'
            'IF(ABS($D$25*$C$26*$C$27*$C$28*$C$29-att_actlr)<0.0000001,'
            '"RECONCILES: factors multiply back to the actual CY LR.",'
            '"DOES NOT RECONCILE - check inputs."))', bold=True)
    ctx.define("att_reconciles", "Attribution", "$B$32", "Attribution reconciliation status")
    note(ws, "B34",
         "Steps 1-3 recompute the CY earned rate index / earned mod with actuals substituted "
         "sequentially; step 4 is the remainder. A different ordering would allocate the "
         "interaction terms differently (stated per §7 E2). Taken rows are assumed booked as "
         "logged; only planned rows are re-achieved.")

    # waterfall chart data + chart
    cd = 37
    put(ws, f"B{cd - 1}", "Waterfall chart data (formulas — do not edit)", fnt=F_SMALL_IT)
    cats = [("Plan CY LR", "$D$25", None, True),
            ("Rate magnitude", "$D$25", "$D$26", False),
            ("Rate timing", "$D$26", "$D$27", False),
            ("Mod drift", "$D$27", "$D$28", False),
            ("Loss residual", "$D$28", "$D$29", False),
            ("Actual CY LR", "$D$30", None, True)]
    for j, h in enumerate(["Step", "base", "up", "down", "total"]):
        put(ws, f"{col(2 + j)}{cd}", h, fnt=font(GREY_DARK, size=9))
    for i, (lbl, a, b, is_total) in enumerate(cats):
        rr = cd + 1 + i
        put(ws, f"B{rr}", lbl, fnt=font(GREY_DARK, size=9))
        if is_total:
            for cL, f in (("C", "=0"), ("D", "=0"), ("E", "=0"), ("F", f"={a}")):
                formula(ws, f"{cL}{rr}", f, fmt=FMT_PCT)
        else:
            formula(ws, f"C{rr}", f"=MIN({a},{b})", fmt=FMT_PCT)
            formula(ws, f"D{rr}", f"=MAX(0,{b}-{a})", fmt=FMT_PCT)
            formula(ws, f"E{rr}", f"=MAX(0,{a}-{b})", fmt=FMT_PCT)
            formula(ws, f"F{rr}", "=0", fmt=FMT_PCT)
        for cL in "CDEF":
            ws[f"{cL}{rr}"].font = font(GREY_DARK, size=9)
    chart = BarChart()
    chart.type = "col"
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.gapWidth = 60
    data = Reference(ws, min_col=3, min_row=cd, max_col=6, max_row=cd + 6)
    cats_ref = Reference(ws, min_col=2, min_row=cd + 1, max_row=cd + 6)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.series[0].graphicalProperties.noFill = True
    chart.series[1].graphicalProperties.solidFill = UP_BAR
    chart.series[2].graphicalProperties.solidFill = DOWN_BAR
    chart.series[3].graphicalProperties.solidFill = TOTAL_BAR
    chart.legend = None
    chart.y_axis.number_format = "0%"
    _style_chart(chart, "Plan -> actual CY LR attribution", y_title="Loss ratio",
                 height=8.5, width=14)
    ws.add_chart(chart, "H5")

    set_widths(ws, {"A": 2, "B": 34, "C": 14, "D": 15, "E": 15, "F": 14})
    presentation_setup(ws, gridlines_off=True, tab_color=STEEL)
    print_setup(ws)
