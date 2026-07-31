"""Program Flow sheet (DECISIONS.md D59).

The DESCRIPTIVE twin of Net Delivery: no target, no suggested change — what
the program AS LOGGED is delivering on renewals, state by state and month by
month, split into its rate and pricing (schedule-mod) legs:

  delivered(m) = [W(m)/W(m-12)] x [M_w(m)/M_w(m-12)]

Program basis: taken rows at filed %, planned rows at filed % x achievement —
net-mode supersession is deliberately NOT applied, so for net-target combos
the gap between this tab and the asserted (1+x) is the exhibit's point (Net
Delivery closes it). Everything reads the existing _calc combo blocks via the
blessed stride-INDEX — the raw written index (col N) and the written mod path
(col O) are program-basis even in net mode, and the block results rows carry
the w-weighted plan-year ratio averages (cols H/I/J). Oracle:
engine.program_flow_by_month; every displayed figure ties at 1e-9.
"""

from __future__ import annotations

from openpyxl.formatting.rule import ColorScaleRule
from openpyxl.worksheet.datavalidation import DataValidation

from .build_workbook import Ctx, Layout as L, SHEETS
from .xlstyle import (
    ALIGN_C, BORDER_THIN, F_LABEL, F_SMALL_IT, FAIL_RED, FILL_NAVY, FILL_PANEL,
    GREY_DARK, NAVY, font, formula, header_row, input_cell, jump, label, link,
    nav_bar, presentation_setup, print_setup, put, quote_sheet as _q, section,
    set_widths, title,
)

PCT_S = "+0.0%;-0.0%;0.0%"

# ---- sheet geometry (module-level so the harness can address the exhibit) ----
PF_SUM_HDR = 8                          # summary table header row
PF_SUM_FIRST = 9                        # first roster row


def n_disp(ctx) -> int:
    """Display slots: live states + 3 roster spares (the _lists convention)."""
    return len(ctx.cfg.states) + 3


def pf_tot(ctx) -> int:
    return PF_SUM_FIRST + n_disp(ctx) + 1


def grid_starts(ctx) -> list[int]:
    """Section rows of the three grids (rate leg, mod leg, delivered)."""
    nd = n_disp(ctx)
    g1 = pf_tot(ctx) + 4
    g2 = g1 + 3 + nd + 3
    return [g1, g2, g2 + 3 + nd + 3]


def build_program_flow(ctx: Ctx):
    ws = ctx.wb[SHEETS.PROGRAM_FLOW]
    nd = n_disp(ctx)
    cbf, stride = L.CALC_BLOCK_FIRST, L.CALC_BLOCK_STRIDE
    title(ws, "A1", f"Program Flow — what the logged program delivers ({ctx.lob.name})",
          "Month-by-month YoY change on renewals from the rate changes and mod path AS "
          "LOGGED — taken rows plus planned rows at achievement. The descriptive twin of "
          "Net Delivery: no target, no suggestion, just the flow you are on today.")
    nav_bar(ws, 3, 1, ["Control", "State Summary", "Net Delivery", "Flow Dashboard",
                       "Rate Log", "Checks"], step=2)

    # ---- controls ----
    label(ws, "A5", "Business unit in view", bold=True)
    c = input_cell(ws, "B5", ctx.we_row["bu"])
    c.alignment = ALIGN_C
    ctx.define("pf_BU", SHEETS.PROGRAM_FLOW, "$B$5",
               "Program Flow BU in view (flow needs a single BU's rate history)")
    dv_bu = DataValidation(type="list", formula1="=lst_bu_all", allow_blank=False,
                           showErrorMessage=True)
    ws.add_data_validation(dv_bu)
    dv_bu.add("B5")
    formula(ws, "A6",
            '=IF(pf_BU="All","Pick a single business unit — the grids and averages need '
            'one BU\'s rate history.",'
            'IF(nr_ModAdjMaster="OFF","NOTE: master mod toggle is OFF — mod legs show a '
            'dash and delivered = the rate leg alone.",'
            'IF(SUMPRODUCT(calc_netmode*(calc_bu=pf_BU))>0,"Net-target combo(s) in view: '
            'these grids show the program AS LOGGED — the gap vs the asserted net is the '
            'point. See Net Delivery for the prescriptive view.","")))')
    ws["A6"].font = font(FAIL_RED, size=9, italic=True)

    # ---- per-state summary table ----
    hdr, first = PF_SUM_HDR, PF_SUM_FIRST
    headers = ["State", "Adj plan EP (000s)", "Net target", "Avg YoY rate leg",
               "Avg YoY mod leg", "Avg YoY delivered net", "Δ vs net assertion"]
    header_row(ws, hdr, 1, headers, widths=[7, 12, 10, 12, 12, 13, 12])
    ws.row_dimensions[hdr].height = 42
    for h, cL in (("key", "O"), ("blk row", "P"), ("|gap|", "Q")):
        put(ws, f"{cL}{hdr}", h, fnt=font(GREY_DARK, size=8), align=ALIGN_C)
    for i in range(nd):
        r = first + i
        band_fill = FILL_PANEL if i % 2 else None
        link(ws, f"A{r}", f"=_lists!$B${3 + i}", align=ALIGN_C, fill=band_fill,
             border=BORDER_THIN, bold=True)
        # grey helpers: resolved key and combo-block top row (0 = not in view)
        formula(ws, f"O{r}", f'=IF(OR($A{r}="",pf_BU="All"),"",pf_BU&"|"&$A{r})')
        formula(ws, f"P{r}",
                f'=IF($O{r}="",0,IF(COUNTIF(calc_key,$O{r})=0,0,'
                f"{cbf}+(MATCH($O{r},calc_key,0)-1)*{stride}))", fmt="0")
        formula(ws, f"Q{r}",
                f"=IF($P{r}=0,0,IF(INDEX('_calc'!$L:$L,$P{r}),"
                f"ABS(INDEX('_calc'!$J:$J,$P{r}+51)-1-INDEX('_calc'!$M:$M,$P{r})),0))",
                fmt="0.0000")
        for cL in "OPQ":
            ws[f"{cL}{r}"].font = font(GREY_DARK, size=8)
        formula(ws, f"B{r}",
                f'=IF($A{r}="","",IF(pf_BU="All",SUMIFS(calc_ep,calc_state,$A{r}),'
                f"SUMIFS(calc_ep,calc_state,$A{r},calc_bu,pf_BU)))",
                fmt="#,##0;-#,##0;\"\"", align=ALIGN_C, fill=band_fill, border=BORDER_THIN)
        guard = f'=IF(OR($A{r}="",pf_BU="All"),"",IF($P{r}=0,"—",'
        formula(ws, f"C{r}",
                guard + f"IF(INDEX('_calc'!$L:$L,$P{r}),INDEX('_calc'!$M:$M,$P{r}),\"—\")))",
                fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN)
        formula(ws, f"D{r}",
                guard + f"INDEX('_calc'!$H:$H,$P{r}+51)-1))",
                fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN)
        formula(ws, f"E{r}",
                guard + f"IF(INDEX('_calc'!$O:$O,$P{r})=0,\"—\","
                        f"INDEX('_calc'!$I:$I,$P{r}+51)-1)))",
                fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN)
        formula(ws, f"F{r}",
                guard + f"INDEX('_calc'!$J:$J,$P{r}+51)-1))",
                fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN, bold=True)
        formula(ws, f"G{r}",
                guard + f"IF(INDEX('_calc'!$L:$L,$P{r}),"
                        f"INDEX('_calc'!$J:$J,$P{r}+51)-1-INDEX('_calc'!$M:$M,$P{r}),"
                        f"\"—\")))",
                fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN)
    tot = pf_tot(ctx)
    put(ws, f"A{tot}", "IN VIEW", fnt=font(NAVY, bold=True), align=ALIGN_C)
    formula(ws, f"B{tot}", '=IF(pf_BU="All",SUM(calc_ep),SUMIFS(calc_ep,calc_bu,pf_BU))',
            fmt="#,##0", align=ALIGN_C, bold=True)
    formula(ws, f"C{tot}",
            '=IF(pf_BU="All",SUMPRODUCT(calc_netmode),'
            'SUMPRODUCT(calc_netmode*(calc_bu=pf_BU)))&" net combo(s)"')
    ws[f"C{tot}"].font = F_LABEL
    ctx.define("pfd_states", SHEETS.PROGRAM_FLOW,
               f"$A${first}:$A${first + nd - 1}", "Program Flow roster state column")
    ctx.define("pfd_gap", SHEETS.PROGRAM_FLOW,
               f"$Q${first}:$Q${first + nd - 1}",
               "Abs gap: program-basis delivered avg vs the asserted net (0 = non-net)")
    put(ws, f"A{tot + 2}",
        "Averages are written-weighted means of the monthly YoY ratios over the plan "
        "year (the Net Delivery convention). Deep dive on one combo: set it on Control, "
        "then see the Flow Dashboard's written legs and locked/planned split.",
        fnt=F_SMALL_IT)

    # ---- three state x month grids ----
    g1, g2, g3 = grid_starts(ctx)
    secs = [
        (g1, "The rate leg — YoY written rate change on renewals",
         "History and planned filings at achievement, day-blended in their effective "
         "months. Steps mark anniversaries; a cliff is a filing finishing its year."),
        (g2, "The pricing leg — YoY written schedule-mod change on renewals",
         "Drift along the anchored mod path (dash = the mod adjustment is off for the "
         "combo). Slow-motion price change, same YoY lens."),
        (g3, "Delivered net — rate x pricing, the YoY change customers actually see",
         "The product of the two legs. For net-target combos compare against the "
         "asserted net: the shortfall months are Net Delivery's to close."),
    ]
    for g0, head, note_txt in secs:
        section(ws, g0, "A", head)
        put(ws, f"A{g0 + 1}", note_txt, fnt=F_SMALL_IT)
        put(ws, f"A{g0 + 2}", "State", fnt=font("FFFFFF", bold=True), fill=FILL_NAVY,
            align=ALIGN_C)
        for j in range(12):
            cell = ws.cell(row=g0 + 2, column=2 + j)
            cell.value = f'=TEXT(DATE(nr_PlanYear,{j + 1},1),"mmm")&" "&nr_PlanYear'
            cell.font = font("FFFFFF", bold=True, size=9)
            cell.fill = FILL_NAVY
            cell.alignment = ALIGN_C
        for i in range(nd):
            r = g0 + 3 + i
            rs = first + i
            link(ws, f"A{r}", f"=_lists!$B${3 + i}", align=ALIGN_C, bold=True)
            for j in range(12):
                nr_ = (f"INDEX('_calc'!$N:$N,$P${rs}+{27 + j})"
                       f"/INDEX('_calc'!$N:$N,$P${rs}+{15 + j})")
                or_ = (f"INDEX('_calc'!$O:$O,$P${rs}+{27 + j})"
                       f"/INDEX('_calc'!$O:$O,$P${rs}+{15 + j})")
                if g0 == g1:
                    f = f'=IF($P${rs}=0,"",{nr_}-1)'
                elif g0 == g2:
                    f = (f'=IF($P${rs}=0,"",'
                         f"IF(INDEX('_calc'!$O:$O,$P${rs})=0,\"—\",{or_}-1))")
                else:
                    f = (f'=IF($P${rs}=0,"",{nr_}*'
                         f"IF(INDEX('_calc'!$O:$O,$P${rs})=1,{or_},1)-1)")
                formula(ws, ws.cell(row=r, column=2 + j).coordinate, f,
                        fmt=PCT_S, align=ALIGN_C)
                ws.cell(row=r, column=2 + j).font = font(GREY_DARK, size=9)
    ws.conditional_formatting.add(
        f"B{g3 + 3}:M{g3 + 2 + nd}",
        ColorScaleRule(start_type="min", start_color="D6E8D5",
                       mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                       end_type="max", end_color="F4B8B8"))

    # ---- footer ----
    fn = g3 + 3 + nd + 1
    put(ws, f"A{fn}",
        "Program basis (D59): every Rate Log row enters exactly as logged — taken at "
        "filed %, planned at filed % x achievement. A net rate selection changes NOTHING "
        "here by design; this tab is the reality check against that assertion.",
        fnt=F_SMALL_IT)
    jump(ws, f"A{fn + 1}", f"{_q(SHEETS.NET_DELIVERY)}!A1",
         "Net combos: see Net Delivery for the filing + pricing walk that closes the gap >")
    set_widths(ws, {"A": 8, "B": 12, "C": 10, "D": 12, "E": 12, "F": 13, "G": 12,
                    "H": 10, "I": 10, "J": 10, "K": 10, "L": 10, "M": 10,
                    "N": 2, "O": 14, "P": 7, "Q": 8})
    presentation_setup(ws, gridlines_off=True, freeze=f"B{PF_SUM_FIRST}",
                       tab_color=NAVY)
    print_setup(ws)
