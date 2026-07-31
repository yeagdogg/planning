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
from .sheets_calc import FLOW_PUB
from .xlstyle import (
    ALIGN_C, BORDER_THIN, F_LABEL, F_SMALL_IT, FAIL_RED, FILL_NAVY, FILL_PANEL,
    GREY_DARK, NAVY, col, font, formula, header_row, input_cell, jump, label,
    link, nav_bar, presentation_setup, print_setup, put, quote_sheet as _q,
    section, set_widths, title,
)

PCT_S = "+0.0%;-0.0%;0.0%"

# ---- sheet geometry (module-level so the harness can address the exhibit) ----
PF_SUM_HDR = 8                          # summary table header row
PF_SUM_FIRST = 9                        # first roster row


def n_disp(cfg) -> int:
    """Display slots: live states + 3 roster spares (the _lists convention)."""
    return len(cfg.states) + 3


def pf_tot(cfg) -> int:
    return PF_SUM_FIRST + n_disp(cfg) + 1


def grid_starts(cfg) -> list[int]:
    """Section rows of the three grids (rate leg, mod leg, delivered). Each
    grid block: section, note, month header, nd roster rows, and one IN VIEW
    total row (D60)."""
    nd = n_disp(cfg)
    g1 = pf_tot(cfg) + 4
    g2 = g1 + 3 + nd + 1 + 3
    return [g1, g2, g2 + 3 + nd + 1 + 3]


def build_program_flow(ctx: Ctx):
    ws = ctx.wb[SHEETS.PROGRAM_FLOW]
    nd = n_disp(ctx.cfg)
    cbf, stride = L.CALC_BLOCK_FIRST, L.CALC_BLOCK_STRIDE
    title(ws, "A1", f"Program Flow — what the logged program delivers ({ctx.lob.name})",
          "Month-by-month YoY change on renewals from the rate changes and mod path AS "
          "LOGGED — taken rows plus planned rows at achievement. The descriptive twin of "
          "Net Delivery: no target, no suggestion, just the flow you are on today. "
          "'All' = the EP-weighted book view (D60).")

    def pub(key, j=0):
        cL = col(FLOW_PUB[key] + j)
        return f"'_calc'!${cL}${L.CALC_RES_FIRST}:${cL}${L.CALC_RES_LAST}"
    nav_bar(ws, 3, 1, ["Control", "State Summary", "Net Delivery", "Flow Dashboard",
                       "Rate Log", "Checks"], step=2)

    # ---- controls ----
    label(ws, "A5", "Business unit in view", bold=True)
    c = input_cell(ws, "B5", ctx.we_row["bu"])
    c.alignment = ALIGN_C
    ctx.define("pf_BU", SHEETS.PROGRAM_FLOW, "$B$5",
               "Program Flow BU in view ('All' = the EP-weighted book view, D60)")
    dv_bu = DataValidation(type="list", formula1="=lst_bu_all", allow_blank=False,
                           showErrorMessage=True)
    ws.add_data_validation(dv_bu)
    dv_bu.add("B5")
    formula(ws, "A6",
            '=IF(pf_BU="All","BOOK VIEW: EP-weighted across BUs (D60). Legs are weighted '
            'means — rate x mod need not multiply to delivered exactly under mix; '
            'delivered is the exact statistic. The delta column is per-BU: pick one BU '
            'for targeting.",'
            'IF(nr_ModAdjMaster="OFF","NOTE: master mod toggle is OFF — mod legs show a '
            'dash and delivered = the rate leg alone.",'
            'IF(SUMPRODUCT(calc_netmode*(calc_bu=pf_BU))>0,"Net-target combo(s) in view: '
            'these grids show the program AS LOGGED — the gap vs the asserted net is the '
            'point. See Net Delivery for the prescriptive view.","")))')
    ws["A6"].font = font(FAIL_RED, size=9, italic=True)

    # ---- per-state summary table ----
    hdr, first = PF_SUM_HDR, PF_SUM_FIRST
    headers = ["State", "Adj plan EP (000s)", "Net target", "Avg YoY rate leg",
               "Avg YoY mod leg", "Avg YoY delivered net", "Δ vs net assertion",
               "Plan LR gap: program vs asserted (pts)"]
    header_row(ws, hdr, 1, headers, widths=[7, 12, 10, 12, 12, 13, 12, 14])
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
        # each metric: =IF(blank,"",IF(All,<EP-weighted book>,<single-BU block read>))
        sg = f'IF($P{r}=0,"—",'                       # single-BU inner guard
        stc = f"(calc_state=$A{r})"
        epd = f"SUMIFS(calc_ep,calc_state,$A{r})"

        def _all_avg(key):
            return (f'IF({epd}=0,"—",'
                    f"SUMPRODUCT({stc}*calc_ep*{pub(key)})/{epd}-1)")

        all_c = (f'IF(SUMIFS(calc_netmode,calc_state,$A{r})=0,"—",'
                 f"SUMIFS(calc_netx,calc_state,$A{r})/"
                 f"SUMIFS(calc_netmode,calc_state,$A{r}))")
        me_d = f"SUMPRODUCT({stc}*calc_ep*{pub('modeff')})"
        all_e = (f'IF({me_d}=0,"—",'
                 f"SUMPRODUCT({stc}*calc_ep*{pub('modeff')}*{pub('avg_mod')})/{me_d}-1)")
        one_c = sg + f"IF(INDEX('_calc'!$L:$L,$P{r}),INDEX('_calc'!$M:$M,$P{r}),\"—\"))"
        one_d = sg + f"INDEX('_calc'!$H:$H,$P{r}+51)-1)"
        one_e = sg + (f"IF(INDEX('_calc'!$O:$O,$P{r})=0,\"—\","
                      f"INDEX('_calc'!$I:$I,$P{r}+51)-1))")
        one_f = sg + f"INDEX('_calc'!$J:$J,$P{r}+51)-1)"
        one_g = sg + (f"IF(INDEX('_calc'!$L:$L,$P{r}),"
                      f"INDEX('_calc'!$J:$J,$P{r}+51)-1-INDEX('_calc'!$M:$M,$P{r}),"
                      f"\"—\"))")
        # D65: the same disagreement one level up — plan LR, not the legs
        # the All branch averages over the state's NET combos; the single-BU
        # branch must gate on THAT combo's own net flag, not the state's
        nm_st = f"SUMIFS(calc_netmode,calc_state,$A{r})"
        one_h = (f'IF($P{r}=0,"—",IF(INDEX(calc_netmode,MATCH($O{r},calc_key,0))=0,"—",'
                 f"INDEX(calc_proggap,MATCH($O{r},calc_key,0))))")
        all_h = (f'IF({nm_st}=0,"—",'
                 f"SUMPRODUCT((calc_state=$A{r})*calc_netmode*calc_proggap)/{nm_st})")
        formula(ws, f"H{r}",
                f'=IF($A{r}="","",IF(pf_BU="All",{all_h},{one_h}))',
                fmt='+0.00;-0.00;0.00', align=ALIGN_C, fill=band_fill,
                border=BORDER_THIN)
        for cc, allf, onef, b in (("C", all_c, one_c, False),
                                  ("D", _all_avg("avg_rate"), one_d, False),
                                  ("E", all_e, one_e, False),
                                  ("F", _all_avg("avg_del"), one_f, True),
                                  ("G", '"—"', one_g, False)):
            formula(ws, f"{cc}{r}",
                    f'=IF($A{r}="","",IF(pf_BU="All",{allf},{onef}))',
                    fmt=PCT_S, align=ALIGN_C, fill=band_fill, border=BORDER_THIN,
                    bold=b)
    tot = pf_tot(ctx.cfg)
    put(ws, f"A{tot}", "TOTAL", fnt=font(NAVY, bold=True), align=ALIGN_C)
    formula(ws, f"B{tot}", '=IF(pf_BU="All",SUM(calc_ep),SUMIFS(calc_ep,calc_bu,pf_BU))',
            fmt="#,##0", align=ALIGN_C, bold=True)
    formula(ws, f"C{tot}",
            '="all states, "&IF(pf_BU="All","every business unit",pf_BU)&" — "'
            '&IF(pf_BU="All",SUMPRODUCT(calc_netmode),'
            'SUMPRODUCT(calc_netmode*(calc_bu=pf_BU)))&" net combo(s)"')
    ws[f"C{tot}"].font = F_LABEL
    ctx.define("pfd_states", SHEETS.PROGRAM_FLOW,
               f"$A${first}:$A${first + nd - 1}", "Program Flow roster state column")
    ctx.define("pfd_gap", SHEETS.PROGRAM_FLOW,
               f"$Q${first}:$Q${first + nd - 1}",
               "Abs gap: program-basis delivered avg vs the asserted net (0 = non-net)")
    put(ws, f"A{tot + 2}",
        "Averages are written-weighted means of the monthly YoY ratios over the plan "
        "year (the Net Delivery convention); under 'All' they are EP-weighted across "
        "BUs, where delivered is the exact statistic (D60). Deep dive on one combo: set "
        "it on Control, then see the Flow Dashboard's written legs and locked/planned "
        "split.",
        fnt=F_SMALL_IT)

    # ---- three state x month grids ----
    g1, g2, g3 = grid_starts(ctx.cfg)
    secs = [
        (g1, "rate", "The rate leg — YoY written rate change on renewals",
         "History and planned filings at achievement, day-blended in their effective "
         "months. Steps mark anniversaries; a cliff is a filing finishing its year."),
        (g2, "mod", "The pricing leg — YoY written schedule-mod change on renewals",
         "Drift along the anchored mod path (dash = the mod adjustment is off for the "
         "combo). Slow-motion price change, same YoY lens."),
        (g3, "delivered", "Delivered net — rate x pricing, the YoY change customers "
         "actually see",
         "The product of the two legs. For net-target combos compare against the "
         "asserted net: the shortfall months are Net Delivery's to close."),
    ]
    roll = ("Bottom row: every state above collapsed onto one line, weighted by "
            "adjusted plan EP x written weight — the whole book when the picker is "
            "'All' (BOOK AVG), otherwise the selected business unit (BU AVG). It is "
            "an average of the state rows, not a total.")
    for g0, key, head, note_txt in secs:
        section(ws, g0, "A", head)
        put(ws, f"A{g0 + 1}", f"{note_txt}  {roll}", fnt=F_SMALL_IT)
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
                if key == "rate":
                    one = f'IF($P${rs}=0,"",{nr_}-1)'
                elif key == "mod":
                    one = (f'IF($P${rs}=0,"",'
                           f"IF(INDEX('_calc'!$O:$O,$P${rs})=0,\"—\",{or_}-1))")
                else:
                    one = (f'IF($P${rs}=0,"",{nr_}*'
                           f"IF(INDEX('_calc'!$O:$O,$P${rs})=1,{or_},1)-1)")
                # All view (D60): EP x w weighted mean across the state's BUs
                # (w cancels within a state — state-keyed seasonality)
                sw = f"(calc_state=$A{r})*{pub('epw', j)}"
                if key == "mod":
                    dn = f"SUMPRODUCT({sw}*{pub('modeff')})"
                    alv = (f'IF({dn}=0,"—",'
                           f"SUMPRODUCT({sw}*{pub('modeff')}*{pub('mod', j)})/{dn})")
                else:
                    dn = f"SUMPRODUCT({sw})"
                    alv = (f'IF({dn}=0,"",'
                           f"SUMPRODUCT({sw}*{pub(key, j)})/{dn})")
                f = f'=IF($A{r}="","",IF(pf_BU="All",{alv},{one}))'
                formula(ws, ws.cell(row=r, column=2 + j).coordinate, f,
                        fmt=PCT_S, align=ALIGN_C)
                ws.cell(row=r, column=2 + j).font = font(GREY_DARK, size=9)
        # roll-up row (D60): the BU / book-level monthly flow line — every
        # state in view collapsed onto one line, EP x weight weighted.
        # View condition is purely multiplicative (no IF over arrays):
        # ((pf_BU="All")+(calc_bu=pf_BU)) is 1 for every combo under All,
        # else 1 only for the picked BU's combos.
        r_t = g0 + 3 + nd
        formula(ws, f"A{r_t}", '=IF(pf_BU="All","BOOK AVG","BU AVG")',
                align=ALIGN_C)
        ws[f"A{r_t}"].font = font(NAVY, bold=True, size=9)
        for j in range(12):
            vw = f'((pf_BU="All")+(calc_bu=pf_BU))*{pub("epw", j)}'
            ext = f"*{pub('modeff')}" if key == "mod" else ""
            dn = f"SUMPRODUCT({vw}{ext})"
            f = (f'=IF({dn}=0,"{"—" if key == "mod" else ""}",'
                 f"SUMPRODUCT({vw}{ext}*{pub(key, j)})/{dn})")
            cc = ws.cell(row=r_t, column=2 + j)
            formula(ws, cc.coordinate, f, fmt=PCT_S, align=ALIGN_C)
            cc.font = font(NAVY, bold=True, size=9)
    ws.conditional_formatting.add(
        f"B{g3 + 3}:M{g3 + 2 + nd}",
        ColorScaleRule(start_type="min", start_color="D6E8D5",
                       mid_type="percentile", mid_value=50, mid_color="FFF2CC",
                       end_type="max", end_color="F4B8B8"))

    # ---- footer ----
    fn = g3 + 3 + nd + 2
    put(ws, f"A{fn}",
        "Program basis (D59): every Rate Log row enters exactly as logged — taken at "
        "filed %, planned at filed % x achievement. A net rate selection changes NOTHING "
        "here by design; this tab is the reality check against that assertion. Under "
        "'All' (and on IN VIEW rows) legs are EP x weight means — rate x mod need not "
        "multiply to delivered exactly under mix; delivered is exact (D60).",
        fnt=F_SMALL_IT)
    jump(ws, f"A{fn + 1}", f"{_q(SHEETS.NET_DELIVERY)}!A1",
         "Net combos: see Net Delivery for the filing + pricing walk that closes the gap >")
    set_widths(ws, {"A": 10, "B": 12, "C": 10, "D": 12, "E": 12, "F": 13, "G": 12,
                    "H": 10, "I": 10, "J": 10, "K": 10, "L": 10, "M": 10,
                    "N": 2, "O": 14, "P": 7, "Q": 8})
    presentation_setup(ws, gridlines_off=True, freeze=f"B{PF_SUM_FIRST}",
                       tab_color=NAVY)
    print_setup(ws)
