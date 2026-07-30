"""Sheet builders: _lists (validation), Inputs (the §4 tables), Control.

Dimensioning (D37): each workbook covers ONE line of business; input rows are
BU x STATE combinations and the selector key is "BU|State". The policy term is
a single workbook-level input; tbl_Seasonality is keyed by state.
"""

from __future__ import annotations

from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.worksheet.table import Table, TableStyleInfo

from .build_workbook import Ctx, Layout as L
from .xlstyle import (
    ALIGN_C, ALIGN_L, ALIGN_WRAP, BORDER_THIN, F_HEADER, F_LABEL, F_SMALL, F_SMALL_IT,
    F_SUB, FAIL_RED, FILL_GREY, FILL_NAVY, FILL_PANEL, FILL_PANEL_2, FILL_RED,
    FILL_GREEN, FMT_DATE, FMT_EP, FMT_GEN, FMT_IDX, FMT_INT, FMT_MOD, FMT_PCT,
    FMT_PTS_SIGNED, GREY_DARK, NAVY, PASS_GREEN, STEEL, WARN_AMBER,
    font, formula, header_row, input_cell, label, link, note, presentation_setup,
    print_setup, put, rng, section, set_widths, title,
)

TABLE_STYLE = TableStyleInfo(
    name="TableStyleMedium2", showFirstColumn=False, showLastColumn=False,
    showRowStripes=True, showColumnStripes=False,
)


def _dv(ws, kind, ranges, **kw):
    dv = DataValidation(type=kind, allow_blank=True, showErrorMessage=True, **kw)
    ws.add_data_validation(dv)
    for r in ranges:
        dv.add(r)
    return dv


# ---------------------------------------------------------------------------
# _lists
# ---------------------------------------------------------------------------


def build_lists(ctx: Ctx):
    ws = ctx.wb["_lists"]
    label(ws, "A1", "Validation lists (hidden sheet; referenced by data-validation dropdowns)", bold=True)
    cols = {
        "A": ("lst_bu", "Business units", list(ctx.cfg.business_units)),
        "B": ("lst_state", "States", list(ctx.cfg.states)),
        "C": ("lst_status", "Rate change status", ["taken", "planned"]),
        "D": ("lst_yn", "Yes/No flags", ["Y", "N"]),
        "E": ("lst_basis", "LR basis", ["current", "proposed"]),
        "F": ("lst_onoff", "Toggles", ["ON", "OFF"]),
        "G": ("lst_scenario", "Scenario selector", ["Base", "S1", "S2", "S3", "S4"]),
        "H": ("lst_basisdisp", "Projected-LR KPI display", ["As input", "Current level"]),
        "I": ("lst_bu_all", "State Summary BU filter", ["All"] + list(ctx.cfg.business_units)),
    }
    for letter, (name, desc, values) in cols.items():
        label(ws, f"{letter}2", desc, bold=True)
        for i, v in enumerate(values):
            put(ws, f"{letter}{3 + i}", v, fnt=F_LABEL)
        ctx.define(name, "_lists", f"${letter}$3:${letter}${2 + len(values)}", desc)
    set_widths(ws, {c: 18 for c in cols})


# ---------------------------------------------------------------------------
# Inputs
# ---------------------------------------------------------------------------

# Paste-friendly layout (D40): input columns are CONTIGUOUS — the key and all
# engine helper columns live to the RIGHT of each dataset, never in between.
LR_HEADERS = [
    "BU", "State", "Projected LR", "LR basis", "Selected chg s", "M_ind", "M_0",
    "M_0 as-of", "M_1", "M_prior (opt)", "M_2 (opt)", "Adj plan EP (000s)", "Net trend P+1",
    "A_other", "A_other label", "Mod adj", "Net sel P (opt)", "Net sel P+1 (opt)",
]
RL_HEADERS = ["BU", "State", "Effective", "Filed %", "Status", "Considered?",
              "Achievement %", "Comment"]
RL_HELPER_HEADERS = ["Key", "r_eff", "ln(1+r)", "Eff month", "Days on/after", "Earlier cnt",
                     "Same cnt", "First?", "Dup month?", "Seq"]


def build_inputs(ctx: Ctx):
    ws = ctx.wb["Inputs"]
    p = ctx.p
    title(ws, "A1", f"Inputs — {ctx.lob.name}",
          "All actuarial inputs live here — blue font = enter values; yellow fill = required. "
          "One row per BU x state. Structure is fixed; do not insert or delete rows.")
    put(ws, "A4",
        "SAMPLE DATA: every populated row below is illustrative and should be replaced with "
        f"your book. {ctx.we_key} carries the documented worked example used by the Checks sheet.",
        fnt=font(FAIL_RED, bold=True, italic=True))
    put(ws, "A5" if False else "T4",
        "Each dataset is one contiguous paste block (tbl_LR A:R, tbl_RateLog A:H) — keys and "
        "engine helpers sit to the right and recalculate on their own.",
        fnt=F_SMALL_IT)

    # ------------------------------------------------------------------ tbl_LR
    section(ws, L.LR_HDR - 1, "A",
            "tbl_LR — one row per BU x state (projected LR and mod inputs). "
            "Columns A:R are one contiguous paste block.")
    header_row(ws, L.LR_HDR, 1, LR_HEADERS,
               widths=[9, 8, 11, 10, 11, 8, 8, 11, 8, 11, 9, 12, 11, 8, 18, 9, 11, 11])
    put(ws, f"T{L.LR_HDR}", "Key", fnt=font(GREY_DARK, bold=True, size=9), fill=FILL_GREY)
    put(ws, f"T{L.LR_HDR - 1}", "engine helper — do not edit", fnt=F_SMALL_IT)
    fmt_by_col = {3: FMT_PCT, 4: FMT_GEN, 5: FMT_PCT, 6: FMT_MOD, 7: FMT_MOD, 8: FMT_DATE,
                  9: FMT_MOD, 10: FMT_MOD, 11: FMT_MOD, 12: FMT_EP, 13: FMT_PCT,
                  14: FMT_IDX, 15: FMT_GEN, 16: FMT_GEN, 17: FMT_PCT, 18: FMT_PCT}
    required_cols = {1, 2, 3, 4, 6, 7, 8, 9, 16}
    for i in range(L.LR_ROWS):
        r = L.LR_FIRST + i
        row = ctx.lr_rows[i] if i < len(ctx.lr_rows) else None
        vals = {
            1: row and row["bu"], 2: row and row["state"], 3: row and row["lr_proj"],
            4: row and row["basis"], 5: row and row["s"], 6: row and row["m_ind"],
            7: row and row["m0"], 8: row and row["m0_asof"], 9: row and row["m1"],
            10: row and row["m_prior"], 11: row and row["m2"], 12: row and row["ep"],
            13: row and row["trend"], 14: row and row["a_other"],
            15: (row["a_other_label"] if row else None), 16: row and row["modadj"],
            17: row and row["netp"], 18: row and row["netp1"],
        }
        for c in range(1, 19):
            input_cell(ws, f"{ws.cell(row=r, column=c).coordinate}", vals.get(c),
                       fmt=fmt_by_col.get(c, FMT_GEN), required=(c in required_cols))
        formula(ws, f"T{r}", f'=IF(OR($A{r}="",$B{r}=""),"",$A{r}&"|"&$B{r})',
                fmt=FMT_GEN, border=BORDER_THIN)
        ws[f"T{r}"].font = font(GREY_DARK, size=9)
    ws[f"P{L.LR_HDR}"].comment = Comment(
        "Per-combo mod adjustment toggle. Set OFF if the indication's premium trend already "
        "reflects schedule mod drift — otherwise the drift is double-counted. The Control "
        "sheet master toggle must also be ON for the adjustment to apply.", "generator")
    ws[f"Q{L.LR_HDR}"].comment = Comment(
        "OPTIONAL net rate selection (DECISIONS.md D39). When set, cohorts written on/after "
        "1/1/P abandon the planned rate program and the projected mod path: each cohort's "
        "combined net price (rate x mod) renews this % above its year-ago cohort. History "
        "before 1/1/P still earns exactly as modeled. Blank = explicit program (classic). "
        "The P+1 column defaults to the P selection when blank.", "generator")
    tbl = Table(displayName="tbl_LR", ref=f"A{L.LR_HDR}:R{L.LR_LAST}")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)
    lr_names = {
        "lr_bu": ("A", "tbl_LR business unit"), "lr_state": ("B", "tbl_LR state"),
        "lr_lrproj": ("C", "Projected loss ratio from the indication"),
        "lr_basis": ("D", "LR basis: current | proposed"),
        "lr_s": ("E", "Selected/indicated change s (used when basis = proposed)"),
        "lr_mind": ("F", "M_ind: avg schedule mod assumed in the indication"),
        "lr_m0": ("G", "M_0: current avg written mod"),
        "lr_m0asof": ("H", "M_0 as-of date (close of day)"),
        "lr_m1": ("I", "M_1: projected avg written mod at 12/31/P"),
        "lr_mprior": ("J", "M_prior: avg mod ~12 months before M_0 (optional)"),
        "lr_m2": ("K", "M_2: projected avg written mod at 12/31/(P+1); blank = M_1"),
        "lr_ep": ("L", "ADJUSTED plan earned premium (000s) — the weight behind Portfolio "
                       "totals and State Summary aggregates"),
        "lr_trend": ("M", "Annual net loss-over-premium trend for the P+1 view"),
        "lr_aother": ("N", "A_other manual adjustment factor"),
        "lr_aotherlbl": ("O", "Label required when A_other <> 1"),
        "lr_modadj": ("P", "Per-combo mod adjustment toggle (ON/OFF)"),
        "lr_netp": ("Q", "OPTIONAL net rate selection for P: YoY combined rate x mod target "
                         "from 1/1/P (blank = explicit program, D39)"),
        "lr_netp1": ("R", "OPTIONAL net selection for P+1 (blank = carry the P selection)"),
        "lr_key": ("T", "Helper: concatenated BU|State key (right of the paste block, D40)"),
    }
    for name, (colL, desc) in lr_names.items():
        ctx.define(name, "Inputs", f"${colL}${L.LR_FIRST}:${colL}${L.LR_LAST}", desc)

    # -------------------------------------------------------------- tbl_RateLog
    section(ws, L.RL_HDR - 1, "A",
            "tbl_RateLog — one row per rate change per BU x state   "
            "(planned rows: effective change = filed % x achievement %). "
            "Columns A:H are one contiguous paste block.")
    header_row(ws, L.RL_HDR, 1, RL_HEADERS, widths=[9, 8, 11, 9, 10, 12, 13, 34])
    header_row(ws, L.RL_HDR, 10, RL_HELPER_HEADERS,
               widths=[14, 8, 8, 11, 12, 10, 9, 7, 11, 6], fill=FILL_GREY,
               fnt=font(GREY_DARK, bold=True, size=9))
    put(ws, f"J{L.RL_HDR - 1}", "engine helpers (formulas — do not edit)", fnt=F_SMALL_IT)
    ws.column_dimensions["I"].width = 2
    for i in range(L.RL_ROWS):
        r = L.RL_FIRST + i
        row = ctx.rate_rows[i] if i < len(ctx.rate_rows) else None
        input_cell(ws, f"A{r}", row and row["bu"])
        input_cell(ws, f"B{r}", row and row["state"])
        input_cell(ws, f"C{r}", row and row["eff"], fmt=FMT_DATE)
        input_cell(ws, f"D{r}", row and row["filed"], fmt=FMT_PCT)
        input_cell(ws, f"E{r}", row and row["status"])
        input_cell(ws, f"F{r}", row and row["considered"])
        input_cell(ws, f"G{r}", row and row["achievement"], fmt=FMT_PCT, required=False)
        input_cell(ws, f"H{r}", (row["comment"] if row else None), required=False)
        ws[f"H{r}"].alignment = ALIGN_L
        # helpers J..S (right of the paste block, D40)
        formula(ws, f"J{r}", f'=IF(OR($A{r}="",$B{r}=""),"",$A{r}&"|"&$B{r})', border=BORDER_THIN)
        formula(ws, f"K{r}", f'=IF($C{r}="","",$D{r}*IF($E{r}="planned",IF($G{r}="",1,$G{r}),1))',
                fmt=FMT_PCT)
        formula(ws, f"L{r}", f'=IF($C{r}="",0,IF(1+$K{r}>0,LN(1+$K{r}),0))', fmt=FMT_IDX)
        formula(ws, f"M{r}", f'=IF($C{r}="","",DATE(YEAR($C{r}),MONTH($C{r}),1))', fmt=FMT_DATE)
        formula(ws, f"N{r}", f'=IF($C{r}="",0,EOMONTH($C{r},0)-$C{r}+1)', fmt=FMT_INT)
        formula(ws, f"O{r}", f'=IF($C{r}="",0,COUNTIFS(rl_key,$J{r},rl_effmonth,$M{r},rl_eff,"<"&$C{r}))',
                fmt=FMT_INT)
        formula(ws, f"P{r}", f'=IF($C{r}="",0,COUNTIFS($J${L.RL_FIRST}:$J{r},$J{r},$M${L.RL_FIRST}:$M{r},$M{r},$C${L.RL_FIRST}:$C{r},$C{r}))',
                fmt=FMT_INT)
        formula(ws, f"Q{r}", f'=IF($C{r}="",0,IF(AND($O{r}=0,$P{r}=1),1,0))', fmt=FMT_INT)
        formula(ws, f"R{r}", f'=IF($C{r}="",0,IF(COUNTIFS(rl_key,$J{r},rl_effmonth,$M{r})>1,1,0))',
                fmt=FMT_INT)
        formula(ws, f"S{r}",
                f'=IF($C{r}="",0,COUNTIFS(rl_key,$J{r},rl_eff,"<"&$C{r})'
                f"+COUNTIFS($J${L.RL_FIRST}:$J{r},$J{r},$C${L.RL_FIRST}:$C{r},$C{r}))",
                fmt=FMT_INT)
        for cL in "JKLMNOPQRS":
            ws[f"{cL}{r}"].font = font(GREY_DARK, size=9)
    tbl = Table(displayName="tbl_RateLog", ref=f"A{L.RL_HDR}:H{L.RL_LAST}")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)
    rl_names = {
        "rl_bu": ("A", "tbl_RateLog business unit"), "rl_state": ("B", "tbl_RateLog state"),
        "rl_eff": ("C", "Rate change effective date (start of day)"),
        "rl_filed": ("D", "Filed rate change (decimal fraction)"),
        "rl_status": ("E", "taken | planned"),
        "rl_cons": ("F", "Considered in the indication? (Y/N)"),
        "rl_ach": ("G", "Achievement % (planned rows; blank = 100%)"),
        "rl_key": ("J", "Helper: BU|State key (right of the paste block, D40)"),
        "rl_reff": ("K", "Helper: effective change = filed x achievement (planned only)"),
        "rl_ln1p": ("L", "Helper: ln(1 + effective change); 0 for blank rows"),
        "rl_effmonth": ("M", "Helper: first day of the effective month"),
        "rl_daysafter": ("N", "Helper: days in effective month on/after the change"),
        "rl_first": ("Q", "Helper: 1 for the first change in its BU|State cohort month"),
        "rl_dup": ("R", "Helper: 1 when >1 change shares a BU|State cohort month"),
        "rl_seq": ("S", "Helper: chronological rank of the change within its BU|State key"),
    }
    for name, (colL, desc) in rl_names.items():
        ctx.define(name, "Inputs", f"${colL}${L.RL_FIRST}:${colL}${L.RL_LAST}", desc)

    # --------------------------------------------------------- tbl_Seasonality
    section(ws, L.SE_HDR - 1, "A",
            "tbl_Seasonality — optional per-STATE monthly written weights (relative; blank row = uniform)")
    months = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"]
    header_row(ws, L.SE_HDR, 1, ["State"] + months + ["Row sum"],
               widths=[9] + [7] * 12 + [9])
    for i in range(L.SE_ROWS):
        r = L.SE_FIRST + i
        row = ctx.season_rows[i] if i < len(ctx.season_rows) else None
        input_cell(ws, f"A{r}", row and row["state"], required=False)
        for m in range(12):
            input_cell(ws, ws.cell(row=r, column=2 + m).coordinate,
                       (row["weights"][m] if row else None), fmt="0.00", required=False)
        formula(ws, f"N{r}", f'=IF($A{r}="","",SUM($B{r}:$M{r}))', fmt="0.00")
        ws[f"N{r}"].font = font(GREY_DARK, size=9)
    tbl = Table(displayName="tbl_Seasonality", ref=f"A{L.SE_HDR}:M{L.SE_LAST}")
    tbl.tableStyleInfo = TABLE_STYLE
    ws.add_table(tbl)
    ctx.define("se_state", "Inputs", f"$A${L.SE_FIRST}:$A${L.SE_LAST}", "Seasonality state column")
    ctx.define("se_block", "Inputs", f"$B${L.SE_FIRST}:$M${L.SE_LAST}", "Seasonality 12-month weight block")
    ctx.define("se_sum", "Inputs", f"$N${L.SE_FIRST}:$N${L.SE_LAST}", "Helper: seasonality row sums")

    # ---------------------------------------------------- workbook parameters
    section(ws, L.WB_HDR - 1, "A", "Workbook parameters")
    label(ws, f"A{L.WB_HDR}", "Line of business")
    put(ws, f"B{L.WB_HDR}", ctx.lob.name, fnt=font(NAVY, bold=True))
    label(ws, f"A{L.TERM_ROW}", "Policy term (months)")
    input_cell(ws, f"B{L.TERM_ROW}", ctx.lob.term_months, fmt=FMT_INT)
    ctx.define("nr_TermMonths", "Inputs", f"$B${L.TERM_ROW}",
               "Workbook policy term in months (this workbook covers one LOB)")
    note(ws, f"C{L.TERM_ROW}",
         "Applies to every BU x state in this workbook (one workbook per LOB, D37).")

    # ------------------------------------------------------------- validations
    # DV sits ON the input cells only (a paste passes straight over it); the
    # paste blocks contain no formula columns (D40).
    fr, lr_ = L.LR_FIRST, L.LR_LAST
    _dv(ws, "list", [f"A{fr}:A{lr_}"], formula1="=lst_bu")
    _dv(ws, "list", [f"B{fr}:B{lr_}"], formula1="=lst_state")
    _dv(ws, "list", [f"D{fr}:D{lr_}"], formula1="=lst_basis")
    _dv(ws, "list", [f"P{fr}:P{lr_}"], formula1="=lst_onoff")
    _dv(ws, "decimal", [f"C{fr}:C{lr_}"], operator="between", formula1="0", formula2="3",
        error="Loss ratio must be between 0 and 300% (entered as a fraction).")
    _dv(ws, "decimal", [f"E{fr}:E{lr_}", f"M{fr}:M{lr_}"], operator="between",
        formula1="-0.5", formula2="1", error="Enter a decimal fraction between -50% and +100%.")
    _dv(ws, "decimal", [f"F{fr}:F{lr_}", f"G{fr}:G{lr_}", f"I{fr}:I{lr_}",
                        f"J{fr}:J{lr_}", f"K{fr}:K{lr_}"], operator="between",
        formula1="0.5", formula2="1.5", error="Schedule mods must lie in [0.5, 1.5].")
    _dv(ws, "decimal", [f"L{fr}:L{lr_}"], operator="greaterThanOrEqual", formula1="0",
        error="Plan EP must be non-negative.")
    _dv(ws, "decimal", [f"N{fr}:N{lr_}"], operator="between", formula1="0.5", formula2="2",
        error="A_other must lie in [0.5, 2.0].")
    _dv(ws, "date", [f"H{fr}:H{lr_}"], operator="between",
        formula1="DATE(1990,1,1)", formula2="DATE(2100,12,31)",
        error="Enter a real as-of date (it must precede 12/31 of the plan year; "
              "the Checks sheet enforces that bound).")
    _dv(ws, "decimal", [f"Q{fr}:Q{lr_}", f"R{fr}:R{lr_}"], operator="between",
        formula1="-0.5", formula2="1",
        error="Net rate selection must lie in [-50%, +100%] (decimal fraction; blank = off).")
    fr, lr_ = L.RL_FIRST, L.RL_LAST
    _dv(ws, "list", [f"A{fr}:A{lr_}"], formula1="=lst_bu")
    _dv(ws, "list", [f"B{fr}:B{lr_}"], formula1="=lst_state")
    _dv(ws, "list", [f"E{fr}:E{lr_}"], formula1="=lst_status")
    _dv(ws, "list", [f"F{fr}:F{lr_}"], formula1="=lst_yn")
    _dv(ws, "decimal", [f"D{fr}:D{lr_}"], operator="between", formula1="-0.5", formula2="2",
        error="Filed change must lie in [-50%, +200%] (decimal fraction).")
    _dv(ws, "decimal", [f"G{fr}:G{lr_}"], operator="between", formula1="0", formula2="1.5",
        error="Achievement must lie in [0%, 150%].")
    _dv(ws, "date", [f"C{fr}:C{lr_}"], operator="between",
        formula1="DATE(1990,1,1)", formula2="DATE(2100,12,31)", error="Enter a real date.")
    _dv(ws, "decimal", [rng(2, L.SE_FIRST, 13, L.SE_LAST)], operator="between",
        formula1="0", formula2="100", error="Seasonality weights are non-negative.")
    _dv(ws, "list", [f"A{L.SE_FIRST}:A{L.SE_LAST}"], formula1="=lst_state")
    _dv(ws, "whole", [f"B{L.TERM_ROW}"], operator="between",
        formula1="1", formula2="12", error="Term must be 1..12 months.")

    presentation_setup(ws, gridlines_off=False, freeze=f"A{L.LR_FIRST}", tab_color=WARN_AMBER)
    print_setup(ws)


# ---------------------------------------------------------------------------
# Control
# ---------------------------------------------------------------------------


def build_control(ctx: Ctx):
    ws = ctx.wb["Control"]
    p = ctx.p
    title(ws, "B2", f"Calendar-Year Plan Loss Ratio — {ctx.lob.name}",
          f"Plan year {p} | one workbook per LOB; combos are BU x state | generated by "
          f"src/build_workbook.py v{ctx.cfg.version} | classic formula mode")

    # Selectors
    section(ws, 5, "B", "Selection")
    label(ws, "B6", "Plan year (P)")
    input_cell(ws, "C6", p, fmt="0")
    label(ws, "B7", "Business unit")
    input_cell(ws, "C7", ctx.we_row["bu"])
    label(ws, "B8", "State")
    input_cell(ws, "C8", ctx.we_row["state"])
    label(ws, "B9", "Scenario in focus")
    input_cell(ws, "C9", "Base", required=False)
    label(ws, "E6", "Selected key")
    formula(ws, "F6", '=$C$7&"|"&$C$8')
    label(ws, "E7", "Combo found in tbl_LR?")
    formula(ws, "F7", "=COUNTIF(lr_key,$F$6)>0")
    label(ws, "E8", "Policy term (months)")
    link(ws, "F8", "=nr_TermMonths", fmt="0")
    label(ws, "E9", "Line of business")
    put(ws, "F9", ctx.lob.name, fnt=font(NAVY, bold=True))

    ctx.define("nr_PlanYear", "Control", "$C$6", "Plan year P")
    ctx.define("nr_SelBU", "Control", "$C$7", "Selected business unit")
    ctx.define("nr_SelState", "Control", "$C$8", "Selected state")
    ctx.define("nr_SelScenario", "Control", "$C$9", "Scenario surfaced in the Control KPI row")
    ctx.define("nr_SelKey", "Control", "$F$6", "Selected BU|State key")
    ctx.define("nr_SelOK", "Control", "$F$7", "TRUE when the selected combo exists in tbl_LR")

    # Toggles
    section(ws, 11, "B", "Global toggles")
    label(ws, "B12", "Seasonality (written weights)")
    input_cell(ws, "C12", "OFF")
    note(ws, "D12", "ON applies tbl_Seasonality weights for the selected STATE; OFF = uniform writings (classic parallelogram).")
    label(ws, "B13", "Mod adjustment (master)")
    input_cell(ws, "C13", "ON")
    put(ws, "D13",
        "Set OFF if the indication's premium trend already reflects schedule mod drift — "
        "otherwise the drift is double-counted.",
        fnt=font(FAIL_RED, size=9, italic=True), align=ALIGN_WRAP)
    label(ws, "B14", "Projected-LR KPI shows")
    input_cell(ws, "C14", "As input", required=False)
    label(ws, "B15", f"Default net trend for CY {p + 1}")
    input_cell(ws, "C15", 0.0, fmt=FMT_PCT)
    note(ws, "D15",
         f"Annual net loss-over-premium trend applied once for the CY {p + 1} view wherever a "
         "combo's tbl_LR trend cell is blank; a per-state entry overrides this default.")
    ctx.define("nr_SeasonOn", "Control", "$C$12", "Global seasonality toggle (ON/OFF)")
    ctx.define("nr_ModAdjMaster", "Control", "$C$13", "Master mod-adjustment toggle (ON/OFF)")
    ctx.define("nr_BasisDisp", "Control", "$C$14", "Projected-LR KPI display basis")
    ctx.define("nr_TrendDefault", "Control", "$C$15",
               "Default net trend for the P+1 view (per-combo blank cells inherit it)")

    put(ws, "B16",
        '=IF(nr_SelOK,"","WARNING: the selected BU | state has no row in tbl_LR - engines show neutral factors until a row is added.")',
        fnt=font(FAIL_RED, bold=True))

    # KPI cards
    section(ws, 17, "B", "Key results — selected BU x state")
    cards = [
        ('=IF(nr_BasisDisp="As input","Projected LR (as input)","Projected LR (current level)")',
         '=IF(nr_BasisDisp="As input",nr_LRproj,nr_LRcur)', FMT_PCT,
         "from the rate level indication"),
        (f'="CY "&nr_PlanYear&" plan loss ratio"', "=nr_CYLR_P", FMT_PCT,
         "projected LR x A_rate x A_mod x A_other"),
        ('="CY earned rate chg vs indication"', "=nr_EChgVsInd", "+0.0%;-0.0%;0.0%",
         "E_CY(P) / CRL_ind - 1"),
        (f'="Carryover into "&(nr_PlanYear+1)', "=nr_YoY_P1", "+0.0%;-0.0%;0.0%",
         "E_CY(P+1) / E_CY(P) - 1"),
        ('="Written mod: current -> projected"',
         '=TEXT(nr_M0,"0.000")&"  ->  "&TEXT(nr_M1,"0.000")', FMT_GEN,
         "M_0 (as-of) -> M_1 (12/31/P)"),
        ('="Checks status"', "='Checks'!$C$3", FMT_GEN, "see the Checks sheet"),
    ]
    for i, (lbl_f, val_f, fmt, sub) in enumerate(cards):
        c1 = 2 + i * 2  # B, D, F, H, J, L
        put(ws, ws.cell(row=18, column=c1).coordinate, lbl_f, fnt=F_SMALL,
            fill=FILL_PANEL_2, align=ALIGN_L)
        put(ws, ws.cell(row=18, column=c1 + 1).coordinate, None, fill=FILL_PANEL_2)
        put(ws, ws.cell(row=19, column=c1).coordinate, val_f,
            fnt=font("1F3864", bold=True, size=16), fmt=fmt, fill=FILL_PANEL, align=ALIGN_L)
        put(ws, ws.cell(row=19, column=c1 + 1).coordinate, None, fill=FILL_PANEL)
        put(ws, ws.cell(row=20, column=c1).coordinate, sub, fnt=F_SMALL_IT, fill=FILL_PANEL)
        put(ws, ws.cell(row=20, column=c1 + 1).coordinate, None, fill=FILL_PANEL)
    ws.conditional_formatting.add(
        "L19",
        CellIsRule(operator="equal", formula=['"ALL CHECKS PASS"'], fill=FILL_GREEN,
                   font=font(PASS_GREEN, bold=True, size=14)))
    ws.conditional_formatting.add(
        "L19",
        CellIsRule(operator="notEqual", formula=['"ALL CHECKS PASS"'], fill=FILL_RED,
                   font=font(FAIL_RED, bold=True, size=14)))

    # Summary table P vs P+1
    section(ws, 22, "B", "Rate and price flow summary")
    put(ws, "B23", "Metric", fnt=F_HEADER, fill=FILL_NAVY)
    put(ws, "C23", '="CY "&nr_PlanYear', fnt=F_HEADER, fill=FILL_NAVY, align=ALIGN_C)
    put(ws, "D23", '="CY "&(nr_PlanYear+1)&" (indicative)"', fnt=F_HEADER, fill=FILL_NAVY, align=ALIGN_C)
    rows = [
        ("Earned rate index E_CY", "=nr_ECY_P", "=nr_ECY_P1", FMT_IDX),
        ("Rate adjustment A_rate", "=nr_Arate_P", "=nr_Arate_P1", FMT_IDX),
        ("Earned schedule mod", "=nr_MEarned_P", "=nr_MEarned_P1", FMT_MOD),
        ("Mod adjustment A_mod", "=nr_Amod_P", "=nr_Amod_P1", FMT_IDX),
        ("Earned rate chg (year over year)", "=nr_YoY_P", "=nr_YoY_P1", "+0.0%;-0.0%;0.0%"),
        ("CY plan loss ratio", "=nr_CYLR_P", "=nr_CYLR_P1", FMT_PCT),
    ]
    for i, (lbl_t, fp_, fp1, fmt) in enumerate(rows):
        r = 24 + i
        label(ws, f"B{r}", lbl_t)
        link(ws, f"C{r}", fp_, fmt=fmt, align=ALIGN_C, bold=(i == len(rows) - 1))
        link(ws, f"D{r}", fp1, fmt=fmt, align=ALIGN_C, bold=(i == len(rows) - 1))
    note(ws, "B30",
         f"The CY {p + 1} column is indicative only: it assumes no new indication and applies "
         "the net trend input once. Trend defaults to 0.0% — enter one in tbl_LR for a meaningful P+1 view.")

    # Scenario spotlight
    section(ws, 32, "B", "Scenario spotlight")
    label(ws, "B33", "Scenario in focus")
    link(ws, "C33", "=nr_SelScenario", align=ALIGN_C)
    label(ws, "B34", "CY plan LR under scenario")
    link(ws, "C34", "=INDEX(sc_res_cylr,MATCH(nr_SelScenario,sc_res_names,0))", fmt=FMT_PCT,
         align=ALIGN_C, bold=True)
    label(ws, "B35", "Delta vs Base")
    link(ws, "C35", "=(INDEX(sc_res_cylr,MATCH(nr_SelScenario,sc_res_names,0))-nr_CYLR_P)*100",
         fmt=FMT_PTS_SIGNED, align=ALIGN_C)
    note(ws, "B37", "Scenario definitions live on the Scenarios sheet; the Bridge and engine "
                    "sheets always show Base inputs (see Methodology).")

    _dv(ws, "list", ["C7"], formula1="=lst_bu")
    _dv(ws, "list", ["C8"], formula1="=lst_state")
    _dv(ws, "list", ["C9"], formula1="=lst_scenario")
    _dv(ws, "list", ["C12", "C13"], formula1="=lst_onoff")
    _dv(ws, "list", ["C14"], formula1="=lst_basisdisp")
    _dv(ws, "whole", ["C6"], operator="between", formula1="2000", formula2="2100",
        error="Plan year must be a 4-digit year.")
    _dv(ws, "decimal", ["C15"], operator="between", formula1="-0.5", formula2="1",
        error="Enter a decimal fraction between -50% and +100%.")

    set_widths(ws, {"A": 2, "B": 30, "C": 16, "D": 16, "E": 22, "F": 16, "G": 14,
                    "H": 16, "I": 14, "J": 16, "K": 14, "L": 16, "M": 14})
    presentation_setup(ws, gridlines_off=True, tab_color=NAVY)
    print_setup(ws)
