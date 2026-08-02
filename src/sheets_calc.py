"""_calc sheet: hidden engine blocks.

Sections, top to bottom:
  1. Results table (rows CALC_RES_FIRST..CALC_RES_LAST) — one row per tbl_LR
     row, read by Portfolio / State Summary / Program Flow. Columns A..AL are
     the D38 annual projection; AM.. are the D60 published program-flow
     columns (monthly legs + EP x weight products; see FLOW_PUB); CU..DH the
     book's rate-change history (BOOK_PUB); DI.. the D68 prior-year legs
     (PRIOR_PUB).
  2. Combo engine blocks (one per tbl_LR row), stride 54 rows.
  3. Scenario-transformed rate logs (4 x 9 columns) + 4 scenario blocks
     for the SELECTED combo.
  4. Attribution rank/transformed logs + 3 attribution blocks (magnitude,
     timing, actual-mod) for the selected combo.
  5. Solver base block (taken rows only) with cumulative columns for Mode B.
  6. Program-flow locked block (taken rows only, keyed on nr_SelKey, D59).

Every block is emitted by sheets_engine.write_cohort_block so the formula
pattern is identical to the visible Rate Engine. Dynamic anchors are stored in
ctx.lay_dyn for the visible sheets (Scenarios / Solver / Attribution) to link.
"""

from __future__ import annotations

from .build_workbook import Ctx, Layout as L, SHEETS, mod_adj_on
from .sheets_engine import (
    LogRefs, mod_drift_at_switch, write_cohort_block, write_mod_anchor_cells)
from .xlstyle import (FMT_IDX, FMT_INT, FMT_MOD, GREY_DARK, col, font, formula, label,
                      put, quote_sheet)

RL = quote_sheet(SHEETS.RATE_LOG)   # the rate-log rows the scenario/attr logs mirror

# D60: published program-flow columns on the results-table rows — the flat,
# SUMPRODUCT-able projection of each combo's monthly YoY legs. Column INDICES
# (render with xlstyle.col); monthly families span 12 columns, Jan..Dec P.
# Single source of truth: sheets_programflow and the harness import this map.
# Everything published here must stay STRICTLY NUMERIC (SUMPRODUCT *-form
# aggregation errors on text), so the mod leg is the raw path ratio and any
# gating happens at the aggregation site via the modeff column.
FLOW_PUB = dict(
    delivered=39,   # AM..AX  YoY delivered net leg by plan month
    rate=51,        # AY..BJ  YoY written rate leg
    mod=63,         # BK..BV  YoY written mod leg (raw path, ungated)
    epw=75,         # BW..CH  adjusted plan EP x seasonality weight w(m)
    modeff=87,      # CI      mod-adjustment-in-force echo (0/1)
    avg_rate=88,    # CJ      plan-year avg rate RATIO (block results echo)
    avg_mod=89,     # CK      plan-year avg mod RATIO
    avg_del=90,     # CL      plan-year avg delivered RATIO
    gap=91,         # CM      abs(avg delivered - 1 - net x); 0 for non-net
)

# A_other x EP, the last weighted product State Summary needed to show its
# bridge chain as visible arithmetic (D61). Appended AFTER the D60 flow block
# so no published column moves.
W_AOTHER_COL = 92   # CN

# Program-basis plan LR (D65): the explicit rate x mod counterfactual to a net
# selection, published so every exhibit reads it for one formula. Appended
# after CN so no earlier column moves.
PROG_LR = dict(
    cylr=93,      # CO  CY P plan LR on the explicit paths
    cylr1=94,     # CP  CY P+1
    w_cylr=95,    # CQ  x EP (aggregation)
    w_cylr1=96,   # CR  x EP
    ident=97,     # CS  |program - asserted| on NON-net rows; 0 on net rows
    gap=98,       # CT  signed program - asserted in points (0 when non-net)
)

# Rate-change history per combo (D66): the four chronological slots and the
# status counts, published so the BOOK workbook — which harvests this table and
# has no rate log of its own — can show the same program history the per-LOB
# State Summary shows. Slots are (date, effective %, T/P token) x 4.
BOOK_PUB = dict(
    ntaken=99,    # CU  count of taken rows for the combo
    nplanned=100,  # CV  count of planned rows
    slot=101,     # CW..DH  4 slots x 3 columns, chronological
)
BOOK_SLOTS = 4

# D68: the PRIOR year — the same three monthly legs plus their average, over
# Jan..Dec P-1 (the year currently flowing, whose year-ago base is P-2). The
# cohort blocks already span Jan P-2..Dec P+1, so this publishes rows that
# were always computed and never read. Appended after the book block so no
# published column moves and tools/harvest.py keeps every index it maps.
#
# The epw family is deliberately NOT duplicated: the seasonality weight is
# 12*INDEX(se_block,srow,MONTH(m))/ssum — a function of calendar MONTH alone —
# so w(Jan P-1) = w(Jan P) and the published plan-year weights aggregate the
# prior year correctly. Pinned by tests/test_program_flow.py
# (test_prior_weights_equal_plan_year_weights_by_calendar_month) and by the
# phase-C All-view tie, which would fail if the reuse were ever wrong.
PRIOR_PUB = dict(
    rate=113,       # DI..DT  YoY written rate leg, Jan..Dec P-1
    mod=125,        # DU..EF  YoY written mod leg (raw path, ungated)
    delivered=137,  # EG..ER  YoY delivered net leg
    avg_rate=149,   # ES      prior-year avg rate RATIO (block results echo)
    avg_mod=150,    # ET      prior-year avg mod RATIO
    avg_del=151,    # EU      prior-year avg delivered RATIO
)


def _grey(ws, addr):
    ws[addr].font = font(GREY_DARK, size=9)


def f_grey(ws, addr, f, fmt="General"):
    formula(ws, addr, f, fmt=fmt)
    _grey(ws, addr)


def build_calc(ctx: Ctx):
    ws = ctx.wb["_calc"]
    put(ws, "A1",
        "_calc (hidden): per-combo engine blocks generated by src/build_workbook.py. "
        "Same formula pattern as the visible Rate Engine; do not edit.",
        fnt=font(GREY_DARK, bold=True, size=9))

    # ------------------------------------------------------------------
    # 2. Combo engine blocks (written first so the results table can link)
    # ------------------------------------------------------------------
    block_tops = []
    for i in range(L.LR_ROWS):
        n = i + 1
        t = L.CALC_BLOCK_FIRST + i * L.CALC_BLOCK_STRIDE
        block_tops.append(t)
        # block header row t: key, state, T, seasonality helpers, guarded mod inputs
        f_grey(ws, f"A{t}", f"=INDEX(lr_key,{n})")
        f_grey(ws, f"B{t}", f"=INDEX(lr_state,{n})")
        f_grey(ws, f"C{t}", "=nr_TermMonths", FMT_INT)  # one LOB per workbook (D37)
        f_grey(ws, f"D{t}", f"=IF(COUNTIF(se_state,$B{t})=0,0,MATCH($B{t},se_state,0))", FMT_INT)
        f_grey(ws, f"E{t}", f"=IF($D{t}=0,0,INDEX(se_sum,$D{t}))", "0.00")
        f_grey(ws, f"F{t}", f"=IF(N(INDEX(lr_mind,{n}))=0,1,INDEX(lr_mind,{n}))", FMT_MOD)
        f_grey(ws, f"G{t}", f"=IF(N(INDEX(lr_m0,{n}))=0,1,INDEX(lr_m0,{n}))", FMT_MOD)
        f_grey(ws, f"H{t}",
               f"=IF(N(INDEX(lr_m0asof,{n}))=0,DATE(nr_PlanYear-1,10,1),INDEX(lr_m0asof,{n}))",
               "mm/dd/yyyy")
        f_grey(ws, f"I{t}", f"=IF(N(INDEX(lr_m1,{n}))=0,1,INDEX(lr_m1,{n}))", FMT_MOD)
        f_grey(ws, f"J{t}", f"=INDEX(lr_mprior,{n})", FMT_MOD)
        f_grey(ws, f"K{t}", f"=INDEX(lr_m2,{n})", FMT_MOD)
        # net-rate-selection cells (D39): mode, x(P), x(P+1), mod-adj-in-force
        f_grey(ws, f"L{t}", f"=NOT(ISBLANK(INDEX(lr_netp,{n})))")
        f_grey(ws, f"M{t}", f"=N(INDEX(lr_netp,{n}))", "0.0%")
        f_grey(ws, f"N{t}",
               f"=IF(ISBLANK(INDEX(lr_netp1,{n})),$M${t},N(INDEX(lr_netp1,{n})))", "0.0%")
        f_grey(ws, f"O{t}", f"=IF({mod_adj_on(n)},1,0)", FMT_INT)
        # D70 stepped-mod cells: how many actions this combo has logged, and
        # the level they compound on (blank M_endPrior falls back to M_0)
        f_grey(ws, f"P{t}", f"=COUNTIF(ml_key,$A${t})", FMT_INT)
        f_grey(ws, f"Q{t}",
               f"=IF(N(INDEX(lr_mendprior,{n}))=0,"
               + mod_drift_at_switch(f"$G${t}", f"$H${t}", f"$I${t}", f"$J${t}")
               + f",INDEX(lr_mendprior,{n}))", FMT_MOD)
        anchors = write_mod_anchor_cells(
            ws, ctx, t + 1, 1,
            mind_ref=f"$F${t}", m0_ref=f"$G${t}", asof_ref=f"$H${t}",
            m1_ref=f"$I${t}", mprior_ref=f"$J${t}", m2_ref=f"$K${t}",
            steps_cond=f"$P${t}>0", mend_ref=f"$Q${t}")
        blk = write_cohort_block(
            ws, ctx, t + 3, LogRefs(key_cond=f"$A${t}"), t_ref=f"$C${t}",
            srow_ref=f"$D${t}", ssum_ref=f"$E${t}", anchors=anchors,
            header=True, header_row_at=t + 2,
            net=dict(mode=f"$L${t}", x=f"$M${t}", x1=f"$N${t}",
                     modeff=f"$O${t}=1", mind=f"$F${t}"),
            mod_refs=LogRefs(key_cond=f"$A${t}", eff="ml_eff", ln="ml_ln1p",
                             effmonth="ml_effmonth", first="ml_first",
                             daysafter="ml_daysafter"),
            mod_col=22, mod_base_ref=f"$Q${t}", mod_count_ref=f"$P${t}",
            # D84: spare tbl_LR rows resolve to a blank key and can match no log
            # row, so their scans are pure cost. The gate returns exactly what
            # the sweep returns for an empty key.
            idle_cond=f'$A${t}=""')
        # block results row
        rr = t + 51
        label(ws, f"A{rr}", "block results:")
        _grey(ws, f"A{rr}")
        f_grey(ws, f"B{rr}", f'=EXP(SUMPRODUCT((rl_key=$A${t})*(rl_cons="Y")*rl_ln1p))', FMT_IDX)
        for cL, ec in (("C", blk["ec_pm1"]), ("D", blk["ec_p"]), ("E", blk["ec_p1"])):
            f_grey(ws, f"{cL}{rr}",
                   f"=SUMPRODUCT({blk['w']},{ec},{blk['U']})/SUMPRODUCT({blk['w']},{ec})",
                   FMT_IDX)
        for cL, ec in (("F", blk["ec_p"]), ("G", blk["ec_p1"])):
            f_grey(ws, f"{cL}{rr}",
                   f"=SUMPRODUCT({blk['w']},{ec},{blk['Mw']})/SUMPRODUCT({blk['w']},{ec})",
                   FMT_MOD)
        # earned rate index on the EXPLICIT written path (col N) rather than the
        # in-force path (col U): identical off net mode, the counterfactual on
        # it — the basis for the program-basis plan LR (D65)
        for cL, ec in (("K", blk["ec_p"]), ("L", blk["ec_p1"])):
            f_grey(ws, f"{cL}{rr}",
                   f"=SUMPRODUCT({blk['w']},{ec},{blk['W']})/SUMPRODUCT({blk['w']},{ec})",
                   FMT_IDX)
        # program-flow ratio averages (D59): w-weighted means of the monthly
        # YoY RATIOS over Jan..Dec P (vs year-ago) — NOT the E_CY aggregate-
        # ratio convention. Read by the Program Flow summary via stride-INDEX.
        wr = f"$H${t + 27}:$H${t + 38}"
        np_, ny = f"$N${t + 27}:$N${t + 38}", f"$N${t + 15}:$N${t + 26}"
        op, oy = f"$O${t + 27}:$O${t + 38}", f"$O${t + 15}:$O${t + 26}"
        f_grey(ws, f"H{rr}", f"=SUMPRODUCT({wr},{np_}/{ny})/SUM({wr})", FMT_IDX)
        f_grey(ws, f"I{rr}", f"=SUMPRODUCT({wr},{op}/{oy})/SUM({wr})", FMT_IDX)
        f_grey(ws, f"J{rr}",
               f"=IF($O${t}=1,SUMPRODUCT({wr},{np_}/{ny}*{op}/{oy})/SUM({wr}),$H{rr})",
               FMT_IDX)
        # D68: the same three averages one year back (Jan..Dec P-1 over its own
        # P-2 base) — the year currently flowing, shown beside the plan year.
        # M/N/O mirror H/I/J exactly; the results row sits outside every cohort
        # range so reusing those letters cannot collide with the block columns.
        wrp = f"$H${t + 15}:$H${t + 26}"
        npp, nyp = f"$N${t + 15}:$N${t + 26}", f"$N${t + 3}:$N${t + 14}"
        opp, oyp = f"$O${t + 15}:$O${t + 26}", f"$O${t + 3}:$O${t + 14}"
        f_grey(ws, f"M{rr}", f"=SUMPRODUCT({wrp},{npp}/{nyp})/SUM({wrp})", FMT_IDX)
        f_grey(ws, f"N{rr}", f"=SUMPRODUCT({wrp},{opp}/{oyp})/SUM({wrp})", FMT_IDX)
        f_grey(ws, f"O{rr}",
               f"=IF($O${t}=1,SUMPRODUCT({wrp},{npp}/{nyp}*{opp}/{oyp})/SUM({wrp}),"
               f"$M{rr})", FMT_IDX)

    # ------------------------------------------------------------------
    # 1. Results table rows 4..21 (linked by Portfolio via calc_* names)
    # ------------------------------------------------------------------
    hdr = ["key", "T", "CRL", "E_CY(P-1)", "E_CY(P)", "E_CY(P+1)", "Mbar(P)", "Mbar(P+1)",
           "A_rate(P)", "A_rate(P+1)", "A_mod(P)", "A_mod(P+1)", "LR_current",
           "CY LR(P)", "CY LR(P+1)", "EChg vs ind", "Carryover", "A_other", "trend"]
    for j, h in enumerate(hdr):
        put(ws, f"{col(1 + j)}{L.CALC_RES_HDR}", h, fnt=font(GREY_DARK, bold=True, size=9))
    for i in range(L.LR_ROWS):
        n = i + 1
        r = L.CALC_RES_FIRST + i
        t = block_tops[i]
        rr = t + 51
        f_grey(ws, f"A{r}", f"=$A${t}")
        f_grey(ws, f"B{r}", f"=$C${t}", FMT_INT)
        f_grey(ws, f"C{r}", f"=$B${rr}", FMT_IDX)
        f_grey(ws, f"D{r}", f"=$C${rr}", FMT_IDX)
        f_grey(ws, f"E{r}", f"=$D${rr}", FMT_IDX)
        f_grey(ws, f"F{r}", f"=$E${rr}", FMT_IDX)
        f_grey(ws, f"G{r}", f"=$F${rr}", FMT_MOD)
        f_grey(ws, f"H{r}", f"=$G${rr}", FMT_MOD)
        f_grey(ws, f"I{r}", f"=$C{r}/$E{r}", FMT_IDX)
        f_grey(ws, f"J{r}", f"=$C{r}/$F{r}", FMT_IDX)
        f_grey(ws, f"K{r}",
               f"=IF($L${t},1,IF({mod_adj_on(n)},$F${t}/$G{r},1))", FMT_IDX)
        f_grey(ws, f"L{r}",
               f"=IF($L${t},1,IF({mod_adj_on(n)},$F${t}/$H{r},1))", FMT_IDX)
        f_grey(ws, f"M{r}",
               f'=IF(INDEX(lr_basis,{n})="proposed",N(INDEX(lr_lrproj,{n}))*(1+N(INDEX(lr_s,{n}))),'
               f"N(INDEX(lr_lrproj,{n})))", "0.0%")
        f_grey(ws, f"R{r}", f"=IF(N(INDEX(lr_aother,{n}))=0,1,INDEX(lr_aother,{n}))", FMT_IDX)
        # effective trend: blank tbl_LR cell inherits the Control default (D38)
        f_grey(ws, f"S{r}",
               f"=IF(ISBLANK(INDEX(lr_trend,{n})),N(nr_TrendDefault),N(INDEX(lr_trend,{n})))",
               "0.0%")
        f_grey(ws, f"N{r}", f"=$M{r}*$I{r}*$K{r}*$R{r}", "0.0%")
        f_grey(ws, f"O{r}", f"=$M{r}*(1+$S{r})*$J{r}*$L{r}*$R{r}", "0.0%")
        f_grey(ws, f"P{r}", f"=$E{r}/$C{r}-1", "0.0%")
        f_grey(ws, f"Q{r}", f"=$F{r}/$E{r}-1", "0.0%")
        # State Summary support: dimension echoes, adjusted EP, and EP-weighted
        # numerators (metric x EP) so the exhibit aggregates via short SUMIFS (D38)
        f_grey(ws, f"T{r}", f"=INDEX(lr_state,{n})")
        f_grey(ws, f"U{r}", f"=INDEX(lr_bu,{n})")
        f_grey(ws, f"V{r}", f"=N(INDEX(lr_ep,{n}))", "#,##0")
        f_grey(ws, f"W{r}", f"=N(INDEX(lr_mind,{n}))*$V{r}", FMT_IDX)
        f_grey(ws, f"X{r}", f"=N(INDEX(lr_m0,{n}))*$V{r}", FMT_IDX)
        f_grey(ws, f"Y{r}", f"=N(INDEX(lr_m1,{n}))*$V{r}", FMT_IDX)
        f_grey(ws, f"Z{r}", f"=$C{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AA{r}", f"=$E{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AB{r}", f"=$I{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AC{r}", f"=$G{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AD{r}", f"=$K{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AE{r}", f"=$M{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AF{r}", f"=$N{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AG{r}", f"=$S{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AH{r}", f"=$J{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AI{r}", f"=$L{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AJ{r}", f"=$O{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"AK{r}", f"=IF($L${t},1,0)", FMT_INT)
        f_grey(ws, f"AL{r}", f"=IF($L${t},$M${t},0)", "0.0%")
        # D60 published program-flow columns (see FLOW_PUB above)
        for j in range(12):
            num, den = t + 27 + j, t + 15 + j
            f_grey(ws, f"{col(FLOW_PUB['delivered'] + j)}{r}",
                   f"=$N${num}/$N${den}*IF($O${t}=1,$O${num}/$O${den},1)-1", FMT_IDX)
            f_grey(ws, f"{col(FLOW_PUB['rate'] + j)}{r}",
                   f"=$N${num}/$N${den}-1", FMT_IDX)
            f_grey(ws, f"{col(FLOW_PUB['mod'] + j)}{r}",
                   f"=$O${num}/$O${den}-1", FMT_IDX)
            f_grey(ws, f"{col(FLOW_PUB['epw'] + j)}{r}", f"=$V{r}*$H${num}", FMT_IDX)
        # D68 published prior-year columns (see PRIOR_PUB above): identical
        # formulas one year back — Jan..Dec P-1 over its own P-2 base. No epw
        # twin; the plan-year weights above serve both years.
        for j in range(12):
            num, den = t + 15 + j, t + 3 + j
            f_grey(ws, f"{col(PRIOR_PUB['delivered'] + j)}{r}",
                   f"=$N${num}/$N${den}*IF($O${t}=1,$O${num}/$O${den},1)-1", FMT_IDX)
            f_grey(ws, f"{col(PRIOR_PUB['rate'] + j)}{r}",
                   f"=$N${num}/$N${den}-1", FMT_IDX)
            f_grey(ws, f"{col(PRIOR_PUB['mod'] + j)}{r}",
                   f"=$O${num}/$O${den}-1", FMT_IDX)
        f_grey(ws, f"{col(PRIOR_PUB['avg_rate'])}{r}", f"=$M${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(PRIOR_PUB['avg_mod'])}{r}", f"=$N${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(PRIOR_PUB['avg_del'])}{r}", f"=$O${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(FLOW_PUB['modeff'])}{r}", f"=$O${t}", FMT_INT)
        f_grey(ws, f"{col(FLOW_PUB['avg_rate'])}{r}", f"=$H${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(FLOW_PUB['avg_mod'])}{r}", f"=$I${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(FLOW_PUB['avg_del'])}{r}", f"=$J${t + 51}", FMT_IDX)
        f_grey(ws, f"{col(FLOW_PUB['gap'])}{r}",
               f"=IF($AK{r}=1,ABS(${col(FLOW_PUB['avg_del'])}{r}-1-$AL{r}),0)", FMT_IDX)
        f_grey(ws, f"{col(W_AOTHER_COL)}{r}", f"=$R{r}*$V{r}", FMT_IDX)
        # program-basis plan LR (D65): LR_current x (CRL / E_CY on the explicit
        # path) x the UNFORCED mod drift x A_other. Net mode forces A_mod to 1
        # and earns the asserted path, so these differ only where a net
        # selection is in force — pinned by the identity column below.
        amod_p = f"IF($O${t}=1,$F${t}/$G{r},1)"
        amod_p1 = f"IF($O${t}=1,$F${t}/$H{r},1)"
        cy, cy1 = col(PROG_LR["cylr"]), col(PROG_LR["cylr1"])
        f_grey(ws, f"{cy}{r}", f"=$M{r}*($C{r}/$K${rr})*{amod_p}*$R{r}", "0.0%")
        f_grey(ws, f"{cy1}{r}",
               f"=$M{r}*(1+$S{r})*($C{r}/$L${rr})*{amod_p1}*$R{r}", "0.0%")
        f_grey(ws, f"{col(PROG_LR['w_cylr'])}{r}", f"=${cy}{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"{col(PROG_LR['w_cylr1'])}{r}", f"=${cy1}{r}*$V{r}", FMT_IDX)
        f_grey(ws, f"{col(PROG_LR['ident'])}{r}",
               f"=IF($AK{r}=1,0,ABS(${cy}{r}-$N{r})+ABS(${cy1}{r}-$O{r}))", FMT_IDX)
        f_grey(ws, f"{col(PROG_LR['gap'])}{r}",
               f"=IF($AK{r}=0,0,(${cy}{r}-$N{r})*100)", "+0.00;-0.00;0.00")
        # rate-change history for the book harvest (D66)
        for cidx, status in ((BOOK_PUB["ntaken"], "taken"),
                             (BOOK_PUB["nplanned"], "planned")):
            f_grey(ws, f"{col(cidx)}{r}",
                   f'=COUNTIFS(rl_key,$A{r},rl_status,"{status}",rl_eff,"<>")', FMT_INT)
        for j in range(BOOK_SLOTS):
            c0 = BOOK_PUB["slot"] + j * 3
            cnt = f"COUNTIFS(rl_key,$A{r},rl_seq,{j + 1})"
            f_grey(ws, f"{col(c0)}{r}",
                   f'=IF({cnt}=0,"",SUMIFS(rl_eff,rl_key,$A{r},rl_seq,{j + 1}))',
                   "m/d/yy")
            f_grey(ws, f"{col(c0 + 1)}{r}",
                   f'=IF({cnt}=0,"",SUMIFS(rl_reff,rl_key,$A{r},rl_seq,{j + 1}))',
                   "+0.0%;-0.0%;0.0%")
            f_grey(ws, f"{col(c0 + 2)}{r}",
                   f'=IF({cnt}=0,"",'
                   f'IF(COUNTIFS(rl_key,$A{r},rl_seq,{j + 1},rl_status,"planned")>0,'
                   f'"P","T")&'
                   f'IF(COUNTIFS(rl_key,$A{r},rl_seq,{j + 1},rl_cons,"Y")>0,"","*"))')
    res_names = {
        "calc_key": ("A", "Combo keys, aligned 1:1 with tbl_LR rows"),
        "calc_term": ("B", "Policy term by combo"),
        "calc_crl": ("C", "CRL_ind by combo"),
        "calc_ecy_pm1": ("D", "E_CY(P-1) by combo"), "calc_ecy_p": ("E", "E_CY(P) by combo"),
        "calc_ecy_p1": ("F", "E_CY(P+1) by combo"),
        "calc_mbar_p": ("G", "CY P earned mod by combo"),
        "calc_mbar_p1": ("H", "CY P+1 earned mod by combo"),
        "calc_arate_p": ("I", "A_rate(P) by combo"), "calc_arate_p1": ("J", "A_rate(P+1) by combo"),
        "calc_amod_p": ("K", "A_mod(P) by combo"), "calc_amod_p1": ("L", "A_mod(P+1) by combo"),
        "calc_lrcur": ("M", "LR at current rate level by combo"),
        "calc_cylr_p": ("N", "CY plan LR(P) by combo"),
        "calc_cylr_p1": ("O", "CY LR(P+1) indicative by combo"),
        "calc_echg": ("P", "CY earned rate chg vs indication by combo"),
        "calc_carry": ("Q", "Carryover into P+1 by combo"),
        "calc_aother": ("R", "A_other by combo (1 when blank)"),
        "calc_trend": ("S", "Effective P+1 net trend by combo (blank inherits the default)"),
        "calc_state": ("T", "State echo, aligned with tbl_LR rows"),
        "calc_bu": ("U", "BU echo, aligned with tbl_LR rows"),
        "calc_ep": ("V", "Adjusted plan EP by combo (0 when blank)"),
        "calc_w_mind": ("W", "M_ind x EP (State Summary weighting)"),
        "calc_w_m0": ("X", "M_0 x EP"), "calc_w_m1": ("Y", "M_1 x EP"),
        "calc_w_crl": ("Z", "CRL_ind x EP"), "calc_w_ecy": ("AA", "E_CY(P) x EP"),
        "calc_w_arate": ("AB", "A_rate(P) x EP"), "calc_w_mbar": ("AC", "Earned mod(P) x EP"),
        "calc_w_amod": ("AD", "A_mod(P) x EP"), "calc_w_lrcur": ("AE", "LR_current x EP"),
        "calc_w_cylr": ("AF", "CY plan LR(P) x EP"), "calc_w_trend": ("AG", "Net trend x EP"),
        "calc_w_arate1": ("AH", "A_rate(P+1) x EP"), "calc_w_amod1": ("AI", "A_mod(P+1) x EP"),
        "calc_w_cylr1": ("AJ", "CY LR(P+1) x EP"),
        "calc_netmode": ("AK", "1 when the combo carries a net rate selection (D39)"),
        "calc_netx": ("AL", "Net selection x(P) by combo (0 when off)"),
    }
    for nm, key, desc in (
            ("calc_cylr_prog", "cylr",
             "CY P plan LR on the EXPLICIT rate x mod paths (D65)"),
            ("calc_cylr1_prog", "cylr1", "CY P+1 plan LR on the explicit paths"),
            ("calc_w_cylr_prog", "w_cylr", "Program-basis CY P plan LR x EP"),
            ("calc_w_cylr1_prog", "w_cylr1", "Program-basis CY P+1 plan LR x EP"),
            ("calc_progident", "ident",
             "Non-net rows: |program basis - headline| for both years; 0 on net rows"),
            ("calc_proggap", "gap",
             "Net rows: program basis minus the asserted plan LR, in points")):
        ctx.define(nm, "_calc",
                   f"${col(PROG_LR[key])}${L.CALC_RES_FIRST}:"
                   f"${col(PROG_LR[key])}${L.CALC_RES_LAST}", desc)
    for nm, key, desc in (
            ("calc_ntaken", "ntaken", "Count of TAKEN rate-log rows for the combo (D66)"),
            ("calc_nplanned", "nplanned", "Count of PLANNED rate-log rows for the combo")):
        ctx.define(nm, "_calc",
                   f"${col(BOOK_PUB[key])}${L.CALC_RES_FIRST}:"
                   f"${col(BOOK_PUB[key])}${L.CALC_RES_LAST}", desc)
    ctx.define("calc_w_aother", "_calc",
               f"${col(W_AOTHER_COL)}${L.CALC_RES_FIRST}:"
               f"${col(W_AOTHER_COL)}${L.CALC_RES_LAST}",
               "A_other x EP (State Summary bridge-chain weighting, D61)")
    ctx.define("pfd_bookgap", "_calc",
               f"${col(FLOW_PUB['gap'])}${L.CALC_RES_FIRST}:"
               f"${col(FLOW_PUB['gap'])}${L.CALC_RES_LAST}",
               "Per-combo abs gap: program-basis delivered avg vs the asserted net "
               "(0 = non-net) — book-wide, view-independent (D60)")
    for name, (cL, desc) in res_names.items():
        ctx.define(name, "_calc", f"${cL}${L.CALC_RES_FIRST}:${cL}${L.CALC_RES_LAST}", desc)

    # ------------------------------------------------------------------
    # 3. Scenario section (selected combo)
    # ------------------------------------------------------------------
    sl = block_tops[-1] + L.CALC_BLOCK_STRIDE + 2          # scenario log header row
    put(ws, f"A{sl}", "Scenario-transformed rate logs (S1..S4), row-aligned with tbl_RateLog",
        fnt=font(GREY_DARK, bold=True, size=9))
    sc_logs = []
    for s in range(1, 5):
        c0 = 1 + (s - 1) * 10  # A, K, U, AE
        cEff, cEm, cAch, cReff, cLn, cDay, cEcnt, cScnt, cFirst = (col(c0 + j) for j in range(9))
        for j, h in enumerate(["eff'", "effmo'", "ach'", "r_eff'", "ln'", "days'", "ecnt", "scnt",
                               "first'"]):
            put(ws, f"{col(c0 + j)}{sl + 1}", f"S{s} {h}", fnt=font(GREY_DARK, size=8))
        r0 = sl + 2
        for j in range(L.RL_ROWS):
            r = r0 + j
            ir = L.RL_FIRST + j
            f_grey(ws, f"{cEff}{r}",
                   f'=IF({RL}!$C${ir}="","",IF({RL}!$E${ir}="planned",'
                   f"EDATE({RL}!$C${ir},N(INDEX(sc_shift,{s}))),{RL}!$C${ir}))", "mm/dd/yyyy")
            f_grey(ws, f"{cEm}{r}",
                   f'=IF(${cEff}{r}="","",DATE(YEAR(${cEff}{r}),MONTH(${cEff}{r}),1))',
                   "mm/dd/yyyy")
            f_grey(ws, f"{cAch}{r}",
                   f'=IF(N(INDEX(sc_ach,{s}))=0,IF({RL}!$G${ir}="",1,{RL}!$G${ir}),'
                   f"INDEX(sc_ach,{s}))", "0.0%")
            f_grey(ws, f"{cReff}{r}",
                   f'=IF({RL}!$C${ir}="",0,IF({RL}!$E${ir}="planned",'
                   f"({RL}!$D${ir}+N(INDEX(sc_dpts,{s})))*${cAch}{r},{RL}!$D${ir}))", "0.0%")
            f_grey(ws, f"{cLn}{r}",
                   f'=IF({RL}!$C${ir}="",0,IF(1+${cReff}{r}>0,LN(1+${cReff}{r}),0))', FMT_IDX)
            f_grey(ws, f"{cDay}{r}",
                   f'=IF(${cEff}{r}="",0,EOMONTH(${cEff}{r},0)-${cEff}{r}+1)', FMT_INT)
            f_grey(ws, f"{cEcnt}{r}",
                   f'=IF(${cEff}{r}="",0,COUNTIFS(rl_key,{RL}!$J${ir},'
                   f'${cEm}${r0}:${cEm}${r0 + L.RL_ROWS - 1},${cEm}{r},'
                   f'${cEff}${r0}:${cEff}${r0 + L.RL_ROWS - 1},"<"&${cEff}{r}))', FMT_INT)
            f_grey(ws, f"{cScnt}{r}",
                   f'=IF(${cEff}{r}="",0,COUNTIFS({RL}!$J${L.RL_FIRST}:$J${ir},{RL}!$J${ir},'
                   f"${cEm}${r0}:${cEm}{r},${cEm}{r},${cEff}${r0}:${cEff}{r},${cEff}{r}))",
                   FMT_INT)
            f_grey(ws, f"{cFirst}{r}",
                   f'=IF(${cEff}{r}="",0,IF(AND(${cEcnt}{r}=0,${cScnt}{r}=1),1,0))', FMT_INT)
        sc_logs.append(dict(
            eff=f"${cEff}${r0}:${cEff}${r0 + L.RL_ROWS - 1}",
            effmonth=f"${cEm}${r0}:${cEm}${r0 + L.RL_ROWS - 1}",
            ln=f"${cLn}${r0}:${cLn}${r0 + L.RL_ROWS - 1}",
            daysafter=f"${cDay}${r0}:${cDay}${r0 + L.RL_ROWS - 1}",
            first=f"${cFirst}${r0}:${cFirst}${r0 + L.RL_ROWS - 1}",
        ))

    sb = sl + L.RL_ROWS + 4
    ctx.lay_dyn["scenario_results"] = []
    for s in range(1, 5):
        t = sb + (s - 1) * L.CALC_BLOCK_STRIDE
        put(ws, f"A{t}", f"Scenario S{s} engine block (selected combo)",
            fnt=font(GREY_DARK, bold=True, size=9))
        # D70: a scenario inherits the selected combo's mod actions unchanged —
        # the levers move the rate program and the mod LEVEL, never the log. The
        # delta-M_1 lever shifts the base those actions compound on, which keeps
        # "shift the projected mod path" meaning the same thing on both
        # mechanisms AND keeps the blank-lever run identical to the base. Without
        # this the Checks row "Scenario engine reproduces Base when all levers
        # blank" fails the moment a combo with logged actions is selected.
        f_grey(ws, f"M{t}", "='Mod Engine'!$C$11", FMT_INT)
        f_grey(ws, f"N{t}", f"='Mod Engine'!$C$12+N(INDEX(sc_dm1,{s}))", FMT_MOD)
        anchors = write_mod_anchor_cells(
            ws, ctx, t + 1, 1, mind_ref="nr_MInd", m0_ref="nr_M0", asof_ref="nr_M0Asof",
            m1_ref="nr_M1", mprior_ref="nr_MPrior", m2_ref="nr_M2",
            dm1_expr=f"+N(INDEX(sc_dm1,{s}))",
            steps_cond=f"$M${t}>0", mend_ref=f"$N${t}")
        lg = sc_logs[s - 1]
        blk = write_cohort_block(
            ws, ctx, t + 3,
            LogRefs(key_cond="nr_SelKey", eff=lg["eff"], ln=lg["ln"],
                    effmonth=lg["effmonth"], first=lg["first"], daysafter=lg["daysafter"]),
            t_ref="nr_TermMonths", srow_ref="re_srow", ssum_ref="re_ssum",
            anchors=anchors, header=True, header_row_at=t + 2,
            net=dict(mode="nr_NetMode",
                     x=f"(nr_NetSelP+N(INDEX(sc_dnet,{s})))",
                     x1=f"(nr_NetSelP1+N(INDEX(sc_dnet,{s})))",
                     modeff='nr_ModAdjEff="ON"', mind="nr_MInd"),
            mod_refs=LogRefs(key_cond="nr_SelKey", eff="ml_eff", ln="ml_ln1p",
                             effmonth="ml_effmonth", first="ml_first",
                             daysafter="ml_daysafter"),
            mod_col=22, mod_base_ref=f"$N${t}", mod_count_ref=f"$M${t}")
        rr = t + 51
        label(ws, f"A{rr}", f"S{s} results:")
        _grey(ws, f"A{rr}")
        f_grey(ws, f"B{rr}",
               f"=SUMPRODUCT({blk['w']},{blk['ec_p']},{blk['U']})/SUMPRODUCT({blk['w']},{blk['ec_p']})",
               FMT_IDX)
        f_grey(ws, f"C{rr}",
               f"=SUMPRODUCT({blk['w']},{blk['ec_p1']},{blk['U']})/SUMPRODUCT({blk['w']},{blk['ec_p1']})",
               FMT_IDX)
        f_grey(ws, f"D{rr}",
               f"=SUMPRODUCT({blk['w']},{blk['ec_p']},{blk['Mw']})/SUMPRODUCT({blk['w']},{blk['ec_p']})",
               FMT_MOD)
        ctx.lay_dyn["scenario_results"].append(
            dict(e_p=f"'_calc'!$B${rr}", e_p1=f"'_calc'!$C${rr}", mbar_p=f"'_calc'!$D${rr}"))

    # ------------------------------------------------------------------
    # 4. Attribution section (selected combo)
    # ------------------------------------------------------------------
    al = sb + 4 * L.CALC_BLOCK_STRIDE + 2
    put(ws, f"A{al}", "Attribution: planned-row ranks and transformed logs (selected combo)",
        fnt=font(GREY_DARK, bold=True, size=9))
    for j, h in enumerate(["rank", "r_eff mag", "ln mag", "eff act", "effmo act", "days act",
                           "ecnt", "scnt", "first act"]):
        put(ws, f"{col(1 + j)}{al + 1}", h, fnt=font(GREY_DARK, size=8))
    r0 = al + 2
    for j in range(L.RL_ROWS):
        r = r0 + j
        ir = L.RL_FIRST + j
        f_grey(ws, f"A{r}",
               f'=IF(OR({RL}!$J${ir}<>nr_SelKey,{RL}!$E${ir}<>"planned",{RL}!$C${ir}=""),0,'
               f'COUNTIFS({RL}!$J${L.RL_FIRST}:$J${ir},nr_SelKey,'
               f'{RL}!$E${L.RL_FIRST}:$E${ir},"planned"))', FMT_INT)
        f_grey(ws, f"B{r}",
               f"=IF($A{r}=0,N({RL}!$K${ir}),IF(ISBLANK(INDEX(att_actpct,$A{r})),"
               f"N({RL}!$K${ir}),INDEX(att_actpct,$A{r})))", "0.0%")
        f_grey(ws, f"C{r}",
               f'=IF({RL}!$C${ir}="",0,IF(1+$B{r}>0,LN(1+$B{r}),0))', FMT_IDX)
        f_grey(ws, f"D{r}",
               f'=IF({RL}!$C${ir}="","",IF($A{r}=0,{RL}!$C${ir},'
               f"IF(ISBLANK(INDEX(att_actdate,$A{r})),{RL}!$C${ir},INDEX(att_actdate,$A{r}))))",
               "mm/dd/yyyy")
        f_grey(ws, f"E{r}", f'=IF($D{r}="","",DATE(YEAR($D{r}),MONTH($D{r}),1))', "mm/dd/yyyy")
        f_grey(ws, f"F{r}", f'=IF($D{r}="",0,EOMONTH($D{r},0)-$D{r}+1)', FMT_INT)
        f_grey(ws, f"G{r}",
               f'=IF($D{r}="",0,COUNTIFS(rl_key,{RL}!$J${ir},'
               f'$E${r0}:$E${r0 + L.RL_ROWS - 1},$E{r},$D${r0}:$D${r0 + L.RL_ROWS - 1},"<"&$D{r}))', FMT_INT)
        f_grey(ws, f"H{r}",
               f'=IF($D{r}="",0,COUNTIFS({RL}!$J${L.RL_FIRST}:$J${ir},{RL}!$J${ir},'
               f"$E${r0}:$E{r},$E{r},$D${r0}:$D{r},$D{r}))", FMT_INT)
        f_grey(ws, f"I{r}", f'=IF($D{r}="",0,IF(AND($G{r}=0,$H{r}=1),1,0))', FMT_INT)
    ctx.lay_dyn["att_rank"] = f"'_calc'!$A${r0}:$A${r0 + L.RL_ROWS - 1}"
    mag_ln = f"$C${r0}:$C${r0 + L.RL_ROWS - 1}"
    act_eff = f"$D${r0}:$D${r0 + L.RL_ROWS - 1}"
    act_em = f"$E${r0}:$E${r0 + L.RL_ROWS - 1}"
    act_day = f"$F${r0}:$F${r0 + L.RL_ROWS - 1}"
    act_first = f"$I${r0}:$I${r0 + L.RL_ROWS - 1}"

    ab = al + L.RL_ROWS + 4
    # magnitude block: achieved % at planned dates (rate only)
    t = ab
    put(ws, f"A{t}", "Attribution block 1 — rate magnitude (achieved % at planned dates)",
        fnt=font(GREY_DARK, bold=True, size=9))
    blk = write_cohort_block(
        ws, ctx, t + 3, LogRefs(key_cond="nr_SelKey", ln=mag_ln),
        t_ref="nr_TermMonths", srow_ref="re_srow", ssum_ref="re_ssum",
        anchors=None, include_mod=False, header=True, header_row_at=t + 2)
    rr = t + 51
    f_grey(ws, f"B{rr}",
           f"=SUMPRODUCT({blk['w']},{blk['ec_p']},{blk['W']})/SUMPRODUCT({blk['w']},{blk['ec_p']})",
           FMT_IDX)
    ctx.lay_dyn["att_e_mag"] = f"'_calc'!$B${rr}"

    # timing block: achieved % at actual dates
    t = ab + L.CALC_BLOCK_STRIDE
    put(ws, f"A{t}", "Attribution block 2 — rate timing (achieved % at actual dates)",
        fnt=font(GREY_DARK, bold=True, size=9))
    blk = write_cohort_block(
        ws, ctx, t + 3,
        LogRefs(key_cond="nr_SelKey", eff=act_eff, ln=mag_ln, effmonth=act_em,
                first=act_first, daysafter=act_day),
        t_ref="nr_TermMonths", srow_ref="re_srow", ssum_ref="re_ssum",
        anchors=None, include_mod=False, header=True, header_row_at=t + 2)
    rr = t + 51
    f_grey(ws, f"B{rr}",
           f"=SUMPRODUCT({blk['w']},{blk['ec_p']},{blk['W']})/SUMPRODUCT({blk['w']},{blk['ec_p']})",
           FMT_IDX)
    ctx.lay_dyn["att_e_time"] = f"'_calc'!$B${rr}"

    # actual-mod block
    t = ab + 2 * L.CALC_BLOCK_STRIDE
    put(ws, f"A{t}", "Attribution block 3 — actual written-mod path",
        fnt=font(GREY_DARK, bold=True, size=9))
    f_grey(ws, f"M{t}", "=IF(N(att_m0)=0,nr_M0,att_m0)", FMT_MOD)
    f_grey(ws, f"N{t}", "=IF(N(att_asof)=0,nr_M0Asof,att_asof)", "mm/dd/yyyy")
    f_grey(ws, f"O{t}", "=IF(N(att_m1)=0,nr_M1,att_m1)", FMT_MOD)
    f_grey(ws, f"P{t}", "=IF(N(att_m2)=0,N(nr_M2),att_m2)", FMT_MOD)
    # D70: taken mod actions ARE the actual path, so the actual-mod block reads
    # the log too. With no att_* overrides typed this reproduces the base path
    # exactly and the mod attribution factor is 1 — the property attribution has
    # always had, restored. (Restating a mod action as partly achieved is not yet
    # modelled; there is no actual-mod override to match att_actpct on rates.)
    f_grey(ws, f"Q{t}", "='Mod Engine'!$C$11", FMT_INT)
    f_grey(ws, f"R{t}", "='Mod Engine'!$C$12", FMT_MOD)
    anchors = write_mod_anchor_cells(
        ws, ctx, t + 1, 1, mind_ref="nr_MInd", m0_ref=f"$M${t}", asof_ref=f"$N${t}",
        m1_ref=f"$O${t}", mprior_ref="nr_MPrior", m2_ref=f"$P${t}",
        steps_cond=f"$Q${t}>0", mend_ref=f"$R${t}")
    blk = write_cohort_block(
        ws, ctx, t + 3, LogRefs(key_cond="nr_SelKey"), t_ref="nr_TermMonths",
        srow_ref="re_srow", ssum_ref="re_ssum", anchors=anchors,
        include_rate=False, header=True, header_row_at=t + 2,
        mod_refs=LogRefs(key_cond="nr_SelKey", eff="ml_eff", ln="ml_ln1p",
                         effmonth="ml_effmonth", first="ml_first",
                         daysafter="ml_daysafter"),
        mod_col=22, mod_base_ref=f"$R${t}", mod_count_ref=f"$Q${t}")
    rr = t + 51
    f_grey(ws, f"B{rr}",
           f"=SUMPRODUCT({blk['w']},{blk['ec_p']},{blk['Mw']})/SUMPRODUCT({blk['w']},{blk['ec_p']})",
           FMT_MOD)
    ctx.lay_dyn["att_mbar_act"] = f"'_calc'!$B${rr}"

    # ------------------------------------------------------------------
    # 5. Solver base block (taken rows only) + cumulative columns
    # ------------------------------------------------------------------
    t = ab + 3 * L.CALC_BLOCK_STRIDE
    put(ws, f"A{t}", "Solver base block — taken rows only (planned program excluded, D13); "
                     "keyed on slv_key so the Solver can follow Control OR its own override",
        fnt=font(GREY_DARK, bold=True, size=9))
    # local seasonality for the SOLVE combo's state (may differ from the selection)
    f_grey(ws, f"M{t}", "=IF(COUNTIF(se_state,slv_state)=0,0,MATCH(slv_state,se_state,0))",
           FMT_INT)
    f_grey(ws, f"N{t}", f"=IF($M${t}=0,0,INDEX(se_sum,$M${t}))", "0.00")
    # first="rl_firsttaken": the split flag must be STATUS-AWARE — the plain
    # rl_first is filter-blind, and a planned row earlier in a taken row's
    # cohort month would steal the flag, zero the L column, and silently drop
    # the taken change from that cohort's W0 (D58; oracle filters first).
    # D73: the block carries a MOD leg too, so the same block serves Mode A/B
    # (solve on rate) and Mode C (solve on a dated mod step).
    #   V = mod-action count for the solve combo, +1 for the step being solved
    #       for. The count drives the re-anchor, and the answer will always
    #       contain at least one action, so the base must already be
    #       re-anchored — the workbook form of the oracle's zero-step base.
    #   W = M_endPrior for the solve combo (M_0 when blank)
    # every mod input resolved for slv_key, NOT the Control selection — the
    # Solver's override can detach the two, and reading nr_M0 here would
    # quietly solve the wrong combo
    # D85: read the SOLVE combo's own _calc header cells rather than
    # re-resolving the mod inputs from tbl_LR. This block used to be a fourth
    # hand-synced copy of that resolution and it had dropped the blank guards
    # every other copy carries — so a combo with blank mod fields (legal
    # everywhere else, where it reads as a flat 1.000 path) drove Mode C's mod
    # leg off M_0 = 0 and an as-of date of serial 0, silently. Pointing at the
    # guarded cells fixes the divergence and removes the copy in one edit.
    sr = "MATCH(slv_key,lr_key,0)"
    cl = L.CALC_BLOCK_FIRST + L.LR_ROWS * L.CALC_BLOCK_STRIDE + 60
    base = f"{L.CALC_BLOCK_FIRST}+({sr}-1)*{L.CALC_BLOCK_STRIDE}"

    def hdr(colL: str, absent: str) -> str:
        return (f"=IF(COUNTIF(lr_key,slv_key)=0,{absent},"
                f"INDEX(${colL}$1:${colL}${cl},{base}))")

    f_grey(ws, f"V{t}", "=COUNTIF(ml_key,slv_key)+1", FMT_INT)
    f_grey(ws, f"W{t}", hdr("Q", "1"), FMT_MOD)          # guarded M_endPrior
    f_grey(ws, f"X{t}", hdr("G", "1"), FMT_MOD)          # guarded M_0
    f_grey(ws, f"Y{t}", hdr("H", "DATE(nr_PlanYear-1,10,1)"), "mm/dd/yyyy")
    f_grey(ws, f"Z{t}", hdr("J", "0"), FMT_MOD)          # M_prior (raw, as _calc)
    anchors_s = write_mod_anchor_cells(
        ws, ctx, t + 1, 1,
        mind_ref="nr_MInd", m0_ref=f"$X${t}", asof_ref=f"$Y${t}",
        m1_ref=f"$W${t}", mprior_ref=f"$Z${t}", m2_ref=f"$W${t}",
        steps_cond="TRUE", mend_ref=f"$W${t}")
    blk = write_cohort_block(
        ws, ctx, t + 3, LogRefs(key_cond="slv_key", taken_only=True,
                                first="rl_firsttaken"),
        t_ref="nr_TermMonths", srow_ref=f"$M${t}", ssum_ref=f"$N${t}",
        anchors=anchors_s, header=True, header_row_at=t + 2,
        mod_refs=LogRefs(key_cond="slv_key", eff="ml_eff", ln="ml_ln1p",
                         effmonth="ml_effmonth", first="ml_firsttaken",
                         daysafter="ml_daysafter", taken_only=True,
                         status="ml_status"),
        mod_col=27, mod_base_ref=f"$W${t}", mod_count_ref=f"$V${t}")
    bf, bl = blk["first"], blk["last"]
    for cL, hdr_ in (("S", "w-ec-W0"), ("T", "cum"), ("U", "w-ec-W0eom"),
                     ("AG", "w-ec-M0"), ("AH", "cum mod"), ("AI", "w-ec-Meom")):
        put(ws, f"{cL}{t + 2}", hdr_, fnt=font(GREY_DARK, size=8))
    mc = [col(27 + i) for i in range(5)]     # AA..AE: the mod index helper strip
    for i in range(L.N_COH):
        r = bf + i
        f_grey(ws, f"S{r}", f"=$H{r}*$Q{r}*$N{r}", FMT_IDX)
        f_grey(ws, f"T{r}", f"=SUM($S${bf}:$S{r})", FMT_IDX)
        f_grey(ws, f"U{r}", f"=$H{r}*$Q{r}*$J{r}", FMT_IDX)
        # mod analogues: level in force, running total, and the end-of-month
        # mod the split month's post-share is valued at
        f_grey(ws, f"AG{r}", f"=$H{r}*$Q{r}*$O{r}", FMT_IDX)
        f_grey(ws, f"AH{r}", f"=SUM($AG${bf}:$AG{r})", FMT_IDX)
        f_grey(ws, f"AI{r}", f"=$H{r}*$Q{r}*$W${t}*${mc[1]}{r}", FMT_IDX)
    rr = t + 51
    label(ws, f"A{rr}", "solver totals:")
    _grey(ws, f"A{rr}")
    f_grey(ws, f"B{rr}", f"=SUMPRODUCT({blk['w']},{blk['ec_p']})", FMT_IDX)   # D
    f_grey(ws, f"C{rr}", f"=SUM($S${bf}:$S${bl})", FMT_IDX)                   # total blend
    f_grey(ws, f"D{rr}", f"=SUM($AG${bf}:$AG${bl})", FMT_IDX)                 # total mod
    ctx.lay_dyn["solver"] = dict(first=bf, last=bl)
    slv_names = {
        "slv_absmi": (f"$G${bf}:$G${bl}", "Solver block: absolute cohort month indices"),
        "slv_w": (f"$H${bf}:$H${bl}", "Solver block: written weights"),
        "slv_ecp": (f"$Q${bf}:$Q${bl}", "Solver block: CY P earning fractions"),
        "slv_widx": (f"$N${bf}:$N${bl}", "Solver block: taken-only written index W0 (blended)"),
        "slv_wpre": (f"$I${bf}:$I${bl}", "Solver block: pre-month index W0_pre"),
        "slv_weom": (f"$J${bf}:$J${bl}", "Solver block: end-of-month index W0_eom"),
        "slv_dfirst": (f"$L${bf}:$L${bl}", "Solver block: first-taken-change days on/after"),
        "slv_cum": (f"$T${bf}:$T${bl}", "Solver block: cumulative w-ec-W0"),
        "slv_weom_c": (f"$U${bf}:$U${bl}", "Solver block: w-ec-W0_eom by cohort"),
        "slv_den": (f"$B${rr}", "Solver denominator D = SUM w-ec over CY P"),
        "slv_total": (f"$C${rr}", "Solver total SUM w-ec-W0 over CY P"),
        # D73 mod leg
        "slv_modcum": (f"$AH${bf}:$AH${bl}", "Solver block: cumulative w-ec-M_w (taken "
                                             "mod actions only)"),
        "slv_meom_c": (f"$AI${bf}:$AI${bl}", "Solver block: w-ec-M_eom by cohort"),
        "slv_mtotal": (f"$D${rr}", "Solver total SUM w-ec-M_w over CY P"),
        "slv_mbase": (f"$W${t}", "Solver block: M_endPrior the mod steps compound on"),
    }
    for name, (ref, desc) in slv_names.items():
        ctx.define(name, "_calc", ref, desc)

    # ------------------------------------------------------------------
    # 6. Program-flow locked block (taken rows only, keyed on nr_SelKey)
    # ------------------------------------------------------------------
    t = ab + 4 * L.CALC_BLOCK_STRIDE
    put(ws, f"A{t}", "Program-flow locked block — taken rows only (D13), keyed on "
                     "nr_SelKey (the Solver base cannot be reused: slv_key can detach "
                     "from Control via the override); status-aware first flag per D58",
        fnt=font(GREY_DARK, bold=True, size=9))
    # D77: the mod leg gets the same treatment. The step-index columns filter
    # to TAKEN mod rows, but the regime switch and the level they compound on
    # follow the log's PRESENCE — 'Mod Engine'!$C$11 counts every row. A combo
    # whose log is all-planned therefore stays re-anchored (D70) with a step
    # index of 1.000, which is the honest locked path: nothing done yet, on
    # the level the year actually starts from. Filtering the switch too would
    # put the two legs on different year-ago bases and the planned residual
    # would silently report the re-anchor.
    anchors_tk = write_mod_anchor_cells(
        ws, ctx, t + 1, 1, mind_ref="nr_MInd", m0_ref="nr_M0", asof_ref="nr_M0Asof",
        m1_ref="nr_M1", mprior_ref="nr_MPrior", m2_ref="nr_M2",
        steps_cond="'Mod Engine'!$C$11>0", mend_ref="'Mod Engine'!$C$12")
    blk = write_cohort_block(
        ws, ctx, t + 3, LogRefs(key_cond="nr_SelKey", taken_only=True,
                                first="rl_firsttaken"),
        t_ref="nr_TermMonths", srow_ref="re_srow", ssum_ref="re_ssum",
        anchors=anchors_tk, header=True, header_row_at=t + 2,
        mod_refs=LogRefs(key_cond="nr_SelKey", eff="ml_eff", ln="ml_ln1p",
                         effmonth="ml_effmonth", first="ml_firsttaken",
                         daysafter="ml_daysafter", taken_only=True,
                         status="ml_status"),
        mod_col=22, mod_base_ref="'Mod Engine'!$C$12",
        mod_count_ref="'Mod Engine'!$C$11")
    ctx.define("fd_tkw", "_calc", f"$N${blk['first']}:$N${blk['last']}",
               "Taken-only written index for the Control selection (Flow Dashboard "
               "locked leg, D59)")
    ctx.define("fd_tkm", "_calc", f"$O${blk['first']}:$O${blk['last']}",
               "Taken-only written MOD for the Control selection (Flow Dashboard "
               "locked pricing leg, D77)")
    ctx.lay_dyn["progflow_tk"] = dict(top=t, first=blk["first"], last=blk["last"])
