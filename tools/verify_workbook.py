"""Workbook verification harness (brief §9 "Workbook verification").

Dimensioning (D37): one workbook per LOB; combos are BU x state. The harness
verifies one LOB workbook per invocation (--lob, default Property):

  A. Static scans of the formula layer: no merged cells, no volatile /
     dynamic-array / prohibited functions, no external references, defined
     names intact.
  B. Recalculate (Excel COM / LibreOffice) and prove zero formula errors,
     plus the chart-axis regression guard (D36).
  C. Oracle ties in the default state: worked-example Bridge cells, monthly
     series, EVERY BU x state combo, solver, scenario-blank identity.
  D. Toggle exercises: mutate inputs on scratch copies, recalculate, re-read,
     tie each state to a fresh oracle run, rescan for errors. (Runs a
     representative selector subsample; phase C already ties all combos.)

Usage:
    python tools/verify_workbook.py [--lob "Property"] [--workbook PATH] [--quick]

Exit code 0 = every assertion passed.
"""

from __future__ import annotations

import argparse
import datetime as dt
import math
import re
import shutil
import sys
from dataclasses import replace
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

from src import engine  # noqa: E402
from src.build_workbook import (  # noqa: E402
    Layout as L, degenerate_combo, load_config, output_path, sample_lr_rows,
    sample_rate_rows, sample_seasonality_rows, sample_to_combo,
)
from tools.recalc import recalc  # noqa: E402

ERR_STRINGS = ("#DIV/0!", "#REF!", "#N/A", "#NAME?", "#VALUE!", "#NUM!", "#NULL!",
               "#SPILL!", "#CALC!", "#GETTING_DATA")
BANNED_FUNCS = [
    "OFFSET", "INDIRECT", "TODAY", "NOW", "RAND", "RANDBETWEEN", "RANDARRAY",
    "XLOOKUP", "XMATCH", "FILTER", "SORTBY", "UNIQUE", "SEQUENCE", "LET",
    "LAMBDA", "IFERROR", "IFNA", "TEXTJOIN", "IFS", "SWITCH", "MAXIFS",
    "MINIFS", "CONCAT", "TOCOL", "TOROW", "VSTACK", "HSTACK", "TAKE", "DROP",
]
# word-boundary match so COUNTIFS/SUMIFS don't false-positive on IFS, etc.
BANNED_RE = re.compile(r"(?<![A-Z0-9_.])(" + "|".join(BANNED_FUNCS) + r")\(")

PASS = 0
FAIL = 0
FAILURES: list[str] = []


def check(desc: str, ok: bool, detail: str = ""):
    global PASS, FAIL
    if ok:
        PASS += 1
    else:
        FAIL += 1
        FAILURES.append(f"{desc}  {detail}")
        print(f"  FAIL  {desc}  {detail}")


def approx(a, b, tol=1e-9):
    if a is None or b is None:
        return False
    if isinstance(a, str) or isinstance(b, str):
        return False
    return abs(float(a) - float(b)) <= tol


def name_cell(wb, name: str):
    dn = wb.defined_names[name]
    for sheet, ref in dn.destinations:
        return wb[sheet][ref.replace("$", "")]
    raise KeyError(name)


def nval(wb, name: str):
    return name_cell(wb, name).value


def scan_errors(wb) -> list[str]:
    out = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                if isinstance(c.value, str) and c.value in ERR_STRINGS:
                    out.append(f"{ws.title}!{c.coordinate}={c.value}")
    return out


# ---------------------------------------------------------------------------
# Phase A: static scans
# ---------------------------------------------------------------------------


def phase_a(path: Path):
    print("Phase A: static formula-layer scans")
    wb = openpyxl.load_workbook(path, data_only=False)
    merged = {ws.title: len(ws.merged_cells.ranges) for ws in wb.worksheets
              if len(ws.merged_cells.ranges)}
    check("no merged cells anywhere", not merged, str(merged))

    banned_hits, ext_hits, longest = [], [], 0
    n_formulas = 0
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    n_formulas += 1
                    longest = max(longest, len(v))
                    for hit in BANNED_RE.findall(v.upper()):
                        banned_hits.append(f"{ws.title}!{c.coordinate}: {hit}(")
                    if "[" in v and "]" in v:
                        ext_hits.append(f"{ws.title}!{c.coordinate}")
    check("no volatile / dynamic-array / prohibited functions", not banned_hits,
          "; ".join(banned_hits[:5]))
    check("no external-workbook references", not ext_hits, "; ".join(ext_hits[:5]))
    check("defined names all present and unbroken",
          all("#REF" not in (wb.defined_names[n].attr_text or "")
              for n in wb.defined_names) and len(list(wb.defined_names)) > 80,
          f"{len(list(wb.defined_names))} names")
    print(f"  formulas: {n_formulas:,}; longest: {longest} chars")
    check("automatic calculation mode",
          wb.calculation.calcMode in (None, "auto"), str(wb.calculation.calcMode))
    return wb


# ---------------------------------------------------------------------------
# Phase B/C: recalc + oracle ties in the default state
# ---------------------------------------------------------------------------


def tie_default_state(path: Path, cfg, lob, do_recalc=True):
    print("Phase B: recalculate and scan for errors")
    if do_recalc:
        used = recalc(path)
        print(f"  recalculated with {used}")
    wb = openpyxl.load_workbook(path, data_only=True)
    errs = scan_errors(wb)
    check("zero formula-error cells in every sheet (incl. hidden)", not errs,
          "; ".join(errs[:8]))
    check("Checks sheet: ALL CHECKS PASS", nval(wb, "ck_overall") == "ALL CHECKS PASS",
          str(nval(wb, "ck_overall")))

    # Chart axes must stay visible after the Excel resave (D36).
    import zipfile
    deleted = 0
    total_axes = 0
    with zipfile.ZipFile(path) as z:
        for name in z.namelist():
            if re.match(r"xl/charts/chart\d+\.xml$", name):
                xml = z.read(name).decode("utf-8", errors="replace")
                total_axes += len(re.findall(r"<c:(catAx|valAx|dateAx)>", xml))
                deleted += len(re.findall(r'<c:delete val="1"\s*/>', xml))
    check("chart axes visible after Excel resave (1 intentional hide only)",
          total_axes >= 16 and deleted == 1,
          f"{deleted} deleted of {total_axes} axes")

    print("Phase C: oracle ties (default state)")
    p = cfg.plan_year
    lr_rows = sample_lr_rows(cfg, lob)
    rate_rows = sample_rate_rows(cfg, lob)
    we = next(r for r in lr_rows if r["bu"] == "BU-A" and r["state"] == cfg.states[0])
    combo = sample_to_combo(cfg, lob, we, rate_rows)
    m = engine.run_bridge(p, combo, "monthly")
    c = engine.run_bridge(p, combo, "continuous")

    ties = [
        ("nr_CRLind", m.crl_ind), ("nr_ECY_Pm1", m.e_cy[p - 1]), ("nr_ECY_P", m.e_cy[p]),
        ("nr_ECY_P1", m.e_cy[p + 1]), ("nr_Arate_P", m.a_rate_p),
        ("nr_Arate_P1", m.a_rate_p1), ("nr_MEarned_P", m.m_earned[p]),
        ("nr_MEarned_P1", m.m_earned[p + 1]), ("nr_Amod_P", m.a_mod_p),
        ("nr_Amod_P1", m.a_mod_p1), ("nr_LRcur", m.lr_current), ("nr_CYLR_P", m.cy_lr_p),
        ("nr_CYLR_P1", m.cy_lr_p1), ("nr_EChgVsInd", m.earned_chg_vs_ind),
        ("nr_YoY_P", m.yoy_earned_p), ("nr_YoY_P1", m.yoy_earned_p1),
    ]
    for name, expected in ties:
        check(f"{name} ties oracle monthly", approx(nval(wb, name), expected, 1e-9),
              f"wb={nval(wb, name)} oracle={expected}")
    check("CY LR within 0.10 pts of continuous closed form",
          approx(nval(wb, "nr_CYLR_P"), c.cy_lr_p, 1e-3))
    if lob.term_months == 12 and p == 2027:
        check("worked example CY LR rounds to 63.36%",
              round(float(nval(wb, "nr_CYLR_P")), 4) == 0.6336, str(nval(wb, "nr_CYLR_P")))

    # monthly series: written index, written mod, earned index by month
    re_ws, me_ws = wb["Rate Engine"], wb["Mod Engine"]
    ok_w = ok_m = ok_e = True
    for i in range(L.N_COH):
        r = L.RE_COH_FIRST + i
        ok_w &= approx(re_ws[f"N{r}"].value, m.written_index[i], 1e-9)
        ok_m &= approx(me_ws[f"E{r}"].value, m.written_mod[i], 1e-9)
    for j in range(L.N_MONTHS):
        cell = re_ws.cell(row=L.RE_ROW_EM, column=L.RE_MATRIX_COL + j).value
        exp = m.earned_index_m[j]
        ok_e &= (approx(cell, exp, 1e-9) or (cell in (None, "") and math.isnan(exp)))
    check("48 cohort written-index values tie oracle", ok_w)
    check("48 cohort written-mod values tie oracle", ok_m)
    check("36 monthly earned-index values tie oracle", ok_e)

    # EVERY BU x state combo vs oracle (via the _calc results table)
    calc = wb["_calc"]
    bad = 0
    for i, lrr in enumerate(lr_rows):
        cmb = sample_to_combo(cfg, lob, lrr, rate_rows)
        om = engine.run_bridge(p, cmb, "monthly")
        r = L.CALC_RES_FIRST + i
        for colL, exp in (("C", om.crl_ind), ("E", om.e_cy[p]), ("F", om.e_cy[p + 1]),
                          ("G", om.m_earned[p]), ("I", om.a_rate_p), ("K", om.a_mod_p),
                          ("M", om.lr_current), ("N", om.cy_lr_p), ("O", om.cy_lr_p1)):
            if not approx(calc[f"{colL}{r}"].value, exp, 1e-9):
                bad += 1
                if bad <= 5:
                    key = f"{lrr['bu']}|{lrr['state']}"
                    print(f"  FAIL detail: {key} _calc!{colL}{r} wb="
                          f"{calc[f'{colL}{r}'].value} oracle={exp}")
    check(f"all {len(lr_rows)} BU x state combos tie oracle (9 metrics each)", bad == 0,
          f"{bad} mismatches")

    # solver seed + Mode B
    sres = engine.solve_rate_for_target(p, combo, m.cy_lr_p, dt.date(p, 4, 1))
    check("Solver Mode A returns the seeded +5.0% (acceptance §11.7)",
          approx(nval(wb, "nr_SolverR"), 0.05, 1e-9), str(nval(wb, "nr_SolverR")))
    check("Solver Mode A ties oracle solver", approx(nval(wb, "nr_SolverR"),
          sres.required_change, 1e-9))
    sv = wb["Solver"]
    tt = engine.solver_timing_table(p, combo, r=0.05, target_cy_lr=m.cy_lr_p)
    ok_b = all(approx(sv[f"G{28 + row['month']}"].value, row["cy_lr"], 1e-9) for row in tt)
    check("Solver Mode B 12-month table ties oracle", ok_b)
    if lob.term_months == 12:
        check("Solver Mode B latest qualifying month is April",
              sv["D42"].value == f"Apr {p}", str(sv["D42"].value))

    # scenarios with blank levers reproduce base
    sc = wb["Scenarios"]
    ok_s = all(approx(sc[f"{cL}14"].value, m.cy_lr_p, 1e-9) for cL in "DEFG")
    check("blank-lever scenarios reproduce Base CY LR", ok_s)

    # State Summary in 'All' mode: EP-weighted aggregates per state (D38)
    ss = wb["State Summary"]
    results = {}
    for lrr in lr_rows:
        om = engine.run_bridge(p, sample_to_combo(cfg, lob, lrr, rate_rows), "monthly")
        results[(lrr["bu"], lrr["state"])] = (lrr["ep"], om)
    bad = 0
    tot_ep = tot_l27 = tot_l28 = 0.0
    for si, state in enumerate(cfg.states):
        rows_ = [(ep, om) for (bu, s), (ep, om) in results.items() if s == state]
        eps = sum(ep for ep, _ in rows_)
        w27 = sum(ep * om.cy_lr_p for ep, om in rows_) / eps
        w28 = sum(ep * om.cy_lr_p1 for ep, om in rows_) / eps
        wcrl = sum(ep * om.crl_ind for ep, om in rows_) / eps
        wecy = sum(ep * om.e_cy[p] for ep, om in rows_) / eps
        r = 8 + si
        for colL, exp in (("B", eps), ("O", wcrl), ("P", wecy), ("U", w27), ("Y", w28)):
            if not approx(ss[f"{colL}{r}"].value, exp, 1e-6):
                bad += 1
                if bad <= 4:
                    print(f"  FAIL detail: StateSummary {state} {colL}{r} "
                          f"wb={ss[f'{colL}{r}'].value} exp={exp}")
        tot_ep += eps
        tot_l27 += sum(ep * om.cy_lr_p for ep, om in rows_)
        tot_l28 += sum(ep * om.cy_lr_p1 for ep, om in rows_)
    check("State Summary (All): every state row EP-weighted correctly", bad == 0,
          f"{bad} mismatches")
    tot_row = 8 + len(cfg.states) + 3 + 1   # 3 live-roster spare rows (D42)
    check("State Summary (All): total row EP-weighted correctly",
          approx(ss[f"U{tot_row}"].value, tot_l27 / tot_ep, 1e-9)
          and approx(ss[f"Y{tot_row}"].value, tot_l28 / tot_ep, 1e-9)
          and approx(ss[f"B{tot_row}"].value, tot_ep, 1e-6))
    # live dimension lists (D42): _lists uniques match the seeded roster
    ls = wb["_lists"]
    got_bus = [ls[f"A{3 + i}"].value for i in range(len(cfg.business_units))]
    got_states = [ls[f"B{3 + i}"].value for i in range(len(cfg.states))]
    check("_lists live BU uniques match tbl_LR", got_bus == list(cfg.business_units),
          str(got_bus))
    check("_lists live state uniques match tbl_LR", got_states == list(cfg.states),
          str(got_states[:5]))
    return wb


# ---------------------------------------------------------------------------
# Phase D: toggle exercises
# ---------------------------------------------------------------------------


def _mutate(src: Path, scratch: Path, mutations: dict[tuple[str, str], object]) -> Path:
    shutil.copy2(src, scratch)
    wb = openpyxl.load_workbook(scratch, data_only=False)
    for (sheet, addr), value in mutations.items():
        wb[sheet][addr] = value
    wb.save(scratch)
    recalc(scratch)
    return scratch


def _read(scratch: Path):
    return openpyxl.load_workbook(scratch, data_only=True)


def phase_d(path: Path, cfg, lob, scratch_dir: Path):
    print("Phase D: toggle exercises (mutate -> recalc -> tie to oracle)")
    p = cfg.plan_year
    st = cfg.states
    lr_rows = sample_lr_rows(cfg, lob)
    rate_rows = sample_rate_rows(cfg, lob)
    season_rows = sample_seasonality_rows(cfg)
    we = next(r for r in lr_rows if r["bu"] == "BU-A" and r["state"] == st[0])
    base_combo = sample_to_combo(cfg, lob, we, rate_rows)
    base_m = engine.run_bridge(p, base_combo, "monthly")
    scratch = scratch_dir / "exercise.xlsx"
    we_row_xl = L.LR_FIRST + lr_rows.index(we)

    def combo_of(bu, state, **kw):
        row = next(r for r in lr_rows if r["bu"] == bu and r["state"] == state)
        return sample_to_combo(cfg, lob, row, rate_rows, **kw)

    def run(label_, mutations, assertions):
        _mutate(path, scratch, mutations)
        wb = _read(scratch)
        errs = scan_errors(wb)
        check(f"[{label_}] zero formula errors", not errs, "; ".join(errs[:5]))
        assertions(wb)

    # 1. master mod toggle OFF
    off_m = engine.run_bridge(p, replace(base_combo, mod_adjustment_enabled=False), "monthly")

    def a1(wb):
        check("[mod master OFF] A_mod = 1", approx(nval(wb, "nr_Amod_P"), 1.0, 1e-12))
        check("[mod master OFF] CY LR ties oracle", approx(nval(wb, "nr_CYLR_P"),
              off_m.cy_lr_p, 1e-9))
        check("[mod master OFF] CY LR moved UP (A_mod was < 1)",
              float(nval(wb, "nr_CYLR_P")) > base_m.cy_lr_p)
    run("mod master OFF", {("Control", "C13"): "OFF"}, a1)

    # 2. per-combo mod toggle OFF
    def a2(wb):
        check("[combo mod OFF] CY LR ties oracle", approx(nval(wb, "nr_CYLR_P"),
              off_m.cy_lr_p, 1e-9))
    run("combo mod OFF", {("Inputs", f"P{we_row_xl}"): "OFF"}, a2)

    # 3. basis proposed on the worked-example combo
    prop_m = engine.run_bridge(p, replace(base_combo, lr_basis="proposed", sel_change=0.05),
                               "monthly")

    def a3(wb):
        check("[basis proposed] LR_current = LR x (1+s)",
              approx(nval(wb, "nr_LRcur"), 0.65 * 1.05, 1e-12))
        check("[basis proposed] CY LR ties oracle", approx(nval(wb, "nr_CYLR_P"),
              prop_m.cy_lr_p, 1e-9))
    run("basis proposed", {("Inputs", f"D{we_row_xl}"): "proposed"}, a3)

    # 4. seasonality ON for a state with a seeded profile
    se_state = season_rows[0]["state"] if season_rows else None
    if se_state:
        se_on = combo_of("BU-A", se_state, seasonality_on=True, season_rows=season_rows)
        se_off = combo_of("BU-A", se_state)
        se_on_m = engine.run_bridge(p, se_on, "monthly")
        se_off_m = engine.run_bridge(p, se_off, "monthly")

        def a4(wb):
            check(f"[seasonality ON {se_state}] CY LR ties oracle (seasonal weights)",
                  approx(nval(wb, "nr_CYLR_P"), se_on_m.cy_lr_p, 1e-9),
                  f"wb={nval(wb, 'nr_CYLR_P')} oracle={se_on_m.cy_lr_p}")
            check(f"[seasonality ON {se_state}] seasonality changes some CY index",
                  any(not approx(se_on_m.e_cy[y], se_off_m.e_cy[y], 1e-12)
                      for y in (p - 1, p, p + 1)),
                  f"on={se_on_m.e_cy} off={se_off_m.e_cy}")
        run(f"seasonality ON {se_state}",
            {("Control", "C8"): se_state, ("Control", "C12"): "ON"}, a4)

        def a4b(wb):
            check(f"[{se_state} seasonality OFF] CY LR ties oracle (uniform)",
                  approx(nval(wb, "nr_CYLR_P"), se_off_m.cy_lr_p, 1e-9))
        run(f"{se_state} seasonality OFF", {("Control", "C8"): se_state}, a4b)

    # 5. selector subsample: every BU x first 3 states + every showcase combo
    picks = {(bu, s) for bu in cfg.business_units for s in st[:3]}
    for special in [("BU-B", st[1] if len(st) > 1 else st[0]),
                    ("BU-B", st[2] if len(st) > 2 else st[0]),
                    ("BU-A", st[3] if len(st) > 3 else st[0]),
                    ("BU-A", st[5] if len(st) > 5 else st[0]),
                    ("BU-C", st[6] if len(st) > 6 else st[0]),
                    ("BU-C", st[8] if len(st) > 8 else st[0]),  # net-selection showcase
                    degenerate_combo(cfg)]:
        picks.add(special)
    for bu, state in sorted(picks):
        om = engine.run_bridge(p, combo_of(bu, state), "monthly")

        def a5(wb, om=om, bu=bu, state=state):
            check(f"[selector {bu}|{state}] Bridge ties oracle",
                  approx(nval(wb, "nr_CYLR_P"), om.cy_lr_p, 1e-9),
                  f"wb={nval(wb, 'nr_CYLR_P')} oracle={om.cy_lr_p}")
        run(f"selector {bu}|{state}",
            {("Control", "C7"): bu, ("Control", "C8"): state}, a5)

    # 6. degenerate combo (seeded with no rate rows): factors = 1 exactly
    deg_bu, deg_state = degenerate_combo(cfg)
    deg_m = engine.run_bridge(p, combo_of(deg_bu, deg_state), "monthly")

    def a6(wb):
        check("[degenerate combo] A_rate = 1 exactly", approx(nval(wb, "nr_Arate_P"), 1.0, 1e-12))
        check("[degenerate combo] CY LR ties oracle", approx(nval(wb, "nr_CYLR_P"),
              deg_m.cy_lr_p, 1e-9))
    run("degenerate combo", {("Control", "C7"): deg_bu, ("Control", "C8"): deg_state}, a6)

    # 6b. blank out an existing combo's rate log -> neutral factors
    tgt = next((r for r in lr_rows
                if any(rr["bu"] == r["bu"] and rr["state"] == r["state"] for rr in rate_rows)
                and (r["bu"], r["state"]) != ("BU-A", st[0])), None)
    if tgt:
        rows_xl = [L.RL_FIRST + i for i, rr in enumerate(rate_rows)
                   if rr["bu"] == tgt["bu"] and rr["state"] == tgt["state"]]
        muts = {("Control", "C7"): tgt["bu"], ("Control", "C8"): tgt["state"]}
        for r in rows_xl:
            for colL in "ABCDEFGH":
                muts[("Inputs", f"{colL}{r}")] = None
        blank_m = engine.run_bridge(
            p, replace(combo_of(tgt["bu"], tgt["state"]), rate_changes=()), "monthly")

        def a6b(wb):
            check("[blanked rate log] A_rate = 1 exactly",
                  approx(nval(wb, "nr_Arate_P"), 1.0, 1e-12))
            check("[blanked rate log] CY LR ties oracle",
                  approx(nval(wb, "nr_CYLR_P"), blank_m.cy_lr_p, 1e-9))
        run("blanked rate log", muts, a6b)

    # 7. scenario levers, one at a time
    def transformed(dpts=0.0, shift=0, ach=None, dm1=0.0):
        changes = tuple(
            replace(rc,
                    filed_pct=rc.filed_pct + dpts,
                    effective=engine.add_months(rc.effective, shift),
                    achievement=(ach if ach is not None else rc.achievement))
            if rc.status == "planned" else rc
            for rc in base_combo.rate_changes)
        mods = base_combo.mods
        m2 = mods.m2 if mods.m2 is not None else mods.m1
        mods = replace(mods, m1=mods.m1 + dm1, m2=m2 + dm1)
        return replace(base_combo, rate_changes=changes, mods=mods)

    sc_expected = [
        engine.run_bridge(p, transformed(dpts=0.02), "monthly"),
        engine.run_bridge(p, transformed(shift=2), "monthly"),
        engine.run_bridge(p, transformed(ach=0.5), "monthly"),
        engine.run_bridge(p, transformed(dm1=0.02), "monthly"),
    ]
    muts = {("Scenarios", "C6"): 0.02, ("Scenarios", "D7"): 2,
            ("Scenarios", "E8"): 0.5, ("Scenarios", "F9"): 0.02}

    def a7(wb):
        sc = wb["Scenarios"]
        for s, om in enumerate(sc_expected, start=1):
            got = sc[f"{'CDEFG'[s]}14"].value
            check(f"[scenario S{s}] CY LR ties oracle", approx(got, om.cy_lr_p, 1e-9),
                  f"wb={got} oracle={om.cy_lr_p}")
        check("[scenario S1] rate increase lowers CY LR",
              sc_expected[0].cy_lr_p < base_m.cy_lr_p)
        check("[scenario S2] later timing raises CY LR",
              sc_expected[1].cy_lr_p > base_m.cy_lr_p)
        check("[scenario S3] under-achievement raises CY LR",
              sc_expected[2].cy_lr_p > base_m.cy_lr_p)
        check("[scenario S4] rising mods lower CY LR",
              sc_expected[3].cy_lr_p < base_m.cy_lr_p)
    run("scenario levers", muts, a7)

    # 8. solver custom target / mid-month date / achievement
    s_custom = engine.solve_rate_for_target(p, base_combo, 0.61, dt.date(p, 5, 15),
                                            achievement=0.8)

    def a8(wb):
        check("[solver custom] required r ties oracle",
              approx(nval(wb, "nr_SolverR"), s_custom.required_change, 1e-9),
              f"wb={nval(wb, 'nr_SolverR')} oracle={s_custom.required_change}")
        check("[solver custom] filed equivalent = r / 80%",
              approx(wb["Solver"]["C18"].value, s_custom.required_change / 0.8, 1e-9))
    run("solver custom", {("Solver", "C7"): 0.61, ("Solver", "C8"): dt.date(p, 5, 15),
                          ("Solver", "C9"): 0.8}, a8)

    # 9. attribution actuals
    actuals = [engine.PlannedActual(dt.date(p, 4, 1), 0.04, dt.date(p, 6, 1))]
    mods_act = replace(base_combo.mods, m1=0.90)
    att = engine.attribution(p, base_combo, actuals, mods_act, actual_cy_lr=0.655)

    def a9(wb):
        aw = wb["Attribution"]
        for row, (lbl, f) in zip((26, 27, 28, 29), att.factors):
            check(f"[attribution] factor '{lbl}' ties oracle",
                  approx(aw[f"C{row}"].value, f, 1e-9),
                  f"wb={aw[f'C{row}'].value} oracle={f}")
        check("[attribution] final LR = actual entered",
              approx(aw["D30"].value, 0.655, 1e-12))
        ck = wb["Checks"]
        check("[attribution] reconciliation check row is PASS",
              ck["G16"].value == "PASS", str(ck["G16"].value))
    run("attribution actuals",
        {("Attribution", "E14"): 0.04, ("Attribution", "F14"): dt.date(p, 6, 1),
         ("Attribution", "C8"): 0.90, ("Attribution", "C10"): 0.655}, a9)

    # 9b. State Summary filtered to one BU: rows equal that BU's combos exactly,
    # and the rate-change slots replay the chronological log
    fbu = "BU-B"

    def a9b(wb):
        ss = wb["State Summary"]
        bad = 0
        for si, state in enumerate(cfg.states):
            r = 8 + si
            om = engine.run_bridge(p, combo_of(fbu, state), "monthly")
            row_lr = next(x for x in lr_rows if x["bu"] == fbu and x["state"] == state)
            for colL, exp in (("B", row_lr["ep"]), ("U", om.cy_lr_p), ("Y", om.cy_lr_p1),
                              ("O", om.crl_ind), ("P", om.e_cy[p])):
                if not approx(ss[f"{colL}{r}"].value, exp, 1e-6):
                    bad += 1
            log = sorted([rr for rr in rate_rows
                          if rr["bu"] == fbu and rr["state"] == state],
                         key=lambda rr: rr["eff"])
            n_cell = ss[f"N{r}"].value
            if (n_cell or 0) != len(log):
                bad += 1
            for j, rr in enumerate(log[:4]):
                d = ss.cell(row=r, column=6 + j * 2).value
                v = ss.cell(row=r, column=7 + j * 2).value
                eff_pct = rr["filed"] * ((1.0 if rr["achievement"] is None
                                          else rr["achievement"])
                                         if rr["status"] == "planned" else 1.0)
                if not (hasattr(d, "date") and d.date() == rr["eff"]) and d != rr["eff"]:
                    bad += 1
                if not approx(v, eff_pct, 1e-12):
                    bad += 1
        check(f"[State Summary {fbu}] rows, totals, and rate slots tie", bad == 0,
              f"{bad} mismatches")
    run(f"State Summary filter {fbu}", {("State Summary", "B4"): fbu}, a9b)

    # 9c. trend default inheritance: blank combos inherit 2%, explicit entries win
    trend_m = engine.run_bridge(
        p, sample_to_combo(cfg, lob, we, rate_rows, trend_default=0.02), "monthly")
    ovr_state = st[1] if len(st) > 1 else st[0]
    ovr_row = next(r for r in lr_rows if r["bu"] == "BU-B" and r["state"] == ovr_state)
    ovr_m = engine.run_bridge(
        p, sample_to_combo(cfg, lob, ovr_row, rate_rows, trend_default=0.02), "monthly")
    ovr_calc_row = L.CALC_RES_FIRST + lr_rows.index(ovr_row)

    def a9c(wb):
        check("[trend default 2%] blank-trend combo inherits it for CY P+1",
              approx(nval(wb, "nr_CYLR_P1"), trend_m.cy_lr_p1, 1e-9),
              f"wb={nval(wb, 'nr_CYLR_P1')} oracle={trend_m.cy_lr_p1}")
        check("[trend default 2%] CY P untouched by trend",
              approx(nval(wb, "nr_CYLR_P"), base_m.cy_lr_p, 1e-9))
        got = wb["_calc"][f"O{ovr_calc_row}"].value
        check("[trend default 2%] explicit per-state trend still wins",
              approx(got, ovr_m.cy_lr_p1, 1e-9), f"wb={got} oracle={ovr_m.cy_lr_p1}")
    run("trend default 2%", {("Control", "C15"): 0.02}, a9c)

    # 9d. net rate selection on the worked-example combo (D39)
    net_m = engine.run_bridge(p, replace(base_combo, net_sel_p=0.08), "monthly")

    def a9d(wb):
        check("[net 8%] CY LR ties oracle net path",
              approx(nval(wb, "nr_CYLR_P"), net_m.cy_lr_p, 1e-9),
              f"wb={nval(wb, 'nr_CYLR_P')} oracle={net_m.cy_lr_p}")
        check("[net 8%] CY LR(P+1) ties oracle (selection carried)",
              approx(nval(wb, "nr_CYLR_P1"), net_m.cy_lr_p1, 1e-9))
        check("[net 8%] A_mod collapses to 1", approx(nval(wb, "nr_Amod_P"), 1.0, 1e-12))
        check("[net 8%] net mode flag TRUE", nval(wb, "nr_NetMode") is True)
    run("net selection 8%", {("Inputs", f"Q{we_row_xl}"): 0.08}, a9d)

    # 9e. scenario D-net lever on top of a net selection
    net_s1 = engine.run_bridge(p, replace(base_combo, net_sel_p=0.10), "monthly")

    def a9e(wb):
        got = wb["Scenarios"]["D14"].value
        check("[net + D-net 2%] scenario S1 ties oracle at 10% net",
              approx(got, net_s1.cy_lr_p, 1e-9), f"wb={got} oracle={net_s1.cy_lr_p}")
        check("[net + D-net 2%] base column still ties 8% net",
              approx(wb["Scenarios"]["C14"].value, net_m.cy_lr_p, 1e-9))
    run("net + scenario D-net", {("Inputs", f"Q{we_row_xl}"): 0.08,
                                 ("Scenarios", "C10"): 0.02}, a9e)

    # 9f. in-book rename of a BU and a state (D42): every dropdown list, the
    # selectors, and the State Summary must follow without regeneration
    ren_state = st[1] if len(st) > 1 else st[0]
    ren_row = next(x for x in lr_rows if x["bu"] == "BU-B" and x["state"] == ren_state)
    ren_lr_xl = L.LR_FIRST + lr_rows.index(ren_row)
    muts = {("Inputs", f"A{ren_lr_xl}"): "BU-X", ("Inputs", f"B{ren_lr_xl}"): "ZZ",
            ("Control", "C7"): "BU-X", ("Control", "C8"): "ZZ"}
    for i, rr in enumerate(rate_rows):
        if rr["bu"] == "BU-B" and rr["state"] == ren_state:
            muts[("Inputs", f"A{L.RL_FIRST + i}")] = "BU-X"
            muts[("Inputs", f"B{L.RL_FIRST + i}")] = "ZZ"
    ren_m = engine.run_bridge(p, combo_of("BU-B", ren_state), "monthly")  # values unchanged

    def a9f(wb):
        ls = wb["_lists"]
        bus = {ls[f"A{3 + i}"].value for i in range(10)}
        sts = {ls[f"B{3 + i}"].value for i in range(len(st) + 3)}
        check("[rename] new BU appears in the live BU list", "BU-X" in bus, str(sorted(
            b for b in bus if b))[:80])
        check("[rename] new state appears in the live state list", "ZZ" in sts)
        check("[rename] Bridge ties oracle for the renamed combo",
              approx(nval(wb, "nr_CYLR_P"), ren_m.cy_lr_p, 1e-9),
              f"wb={nval(wb, 'nr_CYLR_P')} oracle={ren_m.cy_lr_p}")
        ss = wb["State Summary"]
        zz_row = next((r for r in range(8, 8 + len(st) + 3)
                       if ss[f"A{r}"].value == "ZZ"), None)
        check("[rename] State Summary grew a live ZZ row", zz_row is not None)
        if zz_row:
            check("[rename] ZZ row EP and CY LR tie the renamed combo",
                  approx(ss[f"B{zz_row}"].value, ren_row["ep"], 1e-6)
                  and approx(ss[f"U{zz_row}"].value, ren_m.cy_lr_p, 1e-9),
                  f"EP={ss[f'B{zz_row}'].value} U={ss[f'U{zz_row}'].value}")
    run("in-book rename BU/state", muts, a9f)

    # 10. plan-year change (fingerprint must go N/A, engines recompute cleanly)
    p2 = p + 1
    m_p2 = engine.run_bridge(p2, base_combo, "monthly")

    def a10(wb):
        check("[plan year P+1] CY LR ties oracle for the new year",
              approx(nval(wb, "nr_CYLR_P"), m_p2.cy_lr_p, 1e-9),
              f"wb={nval(wb, 'nr_CYLR_P')} oracle={m_p2.cy_lr_p}")
        check("[plan year P+1] fingerprint FALSE", nval(wb, "orc_fp") is False)
        check("[plan year P+1] Checks still ALL PASS (oracle rows N/A)",
              nval(wb, "ck_overall") == "ALL CHECKS PASS",
              str(nval(wb, "ck_overall")))
    run("plan year change", {("Control", "C6"): p2}, a10)

    # 10b. plan year set BACK a year (the "reproduce last year" test, D44):
    # values AND the live year labels must both follow the Control input
    pm1 = p - 1
    m_pm1 = engine.run_bridge(pm1, base_combo, "monthly")

    def a10b(wb):
        check(f"[plan year {pm1}] CY LR ties oracle computed for {pm1}",
              approx(nval(wb, "nr_CYLR_P"), m_pm1.cy_lr_p, 1e-9),
              f"wb={nval(wb, 'nr_CYLR_P')} oracle={m_pm1.cy_lr_p}")
        check(f"[plan year {pm1}] E_CY ties oracle",
              approx(nval(wb, "nr_ECY_P"), m_pm1.e_cy[pm1], 1e-9))
        ss_hdr = wb["State Summary"].cell(row=7, column=21).value
        check(f"[plan year {pm1}] State Summary header relabels live",
              ss_hdr == f"CY {pm1} plan LR", str(ss_hdr))
        pf_hdr = wb["Portfolio"].cell(row=L.PF_HDR, column=7).value
        check(f"[plan year {pm1}] Portfolio header relabels live",
              pf_hdr == f"CY {pm1} plan LR", str(pf_hdr))
        banner = wb["Control"]["B10"].value
        check(f"[plan year {pm1}] Control stale-year banner appears",
              isinstance(banner, str) and banner.startswith("NOTE"), str(banner)[:60])
        check(f"[plan year {pm1}] Checks still ALL PASS",
              nval(wb, "ck_overall") == "ALL CHECKS PASS")
    run(f"plan year back to {pm1}", {("Control", "C6"): pm1}, a10b)


# ---------------------------------------------------------------------------


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lob", default="Property")
    ap.add_argument("--workbook", default=None,
                    help="override the workbook path (default derived from config + --lob)")
    ap.add_argument("--config", default="config/config.yaml")
    ap.add_argument("--quick", action="store_true", help="skip phase D exercises")
    ap.add_argument("--skip-recalc", action="store_true",
                    help="assume the workbook already has cached values")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    L.configure(cfg)
    lob = cfg.lob(args.lob)
    path = Path(args.workbook) if args.workbook else output_path(cfg, Path(cfg.output_dir), args.lob)
    print(f"Verifying {path} (LOB: {lob.name}, term {lob.term_months} months)")
    scratch_dir = path.parent / "_verify_scratch"
    scratch_dir.mkdir(exist_ok=True)

    phase_a(path)
    tie_default_state(path, cfg, lob, do_recalc=not args.skip_recalc)
    if not args.quick:
        phase_d(path, cfg, lob, scratch_dir)
    shutil.rmtree(scratch_dir, ignore_errors=True)

    print(f"\n{'=' * 60}\nRESULT: {PASS} passed, {FAIL} failed")
    for f in FAILURES:
        print(f"  FAIL: {f}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    sys.exit(main())
