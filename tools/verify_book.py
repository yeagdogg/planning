"""Verification harness for the combined book workbook (DECISIONS.md D66).

The book's engines are frozen, so this proves something different from the
per-LOB harness: that every harvested row still equals what the oracle says
for that combo, that the aggregation over it is right under every filter
state, and that the file is structurally sound.

    python tools/verify_book.py                 # phases A-D
    python tools/verify_book.py --quick         # skip the filter exercises
"""

from __future__ import annotations

import argparse
import shutil
import sys
from pathlib import Path

import openpyxl

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # noqa: E402

from src import engine  # noqa: E402
from src.build_workbook import (load_config, sample_lr_rows, sample_mod_rows,  # noqa: E402
                                sample_rate_rows,
                                sample_to_combo)
from src.sheets_book import COLS, PF_FIRST, SS_FIRST  # noqa: E402
from tools.build_book import book_path  # noqa: E402
from tools.harvest import harvest  # noqa: E402
from tools.recalc import recalc  # noqa: E402
from tools.verify_workbook import (BANNED_RE, approx, scan_errors)  # noqa: E402

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


def phase_a(path: Path):
    """Static scans — the same house rules the per-LOB harness enforces."""
    print("Phase A: static scans")
    wb = openpyxl.load_workbook(path, data_only=False)
    banned, ext, texted, merged = [], [], [], []
    for ws in wb.worksheets:
        if ws.merged_cells.ranges:
            merged.append(ws.title)
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if not isinstance(v, str):
                    continue
                if v.startswith("="):
                    if BANNED_RE.search(v):
                        banned.append(f"{ws.title}!{c.coordinate}")
                    if "[" in v and "]" in v:
                        ext.append(f"{ws.title}!{c.coordinate}")
                    # D41/D66: a LITERAL that begins with "=" is stored as a
                    # formula; an invalid one makes Excel refuse the file, so
                    # every "=" cell must parse as a formula, not prose
                    body = v[1:].lstrip()
                    if (" " in body and "(" not in v and "&" not in v
                            and '"' not in v):
                        texted.append(f"{ws.title}!{c.coordinate}: {v[:40]}")
    check("no banned functions", not banned, "; ".join(banned[:5]))
    check("no external-workbook references", not ext, "; ".join(ext[:5]))
    check("no prose stored as a formula (the D41 trap)", not texted,
          "; ".join(texted[:3]))
    check("no merged cells", not merged, "; ".join(merged[:3]))
    check("hidden data sheet stays hidden", wb["_book"].sheet_state == "hidden")
    return wb


def phase_bc(path: Path, cfg, do_recalc=True):
    print("Phase B: recalculate and scan for errors")
    if do_recalc:
        recalc(path)
    wb = openpyxl.load_workbook(path, data_only=True)
    errs = scan_errors(wb)
    check("zero formula errors in every sheet", not errs, "; ".join(errs[:5]))
    check("book Checks panel reads ALL CHECKS PASS",
          wb["Checks"]["C3"].value == "ALL CHECKS PASS",
          str(wb["Checks"]["C3"].value))

    print("Phase C: every harvested row vs the oracle")
    p = cfg.plan_year
    book = harvest(cfg)
    bk = wb["_book"]
    # index the oracle by (lob, bu, state) so row order never matters
    oracle = {}
    for lob in cfg.lobs:
        lr_rows = sample_lr_rows(cfg, lob)
        rate_rows = sample_rate_rows(cfg, lob)
        mod_rows = sample_mod_rows(cfg, lob)   # D70: the oracle must see the log
        for rowdef in lr_rows:
            cmb = sample_to_combo(cfg, lob, rowdef, rate_rows, mod_rows)
            om = engine.run_bridge(p, cmb, "monthly")
            prog = engine.program_basis_plan_lr(p, cmb)
            oracle[(lob.name, rowdef["bu"], rowdef["state"])] = (rowdef, om, prog)
    bad = missing = 0
    for i, rec in enumerate(book.rows):
        key = (rec["lob"], rec["bu"], rec["state"])
        if key not in oracle:
            missing += 1
            continue
        rowdef, om, prog = oracle[key]
        r = 2 + i
        pairs = ((COLS["ep"], rowdef["ep"]), (COLS["crl"], om.crl_ind),
                 (COLS["ecy_p"], om.e_cy[p]), (COLS["arate_p"], om.a_rate_p),
                 (COLS["amod_p"], om.a_mod_p), (COLS["lrcur"], om.lr_current),
                 (COLS["cylr_p"], om.cy_lr_p), (COLS["cylr_p1"], om.cy_lr_p1),
                 (COLS["cylr_prog"], prog[0]), (COLS["cylr1_prog"], prog[1]))
        if not all(approx(bk.cell(row=r, column=c).value, v, 1e-9) for c, v in pairs):
            bad += 1
    check("every harvested row ties its per-combo oracle run (10 metrics)",
          bad == 0 and missing == 0, f"{bad} mismatched, {missing} unmatched rows")

    # the aggregates, recomputed in Python over the same rows
    rows = book.rows
    ep_all = sum(r["ep"] for r in rows)
    lr_all = sum(r["ep"] * r["cylr_p"] for r in rows) / ep_all
    prog_all = sum(r["ep"] * r["cylr_prog"] for r in rows) / ep_all
    ss = wb["State Summary"]
    tot = SS_FIRST + len(book.states) + 1
    check("State Summary total EP ties the harvest",
          approx(ss[f"B{tot}"].value, ep_all, 1e-6),
          f"wb={ss[f'B{tot}'].value} py={ep_all}")
    check("State Summary total plan LR ties the EP-weighted harvest",
          approx(ss[f"AA{tot}"].value, lr_all, 1e-9),
          f"wb={ss[f'AA{tot}'].value} py={lr_all}")
    check("State Summary total program basis ties the harvest",
          approx(ss[f"AH{tot}"].value, prog_all, 1e-9))
    # per-state rows, unfiltered
    bad = 0
    for i, st in enumerate(book.states):
        sub = [r for r in rows if r["state"] == st]
        ep_s = sum(r["ep"] for r in sub)
        lr_s = sum(r["ep"] * r["cylr_p"] for r in sub) / ep_s
        r = SS_FIRST + i
        if not (approx(ss[f"B{r}"].value, ep_s, 1e-6)
                and approx(ss[f"AA{r}"].value, lr_s, 1e-9)):
            bad += 1
    check("every State Summary row ties an EP-weighted recomputation", bad == 0,
          f"{bad} mismatched states")
    # Portfolio mirrors the harvest row for row
    pf = wb["Portfolio"]
    bad = sum(0 if (pf[f"A{PF_FIRST + i}"].value == rec["lob"]
                    and approx(pf[f"I{PF_FIRST + i}"].value, rec["cylr_p"], 1e-9))
              else 1 for i, rec in enumerate(rows))
    check("Portfolio grid mirrors every harvested row", bad == 0,
          f"{bad} mismatched rows")
    return wb, book


def phase_d(path: Path, cfg, book, scratch_dir: Path):
    print("Phase D: filter exercises")
    scratch_dir.mkdir(parents=True, exist_ok=True)
    scratch = scratch_dir / "book_exercise.xlsx"
    rows = book.rows
    tot = SS_FIRST + len(book.states) + 1

    def run(label_, lob, bu, state):
        shutil.copy2(path, scratch)
        wb = openpyxl.load_workbook(scratch, data_only=False)
        ctl = wb["Control"]
        ctl["B7"], ctl["B8"], ctl["B9"] = lob, bu, state
        wb.save(scratch)
        recalc(scratch)
        got = openpyxl.load_workbook(scratch, data_only=True)
        errs = scan_errors(got)
        check(f"[{label_}] zero formula errors", not errs, "; ".join(errs[:3]))
        # State Summary honours LOB + BU (never its own dimension)
        sub = [r for r in rows
               if (lob == "All" or r["lob"] == lob)
               and (bu == "All" or r["bu"] == bu)]
        ep = sum(r["ep"] for r in sub)
        ss = got["State Summary"]
        check(f"[{label_}] State Summary total EP matches the subset",
              approx(ss[f"B{tot}"].value, ep, 1e-6),
              f"wb={ss[f'B{tot}'].value} py={ep}")
        if ep:
            lr = sum(r["ep"] * r["cylr_p"] for r in sub) / ep
            check(f"[{label_}] State Summary total plan LR matches the subset",
                  approx(ss[f"AA{tot}"].value, lr, 1e-9),
                  f"wb={ss[f'AA{tot}'].value} py={lr}")
        # the Control KPI band honours all three
        sub3 = [r for r in sub if state == "All" or r["state"] == state]
        ep3 = sum(r["ep"] for r in sub3)
        check(f"[{label_}] Control EP-in-view matches all three filters",
              approx(got["Control"]["C13"].value, ep3, 1e-6),
              f"wb={got['Control']['C13'].value} py={ep3}")
        check(f"[{label_}] combos in view", got["Control"]["A13"].value == len(sub3),
              f"wb={got['Control']['A13'].value} py={len(sub3)}")
        check(f"[{label_}] book Checks still pass",
              got["Checks"]["C3"].value == "ALL CHECKS PASS")

    run("all", "All", "All", "All")
    run(f"line={book.lobs[0]}", book.lobs[0], "All", "All")
    run(f"bu={book.business_units[0]}", "All", book.business_units[0], "All")
    run(f"state={book.states[0]}", "All", "All", book.states[0])
    run("one combo", book.lobs[-1], book.business_units[-1], book.states[-1])
    shutil.rmtree(scratch_dir, ignore_errors=True)


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Verify the combined book workbook.")
    ap.add_argument("--config", default="config/config.yaml")
    ap.add_argument("--workbook", default=None)
    ap.add_argument("--quick", action="store_true", help="skip the filter exercises")
    ap.add_argument("--skip-recalc", action="store_true")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    path = Path(args.workbook) if args.workbook else book_path(cfg, Path(cfg.output_dir))
    print(f"Verifying {path}")
    phase_a(path)
    _wb, book = phase_bc(path, cfg, do_recalc=not args.skip_recalc)
    if not args.quick:
        phase_d(path, cfg, book, path.parent / "_book_scratch")
    print(f"\n{'=' * 60}\nRESULT: {PASS} passed, {FAIL} failed")
    for f in FAILURES:
        print(f"  FAIL: {f}")
    return 1 if FAIL else 0


if __name__ == "__main__":
    raise SystemExit(main())
