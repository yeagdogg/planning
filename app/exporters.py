"""Scenario -> workbook, through the SAME generator the release harness
verifies (P5). There is no second writer: ``scenario_to_carried`` dresses
the app's rows as ``carry.CarriedInputs`` and ``src.build_workbook.build``
does everything else, so an exported file is a real fleet workbook — the
oracle-tested formulas, checks, notes, charts — seeded with the app's
inputs.

Exports land beside the fleet with an ``_APP`` tag (the fleet's own files
are never overwritten — non-interference), and the tag rides the same
mechanism ``--tag`` uses, so ``importers.from_workbook`` can re-open an
app export: workbook → app → workbook → app is one loop.

Values: the generator writes formulas; numbers appear on first open in
Excel (fullCalcOnLoad). The optional recalc shells to the SYSTEM python's
``tools/recalc.py`` COM session — the app venv deliberately carries no
COM stack, and the recalc path stays exactly the one the release harness
exercises.
"""
from __future__ import annotations

import shutil
import subprocess
from pathlib import Path

from src.build_workbook import build
from src.carry import CarriedInputs

from . import importers

ROOT = Path(__file__).resolve().parents[1]
TAG = "APP"


def scenario_to_carried(scn) -> CarriedInputs:
    """The exact inverse of importers.from_workbook's mapping.

    ``source`` must be Path-like — the generator's Inputs banner prints
    ``carried.source.name`` ("CARRIED FORWARD from …"), so the app signs
    its exports there."""
    return CarriedInputs(
        lr_rows=[dict(r) for r in scn.lr_rows],
        rate_rows=[dict(r) for r in scn.rate_rows],
        mod_rows=[dict(r) for r in scn.mod_rows],
        season_rows=[dict(r) for r in scn.season_rows],
        term_months=int(scn.term_months),
        season_on="ON" if scn.season_on else "OFF",
        mod_master="ON" if scn.mod_master else "OFF",
        trend_default=float(scn.trend_default or 0.0),
        source=Path(f"Westfield Planning Workbench ({scn.name})"),
    )


def export_filename(scn, tag: str = TAG) -> str:
    cfg = importers.app_config()
    stem = cfg.filename.format(plan_year=scn.plan_year,
                               lob=scn.lob.replace(" ", "_"))
    return f"{stem}_{tag}.xlsx" if tag else f"{stem}.xlsx"


def export_line(scn, out_dir=None, formula_mode=None, tag: str = TAG) -> Path:
    """Build one line's workbook from its Scenario. Returns the path."""
    if not scn.combo_keys():
        raise ValueError(f"{scn.lob}: no populated tbl_LR rows to export")
    cfg = importers.app_config()
    out = Path(out_dir) if out_dir is not None else ROOT / cfg.output_dir
    out.mkdir(parents=True, exist_ok=True)
    wb = build(cfg, scn.lob, carried=scenario_to_carried(scn),
               formula_mode=formula_mode)
    path = out / export_filename(scn, tag)
    wb.save(path)
    return path


# ------------------------------------------------------------ Excel recalc

def system_python() -> str | None:
    """A python OUTSIDE the app venv (the COM stack lives there). The py
    launcher is preferred — it can never be a venv shim."""
    for cand in ("py", "python"):
        exe = shutil.which(cand)
        if exe and ".venv" not in exe.lower():
            return exe
    return None


_RECALC_SCRIPT = """\
import sys
sys.path.insert(0, {root!r})
from tools.recalc import recalc, shutdown_excel
try:
    for p in sys.argv[1:]:
        recalc(p)
        print("recalced", p, flush=True)
finally:
    shutdown_excel()
"""


def export_fleet_and_book(book: dict, out_dir=None,
                          formula_mode=None) -> tuple:
    """The summary flow (P6): every line UNTAGGED into output/app/ (an
    isolated directory — the verified fleet in output/ is never touched),
    Excel-recalc them, run tools/build_book.py --out there (the harvest
    refuses unrecalced sources by design), then recalc the Book itself.
    Returns (line paths, book path). Needs Excel end to end.
    """
    import sys

    cfg = importers.app_config()
    missing = [name for name in cfg.output_lobs if name not in book]
    if missing:
        raise ValueError(
            "the Book needs every configured line loaded — missing: "
            + ", ".join(missing))
    wrong_year = [s.lob for s in book.values()
                  if s.plan_year != cfg.plan_year]
    if wrong_year:
        raise ValueError(
            f"plan year mismatch vs config ({cfg.plan_year}): "
            + ", ".join(wrong_year))
    out = Path(out_dir) if out_dir is not None else ROOT / cfg.output_dir / "app"
    paths = [export_line(book[name], out_dir=out,
                         formula_mode=formula_mode, tag="")
             for name in cfg.output_lobs]
    recalc_files(paths)
    proc = subprocess.run(
        [sys.executable, "tools/build_book.py",
         "--config", "config/config.yaml", "--out", str(out)],
        cwd=str(ROOT), capture_output=True, text=True, timeout=300)
    if proc.returncode != 0:
        tail = (proc.stdout or proc.stderr or "").strip().splitlines()[-4:]
        raise RuntimeError("book build failed: " + " | ".join(tail))
    book_file = out / f"Plan_LR_Book_{cfg.plan_year}.xlsx"
    recalc_files([book_file])
    return paths, book_file


def recalc_files(paths, timeout: int = 600) -> str:
    """Populate values in exported workbooks via the system python's Excel
    COM session (tools/recalc.py — the release harness's own path). One
    subprocess for the whole batch, so Excel starts once."""
    exe = system_python()
    if exe is None:
        raise RuntimeError(
            "No system Python found for the Excel recalc — the files are "
            "still valid; values populate on first open in Excel.")
    script = _RECALC_SCRIPT.format(root=str(ROOT))
    proc = subprocess.run(
        [exe, "-c", script, *[str(p) for p in paths]],
        cwd=str(ROOT), capture_output=True, text=True, timeout=timeout)
    if proc.returncode != 0:
        tail = (proc.stderr or proc.stdout or "").strip().splitlines()[-6:]
        raise RuntimeError("recalc failed: " + " | ".join(tail))
    return proc.stdout
