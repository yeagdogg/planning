"""Headless workbook recalculation.

openpyxl writes formulas without cached values, so a generated workbook must be
recalculated once before cached values can be read back (data_only=True).

Engines, in order of preference:
  1. Microsoft Excel via COM automation (the authoritative calc engine).
  2. LibreOffice headless (soffice --convert-to xlsx round-trip).

Usage:
    python tools/recalc.py output/Plan_LR_Workbook_2027.xlsx
"""

from __future__ import annotations

import shutil
import subprocess
import sys
import tempfile
from pathlib import Path

SOFFICE_CANDIDATES = [
    "soffice",
    r"C:\Program Files\LibreOffice\program\soffice.exe",
    r"C:\Program Files (x86)\LibreOffice\program\soffice.exe",
]


def recalc_with_excel(path: Path, attempts: int = 3) -> None:
    """Open in Excel via COM, full rebuild, save in place.

    Retries transient COM failures (RPC_E_CALL_REJECTED and friends) that occur
    when an Excel instance is busy starting up or shutting down.
    """
    import time

    import pythoncom
    import win32com.client

    last_exc: Exception | None = None
    for attempt in range(attempts):
        pythoncom.CoInitialize()
        xl = None
        wb = None
        try:
            xl = win32com.client.DispatchEx("Excel.Application")
            xl.Visible = False
            xl.DisplayAlerts = False
            xl.AskToUpdateLinks = False
            xl.EnableEvents = False
            wb = xl.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
            xl.CalculateFullRebuild()
            # Excel recalculates asynchronously; wait for done (xlDone = 0).
            for _ in range(600):
                if xl.CalculationState == 0:
                    break
                pythoncom.PumpWaitingMessages()
            wb.Save()
            return
        except Exception as e:  # noqa: BLE001 - retry transient COM errors
            last_exc = e
        finally:
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
            try:
                if xl is not None:
                    xl.Quit()
            except Exception:  # noqa: BLE001
                pass
            pythoncom.CoUninitialize()
        time.sleep(2.0 * (attempt + 1))
    raise last_exc  # type: ignore[misc]


def _find_soffice() -> str | None:
    for cand in SOFFICE_CANDIDATES:
        if shutil.which(cand) or Path(cand).exists():
            return cand
    return None


def recalc_with_libreoffice(path: Path) -> None:
    """soffice --headless --convert-to xlsx round-trip (recalculates on load)."""
    soffice = _find_soffice()
    if not soffice:
        raise RuntimeError("LibreOffice (soffice) not found")
    with tempfile.TemporaryDirectory() as td:
        subprocess.run(
            [soffice, "--headless", "--calc", "--convert-to", "xlsx", "--outdir", td,
             str(path.resolve())],
            check=True, capture_output=True, timeout=600)
        produced = Path(td) / path.name
        if not produced.exists():
            raise RuntimeError(f"soffice produced no output for {path}")
        shutil.copy2(produced, path)


def recalc(path: str | Path, engine: str = "auto") -> str:
    """Recalculate in place. Returns the engine used ('excel' | 'libreoffice')."""
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    errors = {}
    if engine in ("auto", "excel"):
        try:
            recalc_with_excel(path)
            return "excel"
        except Exception as e:  # noqa: BLE001 - fall through to LibreOffice
            errors["excel"] = e
            if engine == "excel":
                raise
    if engine in ("auto", "libreoffice"):
        try:
            recalc_with_libreoffice(path)
            return "libreoffice"
        except Exception as e:  # noqa: BLE001
            errors["libreoffice"] = e
            if engine == "libreoffice":
                raise
    raise RuntimeError(
        "No recalculation engine available. Tried: "
        + "; ".join(f"{k}: {v}" for k, v in errors.items())
        + ". Open the file once in Excel manually, save, and re-run verification."
    )


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(2)
    used = recalc(sys.argv[1], engine=sys.argv[2] if len(sys.argv) > 2 else "auto")
    print(f"recalculated {sys.argv[1]} with {used}")
