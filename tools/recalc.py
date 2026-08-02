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

import datetime as _dt
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


# A single Excel process, reused across calls (D86). Spawning and quitting one
# per recalculation cost seconds each, and the verification harness calls this
# ~50 times per LOB phase D and ~300 times per release — that lifecycle was the
# bulk of the wall clock, not the calculation. The instance is created lazily,
# torn down at exit, and discarded on any COM error so a wedged Excel cannot
# poison later calls.
_XL = None
_COM_READY = False


def _com_init() -> None:
    """Initialise COM for this process, exactly once.

    Deliberately NOT paired with a CoUninitialize per call: a COM object belongs
    to the apartment that created it, so tearing the apartment down while still
    holding the Excel proxy leaves a handle that later raises AttributeError on
    its own methods (observed: 'Excel.Application.CalculateFullRebuild' partway
    through a phase-D run). The apartment lives as long as the instance does.
    """
    global _COM_READY
    if not _COM_READY:
        import pythoncom
        pythoncom.CoInitialize()
        _COM_READY = True


def _excel():
    """The shared Excel application object, started on first use."""
    global _XL
    import atexit

    import win32com.client

    _com_init()
    if _XL is None:
        _XL = win32com.client.DispatchEx("Excel.Application")
        _XL.Visible = False
        _XL.DisplayAlerts = False
        _XL.AskToUpdateLinks = False
        _XL.EnableEvents = False
        atexit.register(shutdown_excel)
    return _XL


def shutdown_excel() -> None:
    """Quit the shared instance if one is running (idempotent)."""
    global _XL, _COM_READY
    xl, _XL = _XL, None
    if xl is not None:
        try:
            xl.Quit()
        except Exception:  # noqa: BLE001 - it may already be gone
            pass
    if _COM_READY:
        try:
            import pythoncom
            pythoncom.CoUninitialize()
        except Exception:  # noqa: BLE001
            pass
        _COM_READY = False


def recalc_with_excel(path: Path, attempts: int = 3) -> None:
    """Open in the shared Excel instance, full rebuild, save in place.

    Retries transient COM failures (RPC_E_CALL_REJECTED and friends) that occur
    when an Excel instance is busy starting up or shutting down; each retry
    discards the instance first, so a genuinely broken one is replaced rather
    than reused.
    """
    import time

    import pythoncom

    last_exc: Exception | None = None
    for attempt in range(attempts):
        wb = None
        try:
            xl = _excel()
            wb = xl.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
            xl.CalculateFullRebuild()
            # Excel recalculates asynchronously; wait for done (xlDone = 0).
            # Deadline-based, not a bare spin: a spin of N property reads can
            # exhaust in milliseconds and say nothing about whether the
            # calculation finished.
            deadline = time.monotonic() + 300.0
            while xl.CalculationState != 0:
                if time.monotonic() > deadline:
                    raise RuntimeError(
                        f"Excel still calculating after 300s on {path.name}")
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
            wb.Save()
            wb.Close(SaveChanges=False)
            return
        except Exception as e:  # noqa: BLE001 - retry transient COM errors
            last_exc = e
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
            shutdown_excel()          # never retry into a suspect instance
        time.sleep(2.0 * (attempt + 1))
    raise last_exc  # type: ignore[misc]


_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)


def com_value(v):
    """A Python value as Excel would store it, for writing through COM (D86).

    Dates become SERIAL NUMBERS. Handing pywin32 a naive datetime lets it
    round-trip through local-timezone conversion, shifting the value by the UTC
    offset and sometimes onto the previous day — an error that does not look
    like one, because everything downstream still computes, just from the wrong
    day. (It surfaced as 18 harness failures, every one a fourth-decimal drift
    in an exercise that mutates a date.) A serial is exactly what Excel stores
    and exactly what openpyxl wrote, so the cell's date format still renders it.
    """
    if isinstance(v, _dt.datetime):
        delta = v - _EXCEL_EPOCH
        return delta.days + delta.seconds / 86400.0
    if isinstance(v, _dt.date):
        return (v - _EXCEL_EPOCH.date()).days
    return v


def assert_calc_settings(path: Path) -> bool:
    """Put calcMode="auto" and fullCalcOnLoad back into the saved file (D88).

    The generator sets both, and Excel's own save strips both — a shipped
    workbook carries a bare <calcPr calcId="..."/>. That matters because Excel's
    calculation mode is a SESSION property taken from the first workbook opened:
    someone who opens a manual-mode model first, then this workbook, edits an
    input and sees nothing move. Every KPI, the plan LR and the ALL CHECKS PASS
    banner would all display stale cached values, and with TODAY()/NOW() banned
    and no VBA there is nothing inside the file that could notice.

    fullCalcOnLoad is the actual guarantee (it recalculates on open whatever the
    session mode says); calcMode="auto" additionally sets the session right when
    this workbook happens to be opened first. A full rebuild costs about a third
    of a second, so the belt and braces are free.

    Patched at the zip level, deliberately: round-tripping through openpyxl would
    discard the cached values the recalculation just produced, which is the one
    thing this whole module exists to create.
    """
    import re
    import zipfile

    name = "xl/workbook.xml"
    with zipfile.ZipFile(path) as z:
        if name not in z.namelist():
            return False
        entries = [(i, z.read(i.filename)) for i in z.infolist()]
    xml = next(d for i, d in entries if i.filename == name).decode("utf-8")
    m = re.search(r"<calcPr\b[^>]*?/?>", xml)
    if not m:
        return False
    tag = m.group(0)
    new = tag
    for attr, value in (("calcMode", "auto"), ("fullCalcOnLoad", "1")):
        if re.search(rf'\b{attr}\s*=', new):
            new = re.sub(rf'\b{attr}\s*=\s*"[^"]*"', f'{attr}="{value}"', new)
        else:
            new = new.replace("<calcPr", f'<calcPr {attr}="{value}"', 1)
    if new == tag:
        return False
    xml = xml.replace(tag, new, 1)

    tmp = path.with_suffix(".calcpr.tmp")
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as out:
        for info, data in entries:
            out.writestr(info, xml.encode("utf-8") if info.filename == name else data)
    tmp.replace(path)
    return True


def mutate_and_recalc(path: Path, mutations, attempts: int = 3) -> None:
    """Write cells and recalculate, in ONE Excel visit (D86).

    The verification harness used to apply its phase-D mutations by loading the
    workbook in openpyxl, assigning cells, and saving — then handing the file to
    Excel to recalculate. Measured on the sample Property book that round trip is
    4.3 of every 13.9 seconds per exercise (31%), spent parsing and re-emitting
    171k formulas to change four cells. Excel already has to open the file; it
    can take the edits on the way through.

    ``mutations`` maps (sheet name, A1 address) -> literal value.

    Dates are written as Excel SERIAL NUMBERS, not as datetimes. Handing pywin32
    a naive datetime lets it round-trip through local-timezone conversion, which
    shifts the value by the UTC offset and can land it on the previous day — the
    kind of error that does not look like an error: the harness came back with 18
    failures, all of them fourth-decimal-place drifts in exercises that mutate a
    date (day-blends, solver effective dates, attribution actuals). A serial is
    exactly what Excel stores and exactly what openpyxl wrote before, so the
    cell's existing date format renders it and every formula sees the same value.
    """
    import time

    import pythoncom

    last_exc: Exception | None = None
    for attempt in range(attempts):
        wb = None
        try:
            xl = _excel()
            wb = xl.Workbooks.Open(str(path.resolve()), UpdateLinks=0)
            for (sheet, addr), value in mutations.items():
                wb.Worksheets(sheet).Range(addr).Value = com_value(value)
            # Explicit full rebuild: this file descends from an Excel-saved
            # artifact, whose save clears fullCalcOnLoad, so opening it is not
            # itself a guarantee that anything recalculated.
            xl.CalculateFullRebuild()
            deadline = time.monotonic() + 300.0
            while xl.CalculationState != 0:
                if time.monotonic() > deadline:
                    raise RuntimeError(
                        f"Excel still calculating after 300s on {path.name}")
                pythoncom.PumpWaitingMessages()
                time.sleep(0.01)
            wb.Save()
            wb.Close(SaveChanges=False)
            return
        except Exception as e:  # noqa: BLE001 - retry transient COM errors
            last_exc = e
            try:
                if wb is not None:
                    wb.Close(SaveChanges=False)
            except Exception:  # noqa: BLE001
                pass
            shutdown_excel()
        finally:
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


def recalc(path: str | Path, engine: str = "auto", keep_calc_settings=True) -> str:
    """Recalculate in place. Returns the engine used ('excel' | 'libreoffice').

    ``keep_calc_settings`` re-asserts calcMode/fullCalcOnLoad afterwards (D88).
    The harness turns it off for its throwaway scratch copies, which are read
    once by openpyxl and deleted — repacking the zip there is pure cost.
    """
    path = Path(path)
    if not path.exists():
        raise FileNotFoundError(path)
    errors = {}
    if engine in ("auto", "excel"):
        try:
            recalc_with_excel(path)
            if keep_calc_settings:
                assert_calc_settings(path)
            return "excel"
        except Exception as e:  # noqa: BLE001 - fall through to LibreOffice
            errors["excel"] = e
            if engine == "excel":
                raise
    if engine in ("auto", "libreoffice"):
        try:
            recalc_with_libreoffice(path)
            if keep_calc_settings:
                assert_calc_settings(path)
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
