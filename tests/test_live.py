"""The live phase-D session (D105).

Two halves. The value conversions are pure and always run: they are where the
subtle wrongness would live, because every one of them is a place COM disagrees
with openpyxl about the same cell, and a silent disagreement in a verification
harness is worse than a crash.

The rest needs a live Excel and is skipped without one. Those tests are not
about the happy path — the harness itself proves that, 317 assertions at a time.
They exist to prove the two GUARDS fire, because the guards are the whole reason
mutating a shared open workbook is allowed here at all (D86a reverted this idea
once, after a silent write failure graded a default workbook as if it were the
mutated one).

A run of these may print::

    Windows fatal exception: code 0x80010001

That is RPC_E_CALL_REJECTED and it is NOT a crash — it is Excel saying "busy"
while it finishes starting up, which ``busy_retry`` waits out (D100). pytest
enables faulthandler, which reports the structured exception before Python ever
sees it, so the traceback points AT the retry loop that is handling it. The suite
still passes. This note is here because that message cost a full session once.
"""

from __future__ import annotations

import datetime as dt
import shutil
from pathlib import Path

import pytest

from tools.live import ERR_BY_CODE, _py, _split_ref

# ---------------------------------------------------------------------------
# pure: how a COM value has to be read to mean what openpyxl means
# ---------------------------------------------------------------------------


# 0x800A0000 as a signed 32-bit int: the VT_ERROR base pywin32 reports, with
# Excel's error code in the low 16 bits. Pinned by the five live values below.
VT_ERROR_BASE = -2146828288


@pytest.mark.parametrize("code,text", sorted(ERR_BY_CODE.items()))
def test_error_variants_read_back_as_their_error_string(code, text):
    """pywin32 hands back an int whose low 16 bits are Excel's error code. The
    error scan matches on "#DIV/0!" strings, so the int has to become one —
    otherwise a workbook full of errors scans clean."""
    assert _py(VT_ERROR_BASE + code) == text


def test_the_error_codes_are_the_ones_excel_actually_sends():
    """Pinned against a live instance before the module was written: these five
    are the exact ints a real broken formula produced, and they are what fixes
    VT_ERROR_BASE above."""
    assert _py(-2146826281) == "#DIV/0!"      # =1/0
    assert _py(-2146826273) == "#VALUE!"      # =1+"x"
    assert _py(-2146826246) == "#N/A"         # =NA()
    assert _py(-2146826259) == "#NAME?"       # =NOTAFUNC()
    assert _py(-2146826252) == "#NUM!"        # =SQRT(-1)
    assert VT_ERROR_BASE + 2007 == -2146826281


def test_numbers_are_never_mistaken_for_errors():
    """The discriminator is the TYPE: Excel marshals every real number as a
    float, so an int can only be an error variant. If that ever stopped being
    true, 42 would read as an error code and the scan would cry wolf."""
    assert _py(42.0) == 42.0
    assert _py(0.0) == 0.0
    assert _py(-1.5) == -1.5
    assert _py(2007.0) == 2007.0             # the #DIV/0! code, as a value


def test_booleans_survive_the_int_check():
    """bool IS an int in Python, so it has to be tested first or TRUE becomes
    the error whose code is 1."""
    assert _py(True) is True
    assert _py(False) is False


def test_blank_reads_as_none_the_way_openpyxl_reports_it():
    """A formula returning "" has no cached value in the file, so openpyxl says
    None while COM says "". Assertions are written against None."""
    assert _py("") is None
    assert _py(None) is None
    assert _py("—") == "—"                   # a real string is untouched
    assert _py("OFF") == "OFF"


def test_datetimes_lose_the_timezone_pywin32_invents():
    """An Excel serial has no timezone; pywin32 stamps the machine's local one
    onto it. Kept, it makes a tz-aware datetime uncomparable with the oracle's
    naive one — and east of UTC it can shift the date itself."""
    from datetime import timezone, timedelta

    aware = dt.datetime(2027, 4, 1, tzinfo=timezone(timedelta(hours=10)))
    got = _py(aware)
    assert got == dt.datetime(2027, 4, 1)
    assert got.tzinfo is None
    assert got.date() == dt.date(2027, 4, 1)      # the day did not move


def test_defined_name_refs_split_including_quoted_sheets():
    assert _split_ref("=Bridge!$D$15") == ("Bridge", "$D$15")
    assert _split_ref("='State Summary'!$A$1") == ("State Summary", "$A$1")
    assert _split_ref("=_calc!$A$4") == ("_calc", "$A$4")
    # a sheet whose name contains the separator character
    assert _split_ref("='Net!Delivery'!$B$2") == ("Net!Delivery", "$B$2")
    with pytest.raises(ValueError):
        _split_ref("=#REF!")


# ---------------------------------------------------------------------------
# live: the two guards that make in-place mutation safe
# ---------------------------------------------------------------------------

WB = Path("output/Plan_LR_Workbook_2027_Property.xlsx")


def _excel_available() -> bool:
    try:
        import win32com.client  # noqa: F401
    except ImportError:
        return False
    return WB.exists()


live_only = pytest.mark.skipif(
    not _excel_available(),
    reason="needs Excel and a built Property workbook")


@pytest.fixture(scope="module")
def book(tmp_path_factory):
    from tools.live import LiveBook
    from tools.recalc import shutdown_excel

    lb = LiveBook(WB, tmp_path_factory.mktemp("live") / "wb.xlsx",
                  canary=("nr_CYLR_P", "ck_overall"))
    yield lb
    lb.close()
    shutdown_excel()


@live_only
def test_a_mutation_that_does_not_land_raises_instead_of_grading(book, monkeypatch):
    """D86a's exact failure: the write silently does not happen, and the harness
    goes on to assert against the DEFAULT workbook. Those assertions mostly
    fail, which looks like a broken workbook — but on a different day they pass,
    and that is a false green. It must raise."""
    monkeypatch.setattr(book, "_write", lambda *a, **kw: None)
    with pytest.raises(RuntimeError, match="mutation did not take"):
        book.apply({("Control", "C13"): "OFF"})


@live_only
def test_a_restore_that_does_not_land_is_caught_before_the_next_exercise(book):
    """The other direction: state leaks forward and every later exercise is
    graded against a workbook nobody described. The canary is re-read after
    every restore precisely so this cannot go unnoticed."""
    base = book.nval("nr_CYLR_P")
    book.apply({("Control", "C13"): "OFF"})
    assert book.nval("nr_CYLR_P") != base, "the exercise did not move the answer"

    # forget how to put one cell back, exactly as a dropped COM write would
    book._saved[("Control", "C13")] = (False, "OFF")
    with pytest.raises(RuntimeError, match="did not return to its baseline"):
        book.restore()

    # and leave the session usable for whatever runs next
    book._saved[("Control", "C13")] = (False, "ON")
    book.restore()
    assert book.nval("nr_CYLR_P") == base


@live_only
def test_an_exercise_leaves_no_trace_once_it_is_restored(book):
    """The property the whole design rests on: 42 exercises in one workbook have
    to be indistinguishable from 42 fresh copies."""
    before = {n: book.nval(n) for n in
              ("nr_CYLR_P", "nr_CYLR_P1", "nr_Arate_P", "nr_Amod_P", "nr_LRcur")}
    for mutation in ({("Control", "C13"): "OFF"},
                     {("Control", "C6"): 2026},
                     {("Solver", "C7"): 0.61, ("Solver", "C8"): dt.date(2027, 5, 15)}):
        book.apply(mutation)
    book.restore()
    assert {n: book.nval(n) for n in before} == before


@live_only
def test_a_date_mutation_lands_on_the_day_it_was_asked_for(book):
    """The timezone trap, end to end: the serial has to render as the same day
    the exercise named, or the workbook computes a right answer for a wrong
    date and nothing looks broken."""
    book.apply({("Solver", "C8"): dt.date(2027, 5, 15)})
    got = book["Solver"]["C8"].value
    assert isinstance(got, dt.datetime)
    assert got.date() == dt.date(2027, 5, 15)
    book.restore()


@live_only
def test_the_error_scan_finds_an_error_that_is_really_there(book):
    """A scan that cannot fail is not a scan. Break a cell on purpose and make
    sure SpecialCells reports it, with its address.

    Written against the primitives rather than apply(), because apply() proves a
    mutation landed by comparing the cell's VALUE to what was asked for, and a
    formula's value is never its own text. That is a real limit of the guard and
    it is shared with the path this replaced: no phase-D exercise mutates a cell
    into a formula, and if one ever does, both paths will say so loudly.
    """
    from tools.live import scan_errors_live

    assert scan_errors_live(book) == []
    book._snapshot("Scenarios", "C5")
    book._write("Scenarios", "C5", "=1/0")
    book.calculate()
    errs = scan_errors_live(book)
    assert any("Scenarios!C5" in e and "#DIV/0!" in e for e in errs), errs
    book.restore()
    assert scan_errors_live(book) == []
