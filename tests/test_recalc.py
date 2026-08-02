"""Recalculation-tooling contracts that do not need Excel (D86/D88).

The date one is here because it already cost a full harness run to find: the
harness came back with 18 failures, every one a fourth-decimal drift in an
exercise that mutates a date, because a naive datetime handed to pywin32 goes
through local-timezone conversion and can land on the previous day. Nothing
crashes when that happens — the workbook simply computes the right answer for
the wrong day.
"""

from __future__ import annotations

import datetime as dt
import zipfile

import openpyxl
import pytest

from tools.recalc import assert_calc_settings, com_value


def test_date_becomes_an_excel_serial():
    """1900-03-01 is serial 61 in Excel's (leap-bug-compatible) calendar."""
    assert com_value(dt.date(1900, 3, 1)) == 61
    assert com_value(dt.date(1899, 12, 31)) == 1


def test_known_plan_dates_round_trip():
    for d, serial in ((dt.date(2027, 1, 1), 46388),
                      (dt.date(2027, 5, 15), 46522),
                      (dt.date(2026, 10, 1), 46296)):
        assert com_value(d) == serial, d
        # and the serial maps back to the same day
        assert dt.date(1899, 12, 30) + dt.timedelta(days=com_value(d)) == d


def test_datetime_keeps_its_time_as_a_fraction():
    noon = com_value(dt.datetime(2027, 5, 15, 12, 0, 0))
    assert noon == pytest.approx(46522.5)


def test_no_timezone_shift_anywhere_in_the_year():
    """The failure mode was a whole-day slip, so check every day of a year."""
    d = dt.date(2027, 1, 1)
    while d < dt.date(2028, 1, 1):
        back = dt.date(1899, 12, 30) + dt.timedelta(days=com_value(d))
        assert back == d, f"{d} came back as {back}"
        d += dt.timedelta(days=1)


def test_non_dates_pass_through_untouched():
    for v in ("OFF", "SBA", 0.08, 1, 0, "", None, True):
        assert com_value(v) is v or com_value(v) == v


def _tiny_workbook(path, calc_pr=True):
    wb = openpyxl.Workbook()
    wb["Sheet"]["A1"] = 1
    if calc_pr:
        wb.calculation.calcMode = "manual"
        wb.calculation.fullCalcOnLoad = False
    wb.save(path)
    return path


def _calc_pr(path) -> str:
    with zipfile.ZipFile(path) as z:
        xml = z.read("xl/workbook.xml").decode("utf-8")
    import re
    m = re.search(r"<calcPr\b[^>]*?/?>", xml)
    return m.group(0) if m else ""


def test_calc_settings_are_reasserted(tmp_path):
    p = _tiny_workbook(tmp_path / "wb.xlsx")
    assert assert_calc_settings(p)
    tag = _calc_pr(p)
    assert 'fullCalcOnLoad="1"' in tag
    assert 'calcMode="auto"' in tag


def test_calc_settings_patch_is_idempotent(tmp_path):
    p = _tiny_workbook(tmp_path / "wb.xlsx")
    assert assert_calc_settings(p) is True
    assert assert_calc_settings(p) is False, "second call should find nothing to do"


def test_calc_settings_patch_preserves_the_rest_of_the_file(tmp_path):
    """It rewrites a zip; everything but workbook.xml must survive intact."""
    p = _tiny_workbook(tmp_path / "wb.xlsx")
    with zipfile.ZipFile(p) as z:
        before = {i.filename: z.read(i.filename) for i in z.infolist()}
    assert_calc_settings(p)
    with zipfile.ZipFile(p) as z:
        after = {i.filename: z.read(i.filename) for i in z.infolist()}
    assert set(before) == set(after), "entries were added or lost"
    for name in before:
        if name != "xl/workbook.xml":
            assert before[name] == after[name], name
    assert openpyxl.load_workbook(p)["Sheet"]["A1"].value == 1
