"""Checks-sheet formulas that only LOOK like they test every row.

The Checks panel aggregates whole input columns with SUMPRODUCT, and its rows
are the workbook's own self-defence: a FAIL there is what stops a bad paste
becoming a plan. So a Checks formula that silently inspects ONE row is worse
than no formula at all — it reports PASS with authority.

That is exactly what happened. ``N()`` does not broadcast over a range inside
SUMPRODUCT; it collapses to the range's FIRST cell. Written

    =SUMPRODUCT((lr_key<>"")*(N(lr_m0asof)>=DATE(nr_PlanYear,12,31))*1)

the as-of check tested tbl_LR row 1 and nothing else for five versions. Proven
against a live workbook by planting an illegal as-of date: in row 1 the check
returned 63 (the whole populated row COUNT, because one scalar comparison was
multiplied across the mask), and in row 8 it returned 0 — blind. D107 found it
while adding a new row with the same construction.

IS-functions DO broadcast, which is why the paste-hygiene rows have always
worked, and why the fix is ISNUMBER rather than N.
"""

from __future__ import annotations

import re

import pytest

from src.build_workbook import build, load_config

# N(<something>) where <something> is a bare identifier — a defined name rather
# than a cell reference or a nested call. Cell refs and scalars are fine.
_N_OF_NAME = re.compile(r"\bN\(([A-Za-z_][A-Za-z0-9_]*)\)")


@pytest.fixture(scope="module")
def book():
    cfg = load_config("config/config.yaml")
    wb = build(cfg, cfg.output_lobs[0])
    return wb


def _is_multi_cell(wb, name: str) -> bool:
    """True when a defined name spans more than one cell."""
    if name not in wb.defined_names:
        return False
    for sheet, ref in wb.defined_names[name].destinations:
        if sheet not in wb.sheetnames:
            return False
        got = wb[sheet][ref.replace("$", "")]
        return isinstance(got, tuple)
    return False


def test_no_checks_formula_wraps_a_range_in_n(book):
    """The general form of the bug, not the two instances of it.

    Any Checks formula applying N() to a MULTI-CELL name is testing one cell
    while claiming to test a column. Single-cell names are fine — that is what
    N() is for.
    """
    ws = book["Checks"]
    offenders = []
    for row in ws.iter_rows():
        for cell in row:
            f = cell.value
            if not isinstance(f, str) or not f.startswith("="):
                continue
            for name in _N_OF_NAME.findall(f):
                if _is_multi_cell(book, name):
                    offenders.append(f"{cell.coordinate}: N({name})")
    assert not offenders, (
        "N() collapses a range to its first cell inside SUMPRODUCT, so these "
        "rows inspect one row and report on the whole column — use ISNUMBER "
        f"(which broadcasts) instead: {offenders}")


def test_the_two_rows_that_had_it_now_use_isnumber(book):
    """Pin the specific fixes, so a later edit cannot quietly revert them."""
    ws = book["Checks"]
    found = {}
    for row in ws.iter_rows():
        desc = row[2].value if len(row) > 2 else None
        if not isinstance(desc, str):
            continue
        if "M_0 as-of dates precede" in desc:
            found["asof"] = row[4].value
        elif "Target LR agrees with its indication components" in desc:
            found["target"] = row[4].value
        elif "positive weight sum" in desc:
            found["season"] = row[4].value
    assert set(found) == {"asof", "target", "season"}, sorted(found)
    assert "ISNUMBER(lr_m0asof)" in found["asof"], found["asof"]
    assert "N(lr_m0asof)" not in found["asof"]
    assert "ISNUMBER(lr_target)" in found["target"], found["target"]
    assert "N(se_sum)" not in found["season"], found["season"]
