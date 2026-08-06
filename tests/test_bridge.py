"""The Bridge's resolver block: companion notes must sit on the row they
describe, and the block must not grow into the chart staging area below it.

Both are D98. The companion notes used to be keyed by POSITION in the resolver
input list, so inserting the target loss ratio (D96) at index 12 slid every
later note one row down from its label. The A_other warning printed beside "Net
trend", and the net-selection note landed on "Mod adjustment effective" — a
cell holding the text "ON"/"OFF", which makes IF(<text>,...) a permanent
#VALUE!. It shipped in all six workbooks; the Checks panel still said PASS.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

import re

import pytest

from src.build_workbook import Layout as L, build, load_config


@pytest.fixture(scope="module")
def bridge():
    cfg = load_config("config/config.yaml")
    wb = build(cfg, cfg.output_lobs[0])
    return wb, wb["Bridge"]


def _row_of(wb, name: str) -> int:
    """Row of a single-cell defined name on the Bridge."""
    dn = wb.defined_names[name]
    dest = list(dn.destinations)
    assert len(dest) == 1, f"{name} is not a single range"
    sheet, ref = dest[0]
    assert sheet == "Bridge", f"{name} lives on {sheet}, not Bridge"
    return int(re.search(r"\$C\$(\d+)", ref).group(1))


# the note that belongs to each resolver input, by a phrase unique to it
# (nr_AOther's note went with D107: it warned that A_other <> 1 needs a label,
# and the label column is gone.)
COMPANIONS = {
    "nr_MPrior": "backward anchor",
    "nr_M2": "carried flat beyond the plan year",
    "nr_NetMode": "NET SELECTION ACTIVE",
}


@pytest.mark.parametrize("name,phrase", sorted(COMPANIONS.items()))
def test_each_companion_note_sits_on_the_row_it_describes(bridge, name, phrase):
    wb, ws = bridge
    r = _row_of(wb, name)
    got = ws[f"D{r}"].value
    assert got and phrase in got, (
        f"the note for {name} (row {r}) reads {got!r} — a companion note has "
        f"drifted off its label again")


def test_no_advisory_lands_on_a_resolver_row_that_should_not_have_one(bridge):
    """The same drift seen from the other side.

    Checking only that each named row HAS its note would pass a workbook that
    also grew a stray note somewhere else; checking the set catches the note
    that slid onto its neighbour. (A self-reference check would be vacuous
    here — the template is formatted with the row it is written to, so it
    always reads its own row, even when that row is the wrong one.)"""
    wb, ws = bridge
    expected = {_row_of(wb, n) for n in COMPANIONS}
    found = {r for r in range(L.BR_IN_FIRST, L.BR_CHART_DATA - 1)
             if isinstance(ws[f"D{r}"].value, str)
             and ws[f"D{r}"].value.startswith("=")}
    assert found == expected, (
        f"advisory notes sit on rows {sorted(found)}, but the inputs they "
        f"describe are on {sorted(expected)}")


def test_the_resolver_block_does_not_reach_the_chart_staging_rows(bridge):
    """The block is fixed-origin and so is the chart data below it, so it can
    only grow into it — silently, by overwriting the waterfall's own series."""
    wb, _ws = bridge
    last = max(_row_of(wb, n) for n in wb.defined_names
               if n.startswith("nr_")
               and "Bridge" in str(list(wb.defined_names[n].destinations))
               and re.search(r"\$C\$\d+$", str(list(
                   wb.defined_names[n].destinations)[0][1])))
    assert last < L.BR_CHART_DATA - 1, (
        f"the resolver block now ends at row {last}, and the waterfall chart "
        f"data starts at {L.BR_CHART_DATA - 1}")
