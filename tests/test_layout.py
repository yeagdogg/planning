"""Layout.configure arithmetic under the shipped config and synthetic rosters.

Run:  python -m pytest tests -q
"""

from __future__ import annotations

from dataclasses import replace

import pytest

from src.build_workbook import Layout as L, load_config


@pytest.fixture(scope="module")
def cfg():
    return load_config("config/config.yaml")


def _assert_layout_consistent(c):
    L.configure(c)
    n_combos = len(c.business_units) * len(c.states)
    assert L.LR_ROWS == n_combos + c.spare_lr_rows
    assert L.LR_LAST == L.LR_FIRST + L.LR_ROWS - 1
    # the rate log lives on its own sheet with a fixed header
    assert L.RL_HDR == 6
    assert L.RL_FIRST == L.RL_HDR + 1
    assert L.RL_LAST == L.RL_FIRST + c.rate_log_rows - 1
    # seasonality follows tbl_LR directly on Inputs
    assert L.SE_HDR == L.LR_LAST + 4
    assert L.SE_ROWS == len(c.states) + c.spare_seasonality_rows
    assert L.SE_LAST == L.SE_FIRST + L.SE_ROWS - 1
    # parameters live ABOVE the tables (term input is never a long scroll away)
    assert L.TERM_ROW == L.WB_HDR + 1
    assert L.LR_HDR == L.TERM_ROW + 3
    assert L.LR_FIRST == L.LR_HDR + 1
    # _calc results table mirrors tbl_LR 1:1; blocks start after it
    assert L.CALC_RES_LAST == L.CALC_RES_FIRST + L.LR_ROWS - 1
    assert L.CALC_BLOCK_FIRST == L.CALC_RES_LAST + 4
    # Portfolio grid mirrors tbl_LR 1:1
    assert L.PF_LAST == L.PF_FIRST + L.LR_ROWS - 1
    assert L.PF_W_TOTAL == L.PF_LAST + 2
    assert L.PF_S_TOTAL == L.PF_LAST + 3


def test_layout_shipped_config(cfg):
    _assert_layout_consistent(cfg)


def test_layout_synthetic_rosters(cfg):
    for bus, states in (
        (("BU-1", "BU-2", "BU-3", "BU-4"), tuple(f"S{i:02d}" for i in range(25))),
        (("Solo",), ("AA",)),
    ):
        c = replace(cfg, business_units=bus, states=states)
        _assert_layout_consistent(c)
    # restore the shipped geometry for any test that runs after this one
    L.configure(cfg)


def test_layout_configure_is_idempotent(cfg):
    L.configure(cfg)
    first = (L.LR_LAST, L.RL_LAST, L.SE_LAST, L.TERM_ROW, L.CALC_BLOCK_FIRST, L.PF_LAST)
    L.configure(cfg)
    assert first == (L.LR_LAST, L.RL_LAST, L.SE_LAST, L.TERM_ROW,
                     L.CALC_BLOCK_FIRST, L.PF_LAST)


def test_class_defaults_match_shipped_configure(cfg):
    """`python -m src.build_workbook` can load the module twice (__main__ and
    src.build_workbook). The __main__ guard now delegates to the canonical
    module, but keep the class-body defaults in lockstep with the shipped
    config anyway so any OTHER unconfigured import can never silently build a
    different geometry (the D46 dual-module trap)."""
    defaults = {k: getattr(L, k) for k in vars(L)
                if k.isupper() and isinstance(getattr(L, k), int)}
    L.configure(cfg)
    mismatched = {k: (v, getattr(L, k)) for k, v in defaults.items()
                  if getattr(L, k) != v}
    assert not mismatched, f"class defaults diverge from configure(): {mismatched}"


# ---------------------------------------------------------------------------
# D90: the State Summary column map is the single declaration of that exhibit's
# order — for the per-LOB sheet, the Book's mirror, and both harnesses.
# ---------------------------------------------------------------------------


def test_ss_column_map_is_well_formed():
    from src.sheets_main import SS_COLS, SS_COL, SS_HELP, SS_LAST
    keys = [c.key for c in SS_COLS]
    assert len(keys) == len(set(keys)), "duplicate State Summary column key"
    assert SS_LAST == len(SS_COLS) == 36
    assert sorted(SS_COL.values()) == list(range(1, SS_LAST + 1))
    # helpers live strictly to the RIGHT of the exhibit, never inside it
    assert min(SS_HELP.values()) > SS_LAST
    assert sorted(SS_HELP.values()) == list(range(SS_LAST + 1, SS_LAST + 7))


def test_ss_groups_are_contiguous_and_cover_every_column():
    from src.sheets_main import SS_COLS, ss_groups
    runs = ss_groups()
    assert runs[0][0] == 1 and runs[-1][1] == len(SS_COLS)
    for (_, prev_last, _), (nxt_first, _, _) in zip(runs, runs[1:]):
        assert nxt_first == prev_last + 1, "a group span has a hole in it"
    # a caption may not appear in two separate runs: the two-tier header writes
    # it once at the run's first column, so a split group would print it twice
    caps = [c for _f, _l, c in runs if c]
    assert len(caps) == len(set(caps))


def _ss_px(upto_key: str, collapse_at: int) -> float:
    """Pixels to the right edge of `upto_key` with every column at outline level
    >= collapse_at hidden. Excel renders a column of width w at roughly 7w+5px;
    the roster column is frozen, so this is what a reader actually sees."""
    from src.sheets_main import SS_COLS, ss_c
    return sum(7 * c.width + 5 for c in SS_COLS[:ss_c(upto_key)]
               if c.level < collapse_at) + 30


def test_the_headline_is_on_the_first_screen_in_every_collapse_state():
    """D90 found Plan LR stranded at ~1,600px and fixed it by moving the bridge
    ahead of the inputs. D101 puts the inputs back in front — the exhibit reads
    as a derivation, not a dashboard — and solves the reach problem by COLLAPSING
    them instead, which is what the outline levels are for. So the property to
    pin is no longer raw column order; it is that the headline is reachable in
    the state the file actually SHIPS in, and closer still when collapsed."""
    from src.sheets_main import SS_LV_HIST, SS_LV_INPUT
    as_shipped = _ss_px("planlr", SS_LV_HIST)      # chronology collapsed
    assert as_shipped < 1100, f"plan LR sits {as_shipped:.0f}px in — off screen"
    assert _ss_px("target", SS_LV_HIST) < 1100     # and the benchmark beside it
    # collapsing the whole input region has to be worth doing
    answer_only = _ss_px("planlr", SS_LV_INPUT)
    assert answer_only < 600, f"inputs collapsed still puts plan LR at {answer_only:.0f}px"
    assert answer_only < as_shipped - 300
    # fully expanded it is a wide exhibit, and that is the point of the toggle
    assert _ss_px("planlr", 99) > 1400


def test_the_exhibit_reads_as_a_derivation():
    """Inputs, then the walk, then the answer — and the following year's answer
    beside this year's rather than 20 columns away (D101)."""
    from src.sheets_main import SS_LV_HIST, SS_LV_INPUT, ss_c
    # the bridge itself reads left to right, ending in the number it produces
    chain = [ss_c(k) for k in ("lrcur", "arate", "amod", "aother", "planlr")]
    assert chain == list(range(chain[0], chain[0] + 5))
    # then the benchmark, immediately beside the answer (D96), then the mix
    # residual — target is NOT part of the product, only read against it
    assert ss_c("target") == ss_c("planlr") + 1
    assert ss_c("mix") == ss_c("target") + 1
    # every input feeding the bridge precedes it, and is collapsible
    from src.sheets_main import SS_COLS
    lv = {c.key: c.level for c in SS_COLS}
    for k in ("chg1_date", "ntaken", "mind", "m0", "m1", "crl", "ecy", "mbar"):
        assert ss_c(k) < ss_c("lrcur"), f"{k} is an input and belongs before the bridge"
        assert lv[k] >= SS_LV_INPUT, f"{k} is an input and must be collapsible"
    # the chronology collapses independently of the rest, so it is nested deeper
    assert lv["chg1_date"] == SS_LV_HIST > SS_LV_INPUT == lv["mind"]
    # P and P+1 are the same quantity for two years: keep them within one screen
    assert ss_c("planlr1") - ss_c("planlr") <= 8, "the two years drifted apart again"
    # nothing the reader came for may be behind a collapse
    answers = ("lrcur", "arate", "amod", "aother", "planlr", "target", "mix",
               "trend", "arate1", "amod1", "planlr1")
    hidden = [k for k in answers if lv[k] != 0]
    assert not hidden, f"answer columns are behind a collapse: {hidden}"


def test_book_mirror_shares_the_column_map():
    """The Book's State Summary is the same exhibit one dimension deeper. It
    used to carry its own hand-maintained copy of the 35-column order, which is
    exactly the arrangement that lets one drift from the other."""
    import src.sheets_book as bk
    import src.sheets_main as sm
    assert bk.SS_COLS is sm.SS_COLS
    assert bk.ss_l("planlr") == sm.ss_l("planlr")


def test_both_state_summaries_get_the_same_conditional_formats():
    """D104. The Book's mirror had NO conditional formatting for four versions:
    no plan-LR heatmap, no EP data bar, and no amber on planned rate changes.
    The two builders had already been unified on the column map (D90), the
    group captions and the outline levels (D101) — formatting was the last
    thing still hand-written in one place, so it was the last thing to drift.

    Both now call ss_conditional_formats, so pin what it produces rather than
    that it was called: the rules must land on the columns the map names, in
    any collapse state and after any future reorder."""
    from openpyxl import Workbook

    from src.sheets_main import ss_conditional_formats, ss_l

    ws = Workbook().active
    ss_conditional_formats(ws, 8, 31, "Calibri")
    got = {}
    for rng in ws.conditional_formatting:
        cols = {"".join(ch for ch in part.split("$")[-2] if ch.isalpha())
                if "$" in part else "".join(ch for ch in part if ch.isalpha())
                for part in str(rng.sqref).replace(":", " ").split()}
        for rule in rng.rules:
            for c in cols:
                got.setdefault(c, set()).add(rule.type)
    assert got.get(ss_l("planlr")) == {"colorScale"}
    assert got.get(ss_l("planlr1")) == {"colorScale"}
    assert got.get(ss_l("ep")) == {"dataBar"}
    for j in range(1, 5):
        assert got.get(ss_l(f"chg{j}_tok")) == {"expression"}, f"chg{j}_tok unformatted"
    # nothing lands anywhere else — a stray rule means a literal column letter
    expected = {ss_l(k) for k in ("planlr", "planlr1", "ep")} | {
        ss_l(f"chg{j}_tok") for j in range(1, 5)}
    assert set(got) == expected, f"unexpected columns formatted: {set(got) - expected}"


def test_the_token_rule_anchors_on_its_own_range():
    """The subtle half of the same bug: a formula whose anchor row does not
    match the range start makes every row test its neighbour's cell."""
    from openpyxl import Workbook

    from src.sheets_main import ss_conditional_formats

    ws = Workbook().active
    ss_conditional_formats(ws, 8, 31, "Calibri")
    n = 0
    for rng in ws.conditional_formatting:
        for rule in rng.rules:
            if rule.type != "expression":
                continue
            n += 1
            ref = str(rng.sqref).split(":")[0]          # e.g. E8
            col_ = "".join(ch for ch in ref if ch.isalpha())
            row_ = "".join(ch for ch in ref if ch.isdigit())
            assert rule.formula == [f'LEFT(${col_}{row_},1)="P"'], rule.formula
    assert n == 4, f"expected 4 token rules, found {n}"
