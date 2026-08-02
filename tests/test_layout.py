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
    assert SS_LAST == len(SS_COLS) == 35
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


def test_the_bridge_and_its_headline_fit_the_first_screen():
    """The reason the exhibit was reordered (D90).

    Plan LR used to sit at column 27, about 1,600px in — off the right edge of
    a 1080p window, so the flagship leadership exhibit did not show its own
    headline without scrolling. Excel renders a column of width w at roughly
    7w + 5 px; the roster column is frozen, so everything up to and including
    Plan LR must fit."""
    from src.sheets_main import SS_COLS, ss_c
    px = sum(7 * c.width + 5 for c in SS_COLS[:ss_c("planlr")]) + 30
    assert px < 1100, f"plan LR sits {px:.0f}px in — off screen again"
    # and the bridge must read left to right, ending in the number it produces
    order = [ss_c(k) for k in ("lrcur", "arate", "amod", "aother", "planlr", "mix")]
    assert order == sorted(order) and order == list(range(order[0], order[0] + 6))
    # the chronology is reference, and follows the answer
    assert ss_c("chg1_date") > ss_c("planlr")


def test_book_mirror_shares_the_column_map():
    """The Book's State Summary is the same exhibit one dimension deeper. It
    used to carry its own hand-maintained copy of the 35-column order, which is
    exactly the arrangement that lets one drift from the other."""
    import src.sheets_book as bk
    import src.sheets_main as sm
    assert bk.SS_COLS is sm.SS_COLS
    assert bk.ss_l("planlr") == sm.ss_l("planlr")
