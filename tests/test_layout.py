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
    # each table starts 4 rows after the previous one ends (caption + header gap)
    assert L.RL_HDR == L.LR_LAST + 4
    assert L.RL_FIRST == L.RL_HDR + 1
    assert L.RL_LAST == L.RL_FIRST + c.rate_log_rows - 1
    assert L.SE_HDR == L.RL_LAST + 4
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
