"""P2: paste coercion, grid round-trips, validators, and the live edit loop.

The paste/grid/validate layers are UI-free, so most of this runs on both
interpreters; the edit-loop test at the bottom (build the Inputs page, patch
a grid frame, watch the engine answer move) is venv-only.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"


# ------------------------------------------------------------------ paste

def test_coerce_accepts_what_excel_displays():
    from app.paste import coerce
    assert coerce("pct", "10.3%") == pytest.approx(0.103)
    assert coerce("pct", "0.103") == pytest.approx(0.103)
    assert coerce("pct", "(2.0%)") == pytest.approx(-0.02)
    assert coerce("dollar", "36,000") == 36000.0
    assert coerce("num", "1.1") == 1.1
    assert coerce("date", "3/31/2026") == dt.date(2026, 3, 31)
    assert coerce("date", "2026-03-31") == dt.date(2026, 3, 31)
    assert coerce("text", "  ABD ") == "ABD"
    assert coerce("choice:taken,planned", "taken") == "taken"
    for kind in ("pct", "date", "text", "num"):
        assert coerce(kind, None) is None
        assert coerce(kind, "") is None


def test_apply_block_names_the_bad_cell():
    from app.paste import apply_block
    from app.grids import RL_SCHEMA
    good = "ABD\tAZ\t3/31/2026\t10.3%\ttaken\tY\t\t"
    rows = apply_block(RL_SCHEMA, good)
    assert rows[0]["eff"] == dt.date(2026, 3, 31)
    assert rows[0]["filed"] == pytest.approx(0.103)
    assert rows[0]["achievement"] is None
    with pytest.raises(ValueError, match="column 'Effective'"):
        apply_block(RL_SCHEMA, "ABD\tAZ\tnot-a-date\t10%\ttaken")


# ------------------------------------------------------- df/rows round trip

def test_lr_rows_survive_the_grid_round_trip():
    from app import importers
    from app.grids import LR_SCHEMA, df_to_rows, rows_to_df
    scn = importers.from_workbook(WB)
    back = df_to_rows(LR_SCHEMA, rows_to_df(LR_SCHEMA, scn.lr_rows))
    assert back == scn.lr_rows          # spares dropped, values intact


def test_season_adapters_round_trip():
    from app.grids import grid_to_season, season_to_grid
    rows = [{"state": "AZ", "weights": [1.0] * 6 + [2.0] * 6}]
    assert grid_to_season(season_to_grid(rows)) == rows


# -------------------------------------------------------------- validators

def test_validators_flag_the_planted_problems():
    import dataclasses

    from app import importers
    from app.validate import validate
    scn = importers.from_workbook(WB)
    base = validate(scn)
    # the shipped sample seeds ONE deliberate target-identity WARN (D107)
    assert any("Target LR" in m for sev, m in base if sev == "WARN")
    assert not any(sev == "FAIL" for sev, _m in base)

    rows = [dict(r) for r in scn.lr_rows]
    rows[0] = dict(rows[0], netp=1.5)               # out of range
    rows.append(dict(rows[1]))                       # duplicate key
    rl = [dict(r) for r in scn.rate_rows]
    rl[0] = dict(rl[0], status="maybe")             # bad status
    bad = dataclasses.replace(scn, lr_rows=rows, rate_rows=rl)
    msgs = validate(bad)
    fails = [m for sev, m in msgs if sev == "FAIL"]
    assert any("outside" in m for m in fails)
    assert any("duplicate" in m for m in fails)
    assert any("taken/planned" in m for m in fails)


# ------------------------------------------------------- live edit loop (UI)

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_edit_a_rate_row_and_the_bridge_moves():
    from app import compute, importers
    from app.glue.session import PlanSession
    from app.pages import inputs as inputs_page

    session = PlanSession()
    page = inputs_page.build(session)
    scn = importers.from_workbook(WB)
    session.replace_config(scn)                     # reseeds the grids

    key = scn.combo_keys()[0]
    bu, state = key.split("|")
    before = compute.results(_snap(session))[key].cy_lr_p

    # NOTE the first attempt at this test edited the combo's OLDEST change
    # and the bridge did not move — correctly: a fully-earned considered
    # change scales CRL_ind and E_CY together and cancels out of A_rate
    # (that is what on-leveling is). Type a NEW plan-year, not-considered
    # change into a spare row instead: pure earn-in, guaranteed to move.
    grid = page["tables"]["Rate Log"]
    df = grid.value.copy()
    spare = df.index[df["bu"].isna() | (df["bu"].isnull())]
    row = spare[0]
    df.loc[row, ["bu", "state", "eff", "filed", "status"]] = \
        [bu, state, f"{scn.plan_year}-04-01", 0.10, "taken"]
    n_before = len(scn.rate_rows)
    rev0 = session.bus.rev
    grid.value = df                                 # -> write-back

    assert session.bus.rev == rev0 + 1
    assert len(session.page.config.rate_rows) == n_before + 1
    after = compute.results(_snap(session))[key].cy_lr_p
    assert abs(after - before) > 1e-4, "bridge did not move"

    # the validation rail sees a planted FAIL through the same path
    df2 = grid.value.copy()
    df2.loc[row, "status"] = "maybe"
    grid.value = df2
    import app.validate as v
    assert any(sev == "FAIL" for sev, _m in v.validate(session.page.config))


def _snap(session):
    """compute.results wants only the plain attribute surface."""
    return SimpleNamespace(page=session.page, bus=session.bus,
                           ctx=SimpleNamespace(results_rev=0))
