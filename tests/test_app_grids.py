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


def test_apply_block_skips_a_copied_header_row():
    from app.paste import apply_block
    from app.grids import RL_SCHEMA
    text = ("BU\tState\tEffective\tFiled %\tStatus\tConsidered"
            "\tAchievement\tComment\n"
            "ABD\tAZ\t3/31/2026\t10.3%\ttaken\tY\t\t")
    rows = apply_block(RL_SCHEMA, text)
    assert len(rows) == 1 and rows[0]["bu"] == "ABD"


def test_apply_block_skips_the_workbooks_own_headers():
    """John's live repro: the WORKBOOK's headers ('Adj plan EP (000s)' etc.)
    share no vocabulary with the app's titles — detection must be
    content-based, not title-matching."""
    from app.grids import LR_SCHEMA
    from app.paste import apply_block
    hdr = "\t".join(["Business unit", "State", "Adj plan EP (000s)",
                     "Projected loss ratio (current rate level)",
                     "Prospective premium trend", "Prospective loss trend",
                     "Expense ratio", "ALAE factor", "ULAE factor",
                     "Combined ratio", "Target loss ratio", "Cat load",
                     "Large loss load", "Mod assumed in indication",
                     "Mod ~1yr before as-of", "Current avg written mod",
                     "Current mod as-of date", "Projected mod end current yr",
                     "Projected mod end plan yr", "Projected mod end plan+1",
                     "Net trend", "A_other", "Mod adj", "Net rate P",
                     "Net rate P+1"])
    data = ("ABD\tAZ\t36,000\t65.0%\t3.0%\t4.0%\t28.0%\t1.10\t1.05\t95.0%"
            "\t58.0%\t2.0%\t1.0%\t1.008\t\t1.052\t9/30/2026\t\t1.052\t"
            "\t\t1.0\t\t\t")
    notes: list = []
    rows = apply_block(LR_SCHEMA, hdr + "\n" + data, notes=notes)
    assert notes == ["header row skipped"]
    assert len(rows) == 1
    assert rows[0]["ep"] == 36000.0
    assert rows[0]["m0_asof"] == dt.date(2026, 9, 30)

    # a DATA row with a couple of typos is NOT mistaken for a header —
    # it errors loudly, naming the cell
    bad = data.replace("36,000", "thirty-six")
    with pytest.raises(ValueError, match="Plan EP"):
        apply_block(LR_SCHEMA, bad)


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


# ----------------------------------------------------- blank integrity (W2a)

def test_empty_value_columns_cover_numerics_and_dates_only():
    """The zero-commit fix: every column whose blank means 'no value' gets
    editorEmptyValue None; choice/text columns — where '' is a legitimate
    value — must NOT (a table-wide rule would rewrite them to None)."""
    from app.grids import LR_SCHEMA, RL_SCHEMA, _empty_value_columns
    lr = {c["field"]: c for c in _empty_value_columns(LR_SCHEMA)}
    assert all(c["editorEmptyValue"] is None for c in lr.values())
    assert {"netp", "netp1", "ep", "m0_asof", "trend"} <= set(lr)
    assert "bu" not in lr and "state" not in lr and "modadj" not in lr
    rl = {c["field"] for c in _empty_value_columns(RL_SCHEMA)}
    assert "eff" in rl and "filed" in rl
    assert "status" not in rl and "considered" not in rl and "comment" not in rl


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("bokeh") is None,
    reason="bokeh not installed (system interpreter — app venv runs this)")
def test_blank_cells_render_blank_not_dash_or_zero():
    from app.grids import LR_SCHEMA, _formatters
    fmt = _formatters(LR_SCHEMA)
    assert fmt["netp"].null_format == "" and fmt["netp"].nan_format == ""


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_grid_configuration_carries_the_empty_value_fix():
    from app.grids import LR_SCHEMA, make_grid
    grid = make_grid(LR_SCHEMA, [])
    # the constructor kwarg is popped into the private _configuration attr
    # (panel/widgets/tables.py Tabulator.__init__) — no public param exists
    cols = {c["field"]: c for c in grid._configuration["columns"]}
    assert cols["netp"]["editorEmptyValue"] is None
    assert "modadj" not in cols
    assert grid._configuration["clipboard"] == "copy"


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


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_paste_card_button_flow():
    """The user's actual motion: paste into the box (value_input — the
    keystroke-synced param, since ``value`` only syncs on blur and can race
    the click), press Apply, rows land, feedback shows. Plus the two
    failure paths that must never be silent."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import inputs as inputs_page

    session = PlanSession()
    page = inputs_page.build(session)
    ta, btn, flash = page["paste"]["Rate Log"]

    # no scenario yet -> loud, not silent
    ta.value_input = "ABD\tAZ\t3/31/2026\t10.3%\ttaken"
    btn.clicks += 1
    assert "No scenario open" in flash.object

    session.replace_config(importers.from_workbook(WB))
    scn = session.page.config

    # empty box -> loud, not silent
    ta.value_input = ""
    btn.clicks += 1
    assert "Nothing to paste" in flash.object

    # the real motion: header row + one data row, via value_input only
    ta.value_input = ("BU\tState\tEffective\tFiled %\tStatus\tConsidered"
                      "\tAchievement\tComment\n"
                      f"ABD\tAZ\t{scn.plan_year}-05-01\t12.0%\ttaken\t\t\t")
    rev0 = session.bus.rev
    btn.clicks += 1
    assert "Applied" in flash.object and "1 row" in flash.object
    assert session.bus.rev == rev0 + 1
    assert len(scn.rate_rows) == 1                 # paste REPLACES the table
    assert scn.rate_rows[0]["filed"] == pytest.approx(0.12)
    assert scn.rate_rows[0]["eff"] == dt.date(scn.plan_year, 5, 1)
    assert ta.value_input == "" and ta.value == ""  # box cleared after apply
