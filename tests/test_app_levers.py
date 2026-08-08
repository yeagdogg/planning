"""W2c: the State Summary's live levers and the undo stack.

The undo unit tests are UI-free (both interpreters). The router tests
drive the REAL page through synthetic edit events (the pattern the P4
click-through tests set), venv-gated: every gate must reject loudly AND
revert the cell; every accepted gesture must move the engine numbers in
the patched frame and be undoable to the exact prior state.
"""
from __future__ import annotations

import datetime as dt
import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"


# ------------------------------------------------------------- undo (pure)

class _Bus:
    def __init__(self):
        self.rev = 0

    def bump(self):
        self.rev += 1


def _sess(scns):
    page = SimpleNamespace(config=scns[0], data_version=0, results={},
                           caches={}, book={s.lob: s for s in scns})
    return SimpleNamespace(page=page, bus=_Bus(),
                           ctx=SimpleNamespace(results_rev=0))


def test_undo_restores_by_content_match_and_caps():
    from app import importers
    from app.undo import UndoStack, entry
    sess = _sess([importers.from_workbook(WB_PROP)])
    scn = sess.page.book["Property"]
    row = scn.lr_rows[0]
    old = row.get("netp")

    undo = UndoStack(cap=3)
    e = entry("Property", "lr_rows", row, {"netp": old},
              after={"netp": 0.02})
    row["netp"] = 0.02
    undo.push("test edit", [e])
    assert undo.peek() == "test edit" and len(undo) == 1

    n, skipped = undo.pop_apply(sess)
    assert n == 1 and not skipped
    assert row.get("netp") == old                  # exact restore, incl None
    assert sess.bus.rev == 1
    assert undo.peek() is None

    for i in range(5):                             # cap: only the last 3 kept
        undo.push(f"edit {i}", [e])
    assert len(undo) == 3 and undo.peek() == "edit 4"
    undo.clear()
    assert len(undo) == 0


def test_undo_skips_a_vanished_row_loudly():
    from app import importers
    from app.undo import UndoStack, entry
    sess = _sess([importers.from_workbook(WB_PROP)])
    scn = sess.page.book["Property"]
    row = scn.lr_rows[0]
    undo = UndoStack()
    e = entry("Property", "lr_rows", row, {"netp": row.get("netp")},
              after={"netp": 0.02})
    row["netp"] = 0.02
    undo.push("edit", [e])
    # an Inputs-style wholesale replacement changes the row's content
    scn.lr_rows[0] = dict(row, lr_proj=0.99)
    n, skipped = undo.pop_apply(sess)
    assert n == 0
    assert skipped and "changed since the edit" in skipped[0]


# --------------------------------------------------- the router (venv page)

needs_panel = pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")


def _page_session(bus_filter=None):
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page
    session = PlanSession()
    page = ss_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    if bus_filter:
        page["filters"]["bu"].value = bus_filter
    return session, page


def _ev(page, state, column, value):
    df = page["table"].value
    row = df.index[df["state"] == state][0]
    return SimpleNamespace(column=column, row=int(row), value=value,
                           old=df.iloc[int(row)][column])


@needs_panel
def test_net_fanout_honors_the_view_and_undo_restores():
    session, page = _page_session()
    meta = page["holder"]["meta"]
    st = next(s for s in meta.states if meta.cnt[s] > 1)
    combos = meta.view[st]
    olds = {c[2]: c[4].get("netp") for c in combos}
    rev0 = session.bus.rev

    page["edit_router"](_ev(page, st, "netsel", 0.02))
    scn = session.page.book["Property"]
    for _lob, _s, key, _bu, row, _res in combos:
        assert row["netp"] == 0.02, key
    assert session.bus.rev == rev0 + 1
    assert "applied to" in page["flash"].object
    assert not page["undo_btn"].disabled

    # the recompute reached the patched frame: the state is net-flagged now
    df = page["table"].value
    got = df.loc[df["state"] == st, "netsel"].iloc[0]
    assert got == pytest.approx(0.02)

    n, skipped = page["undo"].pop_apply(session)
    assert n == len([k for k, v in olds.items()]) and not skipped
    for _lob, _s, key, _bu, row, _res in combos:
        assert row.get("netp") == olds[key], key   # None comes back as None
    assert scn is session.page.book["Property"]


@needs_panel
def test_blank_netsel_clears_net_mode_loudly():
    session, page = _page_session()
    meta = page["holder"]["meta"]
    # find a state carrying a real net combo in the sample
    st = next(s for s in meta.states if meta.netcnt[s] > 0)
    page["edit_router"](_ev(page, st, "netsel", ""))
    assert "flips those combos out of net mode" in page["flash"].object
    for _lob, _s, _k, _bu, row, _res in meta.view[st]:
        assert row.get("netp") is None
    n, _sk = page["undo"].pop_apply(session)
    assert n >= 1                                   # and back again
    meta2 = page["holder"]["meta"]
    assert meta2.netcnt[st] > 0


@needs_panel
def test_gates_reject_and_revert():
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]

    # taken slot: locked as historical fact
    st, j = next(((s, j) for (s, j), r in meta.slot_map.items()
                  if r.get("status") == "taken"), (None, None))
    assert st is not None
    before = page["table"].value
    cell0 = before.loc[before["state"] == st, f"chg{j}_pct"].iloc[0]
    ev = _ev(page, st, f"chg{j}_pct", "9%")
    page["edit_router"](ev)
    assert "historical fact" in page["flash"].object
    after = page["table"].value
    assert after.loc[after["state"] == st, f"chg{j}_pct"].iloc[0] == cell0
    assert page["undo"].peek() is None              # nothing armed

    # TOTAL row: aggregate
    page["edit_router"](_ev(page, "TOTAL", "netsel", 0.05))
    assert "TOTAL row is an aggregate" in page["flash"].object

    # multi-combo slot edit: needs a single-combo view
    session2, page2 = _page_session()
    meta2 = page2["holder"]["meta"]
    multi = next(s for s in meta2.states if meta2.cnt[s] > 1)
    page2["edit_router"](_ev(page2, multi, "chg1_pct", "5%"))
    assert "single combo" in page2["flash"].object


@needs_panel
def test_planned_pct_edit_moves_planlr_and_ties_the_engine():
    from src import engine

    from app import compute
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]
    st, j = next(((s, j) for (s, j), r in meta.slot_map.items()
                  if r.get("status") == "planned"), (None, None))
    assert st is not None, "sample should hold a planned filing"
    df0 = page["table"].value
    lr_before = df0.loc[df0["state"] == st, "planlr"].iloc[0]

    page["edit_router"](_ev(page, st, f"chg{j}_pct", "9%"))
    assert "Filed %" in page["flash"].object

    df1 = page["table"].value
    row1 = df1.loc[df1["state"] == st].iloc[0]
    assert row1["planlr"] != pytest.approx(lr_before, abs=1e-9)
    assert row1[f"chg{j}_pct"].startswith("+")      # re-rendered filed×ach

    lob, scn, key = meta.single[st]
    res = engine.run_bridge(scn.plan_year,
                            compute.combo_inputs(scn, scn.row(key)),
                            "monthly")
    assert row1["planlr"] == pytest.approx(res.cy_lr_p, rel=1e-12)

    page["undo"].pop_apply(session)
    df2 = page["table"].value
    assert df2.loc[df2["state"] == st, "planlr"].iloc[0] == \
        pytest.approx(lr_before, rel=1e-12)


@needs_panel
def test_slot_date_mod_step_and_achievement_round_trip():
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]

    st, j = next(((s, j) for (s, j), r in meta.slot_map.items()
                  if r.get("status") == "planned"), (None, None))
    r = meta.slot_map[(st, j)]
    old_eff = r["eff"]
    page["edit_router"](_ev(page, st, f"chg{j}_date", "2027-06-15"))
    assert r["eff"] == dt.date(2027, 6, 15)
    page["undo"].pop_apply(session)
    assert meta_row_eff(page, st, j) == old_eff

    # achievement on the next planned filing
    st2 = next((s for s in page["holder"]["meta"].achnext_map), None)
    if st2 is not None:
        m = page["holder"]["meta"]
        _lob, rr = m.achnext_map[st2]
        page["edit_router"](_ev(page, st2, "ach_next", "60%"))
        assert rr["achievement"] == pytest.approx(0.60)
        page["undo"].pop_apply(session)

    # mod step (if the filtered sample carries a planned mod action)
    m3 = page["holder"]["meta"]
    st3 = next((s for s in m3.modstep_map), None)
    if st3 is not None:
        _lob, mr = m3.modstep_map[st3]
        old = mr.get("chg")
        page["edit_router"](_ev(page, st3, "modstep_pct", "-3%"))
        assert mr["chg"] == pytest.approx(-0.03)
        page["undo"].pop_apply(session)
        assert mr.get("chg") == old


def meta_row_eff(page, st, j):
    return page["holder"]["meta"].slot_map[(st, j)]["eff"]


@needs_panel
def test_undo_clears_on_scenario_load():
    from app import importers
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]
    st = meta.states[0]
    page["edit_router"](_ev(page, st, "netsel", 0.03))
    assert not page["undo_btn"].disabled
    session.replace_config(importers.from_workbook(WB_PROP))
    assert page["undo_btn"].disabled
    assert page["undo"].peek() is None


# --------------------------------------------------------- W3c impact strip

@needs_panel
def test_impact_strip_ties_the_recompute():
    """One snapshot per gesture, 1:1 with undo: before == the pre-edit
    frame, after == the fresh engine frame at 1e-12, and the strip pane
    renders the movement."""
    import pytest as _pt

    from app import summary
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]
    st = meta.states[0]
    df0 = page["table"].value
    b_row = float(df0.loc[df0["state"] == st, "planlr"].iloc[0])
    b_tot = float(df0.iloc[-1]["planlr"])

    page["edit_router"](_ev(page, st, "netsel", 0.05))
    stack = page["impact"]["stack"]
    # headless debounce is synchronous, so the bump inside the router has
    # already rendered and filled `after`; in the browser it fills ~250ms
    # later — either way `before` is the PRE-edit frame
    assert len(stack) == 1
    assert stack[-1]["before"][0] == _pt.approx(b_row, abs=1e-12)
    assert stack[-1]["before"][2] == _pt.approx(b_tot, abs=1e-12)

    page["on_show"]()                       # idempotent re-render
    snap = stack[-1]
    assert snap["after"] not in (None, "stale")
    fresh, _m = summary.state_frame(session, bus=["ABD"])
    want_row = float(fresh.loc[fresh["state"] == st, "planlr"].iloc[0])
    want_tot = float(fresh.iloc[-1]["planlr"])
    assert snap["after"][0] == _pt.approx(want_row, abs=1e-12)
    assert snap["after"][2] == _pt.approx(want_tot, abs=1e-12)
    html = page["impact"]["pane"].object
    assert "Last change" in html and st in html and "TOTAL" in html


@needs_panel
def test_impact_strip_pops_with_undo_and_clears_on_load():
    from app import importers
    session, page = _page_session(bus_filter=["ABD"])
    meta = page["holder"]["meta"]
    st = meta.states[0]
    page["edit_router"](_ev(page, st, "netsel", 0.04))
    page["on_show"]()
    assert page["impact"]["pane"].object != ""
    page["undo_btn"].clicks += 1            # pop_apply bumps the bus
    page["on_show"]()
    assert page["impact"]["stack"] == []
    assert page["impact"]["pane"].object == ""

    page["edit_router"](_ev(page, st, "netsel", 0.06))
    assert len(page["impact"]["stack"]) == 1
    session.replace_config(importers.from_workbook(WB_PROP))
    assert page["impact"]["stack"] == []
    assert page["impact"]["pane"].object == ""


# ------------------------------------------------------ W3h live punch list

@needs_panel
def test_click_into_an_empty_cell_is_not_a_gesture():
    """John's live repro: clicking into an EMPTY Net rate sel cell
    committed the editor's empty value — before editorEmptyValue that
    was 0, silently writing net mode ON at 0% to every combo in the
    row's view. The columnDefaults fix makes the commit None; the
    router treats blank -> blank as no gesture: no fan-out, no flash,
    no bump."""
    import math
    from types import SimpleNamespace as NS
    session, page = _page_session()
    meta = page["holder"]["meta"]
    st = next(s for s in meta.states
              if all(c[4].get("netp") is None for c in meta.view[s]))
    rev0, flash0 = session.bus.rev, page["flash"].object
    df = page["table"].value
    row = int(df.index[df["state"] == st][0])
    old = df.iloc[row]["netsel"]
    assert old is None or (isinstance(old, float) and math.isnan(old))
    for committed in (None, "", float("nan")):
        page["edit_router"](NS(column="netsel", row=row,
                               value=committed, old=old))
    assert session.bus.rev == rev0                 # nothing happened
    assert page["flash"].object == flash0
    assert all(c[4].get("netp") is None for c in meta.view[st])
    assert page["undo_btn"].disabled               # nothing to undo
    # and the client half of the fix ships in the configuration
    cfg = page["table"]._configuration
    assert cfg["columnDefaults"]["editorEmptyValue"] is None


@needs_panel
def test_a_lever_on_a_non_active_line_moves_plan_lr():
    """The stale-cache bug distilled at page level (W3h): with two
    lines loaded and Property ACTIVE, a net lever on a General
    Liability row must move that row's plan LR in the patched frame —
    active-line-only cache invalidation served the stale line until
    something moved data_version (John: "inputs on state summary don't
    seem to affect the plan lr")."""
    import math
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page
    session = PlanSession()
    page = ss_page.build(session)
    session.replace_config(importers.from_workbook(WB_GL))
    session.replace_config(importers.from_workbook(WB_PROP))
    assert session.page.config.lob == "Property"   # GL is NOT active
    page["filters"]["line"].value = ["General Liability"]

    df0 = page["table"].value
    body = df0[df0["state"] != "TOTAL"]
    row0 = next(r for _i, r in body.iterrows()
                if isinstance(r["planlr"], float)
                and not math.isnan(r["planlr"]))
    st, before = row0["state"], float(row0["planlr"])

    page["edit_router"](_ev(page, st, "netsel", 0.15))
    assert session.page.config.lob == "Property"   # still not active
    df1 = page["table"].value
    after = float(df1.loc[df1["state"] == st, "planlr"].iloc[0])
    assert abs(after - before) > 1e-6, \
        "the non-active line's plan LR must move without a data_version bump"
