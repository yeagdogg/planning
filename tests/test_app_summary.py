"""W2b: the State Summary frame — every rule tied back to the engine.

The pivot-check pattern throughout: re-aggregate independently and land on
the frame's numbers. Frames run on both interpreters; the page test at
the bottom needs the app venv. Every deliberate divergence from the Excel
exhibit (Book-dialect zero-EP, the target-dilution fix, slot dash
semantics) is pinned HERE so it reads as a decision, not an accident.
"""
from __future__ import annotations

import datetime as dt
import math
import sys
from pathlib import Path
from types import SimpleNamespace

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"


def _stub(scns):
    page = SimpleNamespace(config=scns[0], data_version=0, results={},
                           caches={}, book={s.lob: s for s in scns})
    return SimpleNamespace(page=page, bus=SimpleNamespace(rev=0),
                           ctx=SimpleNamespace(results_rev=0))


@pytest.fixture(scope="module")
def sess():
    from app import importers
    return _stub([importers.from_workbook(WB_PROP),
                  importers.from_workbook(WB_GL)])


@pytest.fixture(scope="module")
def frame_meta(sess):
    from app import summary
    return summary.state_frame(sess)


# ----------------------------------------------------------- the SS_COLS tie

def test_headers_order_and_groups_come_from_SS_COLS(frame_meta):
    """The D90 drift guard: the frame IS the declaration. A workbook column
    change lands here as a loud failure, never a silent shift."""
    from src.sheets_main import SS_COLS

    from app import summary
    df, meta = frame_meta
    assert list(df.columns) == summary.ALL_KEYS
    assert summary.ALL_KEYS[:len(SS_COLS)] == [c.key for c in SS_COLS]
    assert set(summary.SS_KINDS) <= set(summary.ALL_KEYS)
    assert set(summary.live_headers(meta.plan_year)) <= set(summary.SS_KEYS)
    grouped = [k for caption, keys in
               summary.ss_groups_map(meta.plan_year).items() for k in keys]
    assert set(grouped) <= set(summary.ALL_KEYS)
    hidden = summary.hidden_sets()
    assert set(hidden["hist"]) == {c.key for c in SS_COLS if c.level >= 2}
    assert set(hidden["inputs"]) == {c.key for c in SS_COLS if c.level == 1}


# ------------------------------------------------------- aggregation rules

def test_state_row_reaggregates_through_the_engine(sess, frame_meta):
    from app import compute
    df, meta = frame_meta
    results = compute.book_results(sess)
    multi = [st for st in meta.states if meta.cnt[st] > 1]
    assert multi, "sample book should have multi-combo states"
    for st in multi[:3]:
        num = den = 0.0
        for lob, scn, key, _bu, row, _res in meta.view[st]:
            res = results[lob][key]
            ep = float(row.get("ep") or 0.0)
            num += res.cy_lr_p * ep
            den += ep
        got = df.loc[df["state"] == st, "planlr"].iloc[0]
        assert got == pytest.approx(num / den, abs=1e-12), st


def test_single_combo_row_is_the_visible_chain(sess):
    """The ss_mathchk twin: on a single-combo row the displayed factors
    multiply EXACTLY to the displayed plan LR — both years."""
    from app import summary
    df, meta = summary.state_frame(sess, lines=["Property"], bus=["ABD"])
    body = df[df["state"] != "TOTAL"]
    assert (pd.Series([meta.cnt[s] for s in body["state"]]) == 1).all()
    for _i, r in body.iterrows():
        chain = r["lrcur"] * r["arate"] * r["amod"] * r["aother"]
        assert r["planlr"] == pytest.approx(chain, rel=1e-12)
        chain1 = (r["lrcur"] * (1.0 + r["trend"]) * r["arate1"]
                  * r["amod1"] * r["aother"])
        assert r["planlr1"] == pytest.approx(chain1, rel=1e-12)
        assert math.isnan(r["mix"])                 # nothing to disclose


def test_mix_is_the_exact_residual(frame_meta):
    df, meta = frame_meta
    for _i, r in df[df["state"] != "TOTAL"].iterrows():
        if meta.cnt[r["state"]] <= 1 or math.isnan(r["planlr"]):
            continue
        chain = r["lrcur"] * r["arate"] * r["amod"] * r["aother"]
        assert r["mix"] == pytest.approx((r["planlr"] - chain) * 100.0,
                                         abs=1e-9)
    total = df.iloc[-1]
    assert total["state"] == "TOTAL"
    chain = total["lrcur"] * total["arate"] * total["amod"] * total["aother"]
    assert total["mix"] == pytest.approx(
        (total["planlr"] - chain) * 100.0, abs=1e-9)   # never the chain


def test_chronology_matches_the_rank_logic(sess):
    """D95: the LAST four filings, newest last, stable on same-day ties;
    planned pct = filed × achievement; token P/T + * when not considered.
    Multi-combo rows show — in every slot."""
    import copy

    from app import summary
    s2 = _stub([copy.deepcopy(sess.page.book["Property"])])
    scn = s2.page.book["Property"]
    key = scn.combo_keys()[0]
    bu, st = key.split("|", 1)
    scn.rate_rows = [r for r in scn.rate_rows
                     if not (r["bu"] == bu and r["state"] == st)] + [
        dict(bu=bu, state=st, eff=dt.date(2025, 7, 1), filed=0.04,
             status="taken", considered="Y", achievement=None, comment=""),
        dict(bu=bu, state=st, eff=dt.date(2026, 1, 1), filed=0.02,
             status="taken", considered="Y", achievement=None, comment=""),
        dict(bu=bu, state=st, eff=dt.date(2026, 7, 1), filed=0.06,
             status="taken", considered="", achievement=None, comment=""),
        dict(bu=bu, state=st, eff=dt.date(2027, 4, 1), filed=0.05,
             status="planned", considered="", achievement=0.8, comment="A"),
        dict(bu=bu, state=st, eff=dt.date(2027, 4, 1), filed=0.01,
             status="planned", considered="", achievement=None, comment="B"),
        dict(bu=bu, state=st, eff=dt.date(2027, 9, 1), filed=0.03,
             status="planned", considered="", achievement=0.5, comment=""),
    ]
    df, meta = summary.state_frame(s2, bus=[bu])
    row = df[df["state"] == st].iloc[0]
    # six filings -> slots hold the LAST four: 7/2026, 4/2027(A), 4/2027(B),
    # 9/2027 — the same-day pair keeps its log order (stable sort)
    assert row["chg1_date"] == "2026-07-01" and row["chg1_tok"] == "T*"
    assert row["chg2_date"] == "2027-04-01" and row["chg2_pct"] == "+4.0%"
    assert row["chg3_date"] == "2027-04-01" and row["chg3_pct"] == "+1.0%"
    assert row["chg4_date"] == "2027-09-01" and row["chg4_pct"] == "+1.5%"
    assert row["chg4_tok"] == "P*"
    assert row["ntaken"] == 3 and row["nplanned"] == 3
    assert (st, 2) in meta.slot_map                # the router's map exists

    # a multi-combo state shows — in every slot, "" only on the TOTAL row
    df_all, meta_all = summary.state_frame(sess)
    multi = next(s for s in meta_all.states if meta_all.cnt[s] > 1)
    r2 = df_all[df_all["state"] == multi].iloc[0]
    assert all(r2[f"chg{j}_{p}"] == "—"
               for j in (1, 2, 3, 4) for p in ("date", "pct", "tok"))
    assert all(df_all.iloc[-1][f"chg{j}_date"] == "" for j in (1, 2, 3, 4))


def test_target_fix_divergence_is_deliberate(sess):
    """Excel's w_target dilutes a state where some combos carry no target
    (0×EP in the numerator, full EP in the denominator — documented at
    sheets_calc.py:135-138). The app weights over CARRIERS; pin both
    numbers so the divergence reads as a decision."""
    import copy

    from app import summary
    s2 = _stub([copy.deepcopy(sess.page.book["Property"])])
    scn = s2.page.book["Property"]
    df0, meta0 = summary.state_frame(s2)
    multi = next(s for s in meta0.states if meta0.cnt[s] > 1)
    combos = meta0.view[multi]
    victim = combos[0][4]
    assert victim.get("target") is not None
    victim["target"] = None
    df, meta = summary.state_frame(s2)
    carriers = [(float(r["target"]), float(r["ep"]))
                for _l, _s, _k, _b, r, _res in meta.view[multi]
                if r.get("target") is not None]
    expect = sum(v * e for v, e in carriers) / sum(e for _v, e in carriers)
    got = df.loc[df["state"] == multi, "target"].iloc[0]
    assert got == pytest.approx(expect, abs=1e-12)
    diluted = (sum(v * e for v, e in carriers)
               / sum(float(r["ep"]) for _l, _s, _k, _b, r, _r2 in
                     meta.view[multi]))
    assert got != pytest.approx(diluted, abs=1e-9)  # the Excel number


def test_progbasis_invariant_and_gate(sess, frame_meta):
    from src import engine

    from app import compute
    df, meta = frame_meta
    quiet = [st for st in meta.states if meta.netcnt[st] == 0]
    for st in quiet[:2]:
        r = df[df["state"] == st].iloc[0]
        assert math.isnan(r["netsel"]) and math.isnan(r["progbasis"]) \
            and math.isnan(r["proggap"])
    net_states = [st for st in meta.states if meta.netcnt[st] > 0]
    assert net_states, "sample book should carry net combos"
    st = net_states[0]
    r = df[df["state"] == st].iloc[0]
    num = den = 0.0
    for lob, scn, key, _bu, row, res in meta.view[st]:
        ep = float(row.get("ep") or 0.0)
        if row.get("netp") is not None:
            v = engine.program_basis_plan_lr(
                scn.plan_year, compute.combo_inputs(scn, row))[0]
        else:
            v = res.cy_lr_p
        num += v * ep
        den += ep
    assert r["progbasis"] == pytest.approx(num / den, abs=1e-12)
    assert r["proggap"] == pytest.approx(
        (r["progbasis"] - r["planlr"]) * 100.0, abs=1e-9)


def test_zero_ep_state_is_dash_not_mean(sess):
    """The Book dialect, pinned as an anti-Excel divergence: a zero-EP
    state shows — (NaN), never the per-LOB plain-mean fallback."""
    import copy

    from app import summary
    s2 = _stub([copy.deepcopy(sess.page.book["Property"])])
    scn = s2.page.book["Property"]
    df0, meta0 = summary.state_frame(s2)
    st = meta0.states[0]
    for _l, _s, _k, _b, row, _r in meta0.view[st]:
        row["ep"] = None
    df, _meta = summary.state_frame(s2)
    r = df[df["state"] == st].iloc[0]
    assert math.isnan(r["planlr"]) and math.isnan(r["lrcur"])
    assert r["ep"] == 0.0


# ------------------------------------------------------------- page (venv)

@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_summary_page_builds_filters_and_toggles():
    from app import importers, summary
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page

    session = PlanSession()
    page = ss_page.build(session)
    session.replace_config(importers.from_workbook(WB_GL))
    session.replace_config(importers.from_workbook(WB_PROP))

    table = page["table"]
    meta = page["holder"]["meta"]
    assert len(table.value) == len(meta.states) + 1        # + TOTAL
    assert table.value.iloc[-1]["state"] == "TOTAL"
    # the year-bearing group captions refreshed off the empty-book build
    # (found live: they read "CY 0 plan" before the _render groups update)
    assert any(f"CY {meta.plan_year} plan" in c for c in table.groups)

    hist = summary.hidden_sets()["hist"]
    assert set(hist) <= set(table.hidden_columns)          # collapsed ship
    page["toggles"]["hist"].value = True
    assert not (set(hist) & set(table.hidden_columns))
    page["toggles"]["inputs"].value = False
    assert set(summary.hidden_sets()["inputs"]) <= set(table.hidden_columns)

    page["filters"]["line"].value = ["Property"]
    assert page["holder"]["meta"] is not meta              # re-rendered
    assert table.style is not None
    css = summary.ss_styles(table.value)
    assert (css != "").any().any()                         # styles non-empty


# ------------------------------------------------------------- W3c display

def test_display_order_is_a_banded_permutation():
    """The display layer: levers before echoes before outputs, the answer
    last (frozen right) — a PERMUTATION of ALL_KEYS, never a new list of
    columns; every display group is contiguous; the two answer columns
    are ungrouped (Tabulator ignores frozen inside a column group)."""
    from app import summary

    assert sorted(summary.DISPLAY_ORDER) == sorted(summary.ALL_KEYS)
    assert summary.DISPLAY_ORDER[-2:] == ["planlr", "planlr1"]
    groups = summary.display_groups_map(2027)
    grouped = [k for keys in groups.values() for k in keys]
    assert "planlr" not in grouped and "planlr1" not in grouped
    for keys in groups.values():                       # contiguous runs
        idx = [summary.DISPLAY_ORDER.index(k) for k in keys]
        assert idx == list(range(idx[0], idx[0] + len(idx)))
    # netsel finally sits next to netsel1
    d = summary.DISPLAY_ORDER
    assert d.index("netsel1") == d.index("netsel") + 1
    assert any("CY 2027 plan" in c for c in groups)    # page test contract


@pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")
def test_display_layer_reorders_and_freezes_the_answer():
    from app import importers, summary
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page

    session = PlanSession()
    page = ss_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))

    table = page["table"]
    assert list(table.value.columns) == summary.DISPLAY_ORDER
    assert table.frozen_columns == {"state": "left", "planlr": "right",
                                    "planlr1": "right"}
    # any dict value other than left/right silently VANISHES the column
    assert set(table.frozen_columns.values()) <= {"left", "right"}
    # the Styler was computed on the SAME reordered frame
    assert list(table.style.data.columns) == summary.DISPLAY_ORDER
    # layout: banner above the table, notes folded, undo inline not sidebar
    import panel as pn
    main = page["main"]
    banner_at = next(i for i, o in enumerate(main.objects)
                     if "THE BRIDGE IN ONE LINE" in str(
                         getattr(o, "object", "")))
    table_at = next(i for i, o in enumerate(main.objects)
                    if isinstance(o, pn.widgets.Tabulator))
    assert banner_at < table_at
    assert page["fine_print"].collapsed is True
    assert not any(isinstance(o, pn.widgets.Button)
                   for o in page["sidebar"].objects)
