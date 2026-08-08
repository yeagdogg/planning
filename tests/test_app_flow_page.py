"""W3e: the Flow page — the aggregable half, pinned to the engine.

The app combiner is pinned to ``engine.combined_flow_by_month`` at 1e-15
on every slice both compute (per state, full view, and the single-combo
degenerate case); the BOOK AVG row is that same function at 1e-12; the
aggregation doctrine (delivered-is-exact, mod-leg gating, net targets
over net combos only, the never-aggregate required-pricing rule) is
encoded as tests, not prose. The cache discipline is the book_results
contract: a bus bump recomputes only the ACTIVE line.
"""
from __future__ import annotations

import sys
from pathlib import Path
from types import SimpleNamespace

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"

_needs_panel = pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")


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


def _flows(sess, lob="Property"):
    from app import compute
    scn = sess.page.book[lob]
    fr = compute.flow_results(sess)[lob]
    out = []
    for r in scn.lr_rows:
        if not (r.get("bu") and r.get("state")):
            continue
        key = f"{r['bu']}|{r['state']}"
        pf = fr[key]
        if isinstance(pf, tuple):
            continue
        out.append((r.get("ep"), pf, r))
    return scn, out


# --------------------------------------------------------------- combiner

def test_combiner_pinned_to_the_engine(sess):
    """combine_rows == engine.combined_flow_by_month at 1e-15 — per
    state, over the full view, and degenerate on one combo."""
    from src import engine

    from app import compute
    from app.views.flow import combine_rows
    scn, flows = _flows(sess)

    def engine_combined(recs):
        combos = [compute.combo_inputs(scn, r) for _ep, _pf, r in recs]
        return engine.combined_flow_by_month(scn.plan_year, combos)

    views = {"full": flows}
    by_state: dict = {}
    for ep, pf, r in flows:
        by_state.setdefault(r["state"], []).append((ep, pf, r))
    st, recs = next(iter(by_state.items()))
    views[f"state {st}"] = recs
    views["single"] = flows[:1]

    for name, recs in views.items():
        pairs = [(ep, pf) for ep, pf, _r in recs]
        want = engine_combined(recs)
        got_rows = combine_rows(pairs, "rows", slice(0, 12))
        got_prior = combine_rows(pairs, "prior_rows", slice(0, 12))
        for j in range(12):
            for k in ("rate_leg", "delivered"):
                assert got_rows[j][k] == pytest.approx(
                    want.rows[j][k], abs=1e-15), (name, j, k)
                assert got_prior[j][k] == pytest.approx(
                    want.prior_rows[j][k], abs=1e-15), (name, j, k)
            for got, eng_row in ((got_rows[j], want.rows[j]),
                                 (got_prior[j], want.prior_rows[j])):
                if eng_row["mod_leg"] is None:
                    assert got["mod_leg"] is None
                else:
                    assert got["mod_leg"] == pytest.approx(
                        eng_row["mod_leg"], abs=1e-15)


def test_combiner_needs_ep_and_extends_to_p1(sess):
    from app.views.flow import combine_rows
    _scn, flows = _flows(sess)
    pairs = [(ep, pf) for ep, pf, _r in flows]
    with pytest.raises(ValueError):
        combine_rows([(0.0, pairs[0][1])], "rows", slice(0, 12))
    ext = combine_rows(pairs, "rows", slice(0, 24))     # the P+1 half
    assert len(ext) == 24
    # degenerate single combo: the extension equals the raw rows
    one = combine_rows(pairs[:1], "rows", slice(12, 24))
    for j, row in enumerate(one):
        assert row["delivered"] == pytest.approx(
            pairs[0][1].rows[12 + j]["delivered"], abs=1e-15)


def test_avgs_tie_the_engine_and_delivered_is_exact(sess):
    """The AVG combination reproduces the engine's top-level rule — and
    under mix, weighted rate x weighted mod need NOT equal delivered
    (delivered is the exact statistic; nothing normalises the gap)."""
    from src import engine

    from app import compute
    from app.views.flow import combine_avgs
    scn, flows = _flows(sess)
    pairs = [(ep, pf) for ep, pf, _r in flows]
    combos = [compute.combo_inputs(scn, r) for _ep, _pf, r in flows]
    want = engine.combined_flow_by_month(scn.plan_year, combos)
    a = combine_avgs(pairs)
    assert a["rate"] == pytest.approx(want.avg_rate_ratio - 1, abs=1e-15)
    assert a["delivered"] == pytest.approx(want.avg_delivered_ratio - 1,
                                           abs=1e-15)
    assert a["delivered_prior"] == pytest.approx(
        want.avg_delivered_ratio_prior - 1, abs=1e-15)
    if want.avg_mod_ratio is not None:
        assert a["mod"] == pytest.approx(want.avg_mod_ratio - 1,
                                         abs=1e-15)
        assert abs((1 + a["rate"]) * (1 + a["mod"]) - 1
                   - a["delivered"]) > 0            # mix: not forced equal


# ----------------------------------------------------------------- frames

def test_grid_frame_states_plus_avg_row(sess):
    from app.views.flow import combine_rows, state_grid_frame
    _scn, flows = _flows(sess)
    view: dict = {}
    for ep, pf, r in flows:
        view.setdefault(r["state"], []).append(
            dict(ep=ep, pf=pf, netp=r.get("netp")))
    pairs = [(ep, pf) for ep, pf, _r in flows]
    df = state_grid_frame(view, pairs, "delivered", "BOOK AVG")
    assert df.iloc[-1]["state"] == "BOOK AVG"
    want = combine_rows(pairs, "rows", slice(0, 12))
    for j in range(12):
        assert df.iloc[-1][f"pp{j + 1:02d}"] == pytest.approx(
            want[j]["delivered"], abs=1e-12)
    st = df.iloc[0]["state"]
    st_pairs = [(r["ep"], r["pf"]) for r in view[st]]
    w_st = combine_rows(st_pairs, "prior_rows", slice(0, 12))
    assert df.iloc[0]["pm01"] == pytest.approx(w_st[0]["delivered"],
                                               abs=1e-12)


def test_nd_summary_targets_over_net_combos_only(sess):
    from app.views.flow import nd_summary_frame
    scn, flows = _flows(sess)
    view: dict = {}
    for ep, pf, r in flows:
        view.setdefault(r["state"], []).append(dict(
            lob=scn.lob, key=f"{r['bu']}|{r['state']}", ep=ep, pf=pf,
            netp=r.get("netp"), netp1=r.get("netp1"),
            lr_p=0.6))
    df = nd_summary_frame(view, {}, "BOOK AVG")
    nets = df[df["netcnt"] > 0]
    assert len(nets)                       # the shipped book has SBA|KY
    for _i, row in df.iterrows():
        if row["netcnt"] == 0 and row["state"] != "BOOK AVG":
            assert row["tgt"] is None      # no combo asserts -> no target
        elif row["netcnt"] > 0:
            assert row["tgt"] is not None
            assert row["set1"] in ("entered", "carried", "mixed")
            assert row["gap"] == pytest.approx(
                (row["delivered"] - row["tgt"]) * 100.0, abs=1e-12)


def test_required_pricing_gates_on_single_net_combo(sess):
    from app.views.flow import required_pricing_frame
    scn, flows = _flows(sess)
    view: dict = {}
    for ep, pf, r in flows:
        view.setdefault(r["state"], []).append(dict(
            ep=ep, pf=pf, netp=r.get("netp")))
    df = required_pricing_frame(view, scn.plan_year)
    for _i, row in df.iterrows():
        recs = view[row["state"]]
        single_net = len(recs) == 1 and recs[0]["netp"] is not None
        if single_net:
            x, pf = recs[0]["netp"], recs[0]["pf"]
            want = (1 + x) / (1 + pf.rows[0]["rate_leg"]) - 1
            assert row["pp01"] == pytest.approx(want, abs=1e-15)
        else:
            assert row["pp01"] is None     # dashed, never approximated


def test_flow_results_cache_discipline(sess):
    """The book_results contract: a bus bump recomputes only the ACTIVE
    line; a data_version bump revalidates everything."""
    from unittest.mock import patch

    from src import engine as eng_mod

    from app import compute
    compute.flow_results(sess)                          # warm
    with patch.object(eng_mod, "program_flow_by_month",
                      wraps=eng_mod.program_flow_by_month) as spy:
        compute.flow_results(sess)                      # cached: no calls
        assert spy.call_count == 0
        sess.bus.rev += 1                               # active line only
        compute.flow_results(sess)
        active = sum(1 for r in sess.page.book["Property"].lr_rows
                     if r.get("bu") and r.get("state"))
        assert spy.call_count == active
        spy.reset_mock()
        sess.page.data_version += 1                     # everything
        compute.flow_results(sess)
        total = sum(1 for scn in sess.page.book.values()
                    for r in scn.lr_rows
                    if r.get("bu") and r.get("state"))
        assert spy.call_count == total


# ------------------------------------------------------------- venv page

@_needs_panel
def test_flow_page_builds_and_renders():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import flow as flow_page

    session = PlanSession()
    page = flow_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    session.page.book["General Liability"] = importers.from_workbook(WB_GL)
    page["on_show"]()

    sdf = page["sum_tbl"].value
    assert len(sdf) > 1 and sdf.iloc[-1]["state"] == "BOOK AVG"
    gdf = page["grid_tbl"].value
    assert gdf.iloc[-1]["state"] == "BOOK AVG"
    assert page["grid_tbl"].frozen_rows == [len(gdf) - 1]
    # prior columns hidden until the toggle
    assert "pm01" in page["grid_tbl"].hidden_columns
    page["show_prior"].value = True
    assert "pm01" not in page["grid_tbl"].hidden_columns
    # leg flip re-renders
    page["leg"].value = "rate"
    assert page["grid_tbl"].value.iloc[-1]["state"] == "BOOK AVG"
    assert "plan year 2027" in str(page["main"][1].object)
    ndf = page["nd_tbl"].value
    assert (ndf["netcnt"] > 0).any()
    assert "avg YoY delivered" in page["chips"].object


@_needs_panel
def test_flow_page_click_through_needs_single_combo():
    from types import SimpleNamespace

    from app import importers
    from app.glue.session import PlanSession
    from app.pages import flow as flow_page

    session = PlanSession()
    page = flow_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    page["on_show"]()
    sdf = page["sum_tbl"].value
    multi = next(i for i in range(len(sdf) - 1)
                 if len(page["holder"]["view"][sdf.iloc[i]["state"]]) > 1)
    page["open_state"](SimpleNamespace(row=multi))
    assert "narrow the filters" in page["flash"].object

    page["filters"]["bu"].value = ["ABD"]
    page["on_show"]()
    sdf2 = page["sum_tbl"].value
    single = next(i for i in range(len(sdf2) - 1)
                  if len(page["holder"]["view"][sdf2.iloc[i]["state"]])
                  == 1)
    st = sdf2.iloc[single]["state"]
    page["open_state"](SimpleNamespace(row=single))
    assert session.ctx.focus == f"ABD|{st}"
