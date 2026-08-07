"""Scenario -> engine results, cached on the change bus.

The app has exactly ONE calculation path: the same
src/build_workbook.py::sample_to_combo mapping the release harness verifies
against the workbook, feeding the same src/engine.py the workbook's oracle
ties bind to. There is no second implementation to drift — the app's tests
tie the MAPPING (toggles, carried inputs), not the math.

Caching: one full-book recompute is ~63 x run_bridge, tens of milliseconds —
cheap enough to redo on every real input change. The cache key is
(data_version, bus.rev): any scenario replacement or grid write-back
invalidates wholesale, which is simple and impossible to under-invalidate.
"""
from __future__ import annotations

from dataclasses import replace
from typing import Optional

from src import engine
from src.build_workbook import sample_to_combo

from . import importers

_CACHE_KEY = "engine_results"


def combo_inputs(scn, lr_row: dict) -> engine.ComboInputs:
    """One row's engine inputs, honoring the scenario's global toggles."""
    cfg = importers.app_config()
    lob = cfg.lob(scn.lob)
    if lob.term_months != scn.term_months:
        lob = replace(lob, term_months=scn.term_months)
    ci = sample_to_combo(
        cfg, lob, lr_row, scn.rate_rows, scn.mod_rows,
        seasonality_on=scn.season_on, season_rows=scn.season_rows,
        trend_default=scn.trend_default,
    )
    if not scn.mod_master:
        # nr_ModAdjMaster OFF pins every combo's A_mod regardless of row
        # tokens — same semantics as the workbook's single master test
        ci = replace(ci, mod_adjustment_enabled=False)
    return ci


def results(session) -> dict:
    """{combo key: EngineResult | ('error', msg)} for the current scenario,
    recomputed only when (data_version, bus.rev) moved. Bumps
    ctx.results_rev exactly when a recompute actually happened."""
    scn = session.page.config
    if scn is None:
        return {}
    key = (session.page.data_version, session.bus.rev)
    cached = session.page.caches.get(_CACHE_KEY)
    if cached is not None and cached[0] == key:
        return cached[1]
    out: dict = {}
    for r in scn.lr_rows:
        ck = f"{r.get('bu')}|{r.get('state')}"
        if not (r.get("bu") and r.get("state")):
            continue
        try:
            out[ck] = engine.run_bridge(scn.plan_year, combo_inputs(scn, r),
                                        "monthly")
        except Exception as e:                          # noqa: BLE001
            # a half-typed row must not blank the book — carry the reason
            out[ck] = ("error", str(e))
        session.page.results = out
    session.page.caches[_CACHE_KEY] = (key, out)
    session.ctx.results_rev += 1
    return out


def result_for(session, combo_key: str) -> Optional[object]:
    return results(session).get(combo_key)


def is_error(res) -> bool:
    return isinstance(res, tuple) and len(res) == 2 and res[0] == "error"
