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


def _compute_for(scn) -> dict:
    """{combo key: EngineResult | ('error', msg)} for one scenario — the
    single loop both the active-line and the book paths run."""
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
    return out


def results(session) -> dict:
    """{combo key: EngineResult | ('error', msg)} for the ACTIVE scenario,
    recomputed only when (data_version, bus.rev) moved. Bumps
    ctx.results_rev exactly when a recompute actually happened."""
    scn = session.page.config
    if scn is None:
        return {}
    key = (session.page.data_version, session.bus.rev)
    cached = session.page.caches.get(_CACHE_KEY)
    if cached is not None and cached[0] == key:
        return cached[1]
    out = _compute_for(scn)
    session.page.results = out
    session.page.caches[_CACHE_KEY] = (key, out)
    session.ctx.results_rev += 1
    return out


def book_results(session) -> dict:
    """{lob: {combo key: result}} across every loaded line.

    Cache discipline: between data_version bumps only the ACTIVE line can
    change (the grids write the active scenario alone), so a bus bump
    invalidates just that line's entry; activate()/replace_config() bump
    data_version and revalidate everything once. The active line shares
    ``results()``'s cache, so the Combo page and the Book never compute the
    same edit twice."""
    book = session.page.book
    cache = session.page.caches.setdefault("book_results", {})
    dv, rev = session.page.data_version, session.bus.rev
    active = session.page.config
    out: dict = {}
    for lob, scn in book.items():
        is_active = scn is active
        ent = cache.get(lob)
        if ent is not None and ent[0] == dv and (not is_active
                                                 or ent[1] == rev):
            out[lob] = ent[2]
            continue
        res = results(session) if is_active else _compute_for(scn)
        cache[lob] = (dv, rev, res)
        out[lob] = res
    for lob in list(cache):                    # lines no longer loaded
        if lob not in book:
            del cache[lob]
    return out


def program_results(session) -> dict:
    """{lob: {combo key: (lr_p, lr_p1)}} on the PROGRAM basis — the logged
    rate program and projected mod path instead of any asserted net target
    (engine.program_basis_plan_lr, D65) — computed for NET combos only:
    for a non-net combo the program basis IS the headline (the engine's
    own invariant, tie-tested), so callers reuse the bridge result there.
    Cache discipline mirrors book_results: per line on (data_version,
    bus.rev); only the active line invalidates on a bus bump."""
    book = session.page.book
    cache = session.page.caches.setdefault("prog_results", {})
    dv, rev = session.page.data_version, session.bus.rev
    active = session.page.config
    out: dict = {}
    for lob, scn in book.items():
        is_active = scn is active
        ent = cache.get(lob)
        if ent is not None and ent[0] == dv and (not is_active
                                                 or ent[1] == rev):
            out[lob] = ent[2]
            continue
        res: dict = {}
        for r in scn.lr_rows:
            if r.get("netp") is None:
                continue                      # non-net: program == headline
            ck = f"{r.get('bu')}|{r.get('state')}"
            if not (r.get("bu") and r.get("state")):
                continue
            try:
                res[ck] = engine.program_basis_plan_lr(
                    scn.plan_year, combo_inputs(scn, r))
            except Exception:                           # noqa: BLE001
                # the headline path already carries this row's error
                continue
        cache[lob] = (dv, rev, res)
        out[lob] = res
    for lob in list(cache):
        if lob not in book:
            del cache[lob]
    return out


def book_frame(session):
    """One row per line × BU × state across the loaded book — the app's twin
    of the Book workbook's hidden _book sheet. Errors ride along as text so
    a half-typed row is visible in the roll-up, never silently dropped."""
    import pandas as pd

    all_res = book_results(session)
    recs = []
    for lob, scn in session.page.book.items():
        res_map = all_res.get(lob, {})
        p = scn.plan_year
        for key in scn.combo_keys():
            row = scn.row(key) or {}
            res = res_map.get(key)
            bu, state = key.split("|", 1)
            ep = row.get("ep")
            net = (row.get("netp") is not None
                   or row.get("netp1") is not None)
            rec = dict(lob=lob, bu=bu, state=state,
                       ep=float(ep) if isinstance(ep, (int, float)) else None,
                       lr_p=None, lr_p1=None, a_rate=None, a_mod=None,
                       crl=None, ecy=None, net=net, error="",
                       key=f"{lob}|{key}")
            if res is None:
                rec["error"] = "no result"
            elif is_error(res):
                rec["error"] = res[1]
            else:
                rec.update(lr_p=res.cy_lr_p, lr_p1=res.cy_lr_p1,
                           a_rate=res.a_rate_p, a_mod=res.a_mod_p,
                           crl=res.crl_ind, ecy=res.e_cy[p])
            recs.append(rec)
    cols = ["lob", "bu", "state", "ep", "lr_p", "lr_p1", "a_rate", "a_mod",
            "crl", "ecy", "net", "error", "key"]
    return pd.DataFrame(recs, columns=cols)


_WEIGHTED_COLS = ("lr_p", "lr_p1", "a_rate", "a_mod", "crl", "ecy")


def book_weighted(frame) -> dict:
    """The Book workbook's wtd() rule (sheets_book): every mean is
    adjusted-plan-EP weighted over what is IN VIEW, None when nothing
    carries weight. Error rows and EP-less rows carry no weight but are
    counted. Weighted factors deliberately do NOT compound to the weighted
    LR — the difference is mix, disclosed where both are shown (D103),
    never normalised away."""
    import pandas as pd

    ok = frame[(frame["error"] == "") & frame["lr_p"].notna()]
    # to_numeric, not fillna-on-object: the ep column is object dtype (None
    # for blank EP) and object-fillna downcasting is deprecated
    ep = pd.to_numeric(ok["ep"], errors="coerce").fillna(0.0)
    tot = float(ep.sum())
    out = dict(combos=int(len(frame)),
               errors=int((frame["error"] != "").sum()),
               net=int(frame["net"].fillna(False).astype(bool).sum()),
               lines=int(frame["lob"].nunique()), ep=tot)
    for c in _WEIGHTED_COLS:
        out[c] = (float((ok[c].astype(float) * ep).sum() / tot)
                  if tot > 0 else None)
    return out


def line_rollup(frame):
    """Per-line wtd() aggregates, in book (load) order."""
    import pandas as pd

    recs = []
    for lob, sub in frame.groupby("lob", sort=False):
        w = book_weighted(sub)
        recs.append(dict(lob=lob, combos=w["combos"], ep=w["ep"],
                         lr_p=w["lr_p"], lr_p1=w["lr_p1"],
                         a_rate=w["a_rate"], a_mod=w["a_mod"],
                         errors=w["errors"]))
    return pd.DataFrame(recs, columns=["lob", "combos", "ep", "lr_p",
                                       "lr_p1", "a_rate", "a_mod", "errors"])


def result_for(session, combo_key: str) -> Optional[object]:
    return results(session).get(combo_key)


def is_error(res) -> bool:
    return isinstance(res, tuple) and len(res) == 2 and res[0] == "error"
