"""State Summary — the flagship exhibit's numbers (W2b).

Mirrored from the workbook's OWN declaration: ``SS_COLS``
(src/sheets_main.py, D90) is imported and everything here — column order,
headers, group captions, collapse levels — derives from it, so the app
cannot drift from Excel without a test failing. The D90 lesson, quoted in
the source: parallel position-keyed structures once shipped "a
correct-looking exhibit with one column of the wrong thing in it."

Aggregation follows the BOOK dialect (the app session spans lines): every
metric is adjusted-plan-EP weighted over the combos in view, NaN
(rendered "—") when nothing carries weight — deliberately NOT the
per-LOB exhibit's plain-mean fallback. Deliberate divergences from the
Excel twins, each disclosed on the page and pinned by a test:

- zero-EP states show "—", never a plain mean (Book dialect);
- ``target`` (and the other optional inputs) weight over the combos that
  CARRY a value — Excel's ``w_target`` dilutes them with the full EP of
  blank rows, a flaw its own source documents (sheets_calc.py:135-138);
- chronology slots distinguish "—" (filters don't resolve to one combo)
  from "" (that combo has no such filing);
- APP_EXT columns (net sel +1, mod step, achievement, problems) exist
  only here, grouped under an "App extensions" caption.

Frames are pure pandas over engine results — both interpreters run them.
"""
from __future__ import annotations

import math
from dataclasses import dataclass, field

import pandas as pd

from src.sheets_main import (SS_COLS, SS_G_BRIDGE, SS_G_HIST, SS_G_NET,
                             SS_G_P1, SS_LV_HIST, SS_LV_INPUT)

from app import compute
from app.glue.theme import STEEL, STEEL_LIGHT, WARN_AMBER

SS_KEYS = [c.key for c in SS_COLS]

# app-extension columns — NOT in the Excel exhibit, marked as such
APP_EXT = [("netsel1", "Net rate sel +1"),
           ("modstep_date", "Mod step date"),
           ("modstep_pct", "Mod step %"),
           ("ach_next", "Ach % next planned"),
           ("problems", "⚠")]
EXT_KEYS = [k for k, _t in APP_EXT]
ALL_KEYS = SS_KEYS + EXT_KEYS

_SLOT_KEYS = [c.key for c in SS_COLS if c.key.startswith("chg")]

# key -> TAB_SPECS kind for the NUMERIC columns (floats + NumberFormatter);
# chronology/date/token columns are pre-rendered strings and stay unlisted
SS_KINDS = {
    "ep": "dollar", "ntaken": "count", "nplanned": "count",
    "mind": "mod", "m0": "mod", "m1": "mod",
    "crl": "idx", "ecy": "idx", "mbar": "mod",
    "lrcur": "pct", "arate": "idx", "amod": "idx", "aother": "idx",
    "planlr": "pct", "target": "pct", "mix": "pts",
    "trend": "pct_signed", "arate1": "idx", "amod1": "idx", "planlr1": "pct",
    "netsel": "pct_signed", "progbasis": "pct", "proggap": "pts",
    "netsel1": "pct_signed", "problems": "count",
}


def live_headers(plan_year: int) -> dict:
    """App twins of the SsCol.live year-bearing header formulas (D44)."""
    p = plan_year
    return {"ecy": f"{p} earned rate level", "mbar": f"{p} earned mod",
            "planlr": f"=  CY {p} plan LR", "trend": f"Net trend ({p + 1})",
            "arate1": f"{p + 1} rate earn-in", "amod1": f"{p + 1} mod drift",
            "planlr1": f"CY {p + 1} plan LR"}


def _caption(group: str, plan_year: int) -> str:
    p = plan_year
    if group == SS_G_BRIDGE:
        return f"CY {p} plan — the bridge, left to right"
    if group == SS_G_P1:
        return f"CY {p + 1} indicative"
    if group == SS_G_HIST:
        # the per-LOB caption says "single-BU view"; the app is the Book's
        # one-dimension-deeper twin
        return ("Rate change history — chronological (single-combo view). "
                "T = taken, P = planned, * = not counted in the indication")
    return group                                   # Volume / mods / levels / net


APP_EXT_CAPTION = "App extensions — not in the Excel exhibit"


def ss_groups_map(plan_year: int) -> dict:
    """SsCol.group runs -> the Panel Tabulator ``groups`` dict."""
    groups: dict = {}
    for c in SS_COLS:
        if not c.group:
            continue
        groups.setdefault(_caption(c.group, plan_year), []).append(c.key)
    groups[APP_EXT_CAPTION] = list(EXT_KEYS)
    return groups


def hidden_sets() -> dict:
    """The D101 outline, from SsCol.level: which keys each collapse toggle
    governs (chronology inner, bridge-inputs outer)."""
    return {"hist": [c.key for c in SS_COLS if c.level >= SS_LV_HIST],
            "inputs": [c.key for c in SS_COLS if c.level == SS_LV_INPUT]}


# ------------------------------------------------------------------ helpers

def _num(v):
    return float(v) if isinstance(v, (int, float)) else None


def _wtd(pairs) -> float:
    """Σ(v·ep)/Σep over the pairs whose value is not None (the carriers
    rule — a combo without the input contributes neither numerator nor
    weight; this is the deliberate target-dilution fix). NaN when nothing
    carries weight."""
    num = den = 0.0
    for v, ep in pairs:
        if v is None:
            continue
        den += ep
        num += v * ep
    return (num / den) if den > 0 else float("nan")


def _fmt_slot_pct(v: float) -> str:
    return f"{v:+.1%}"


@dataclass
class SsMeta:
    """Everything the page and the W2c edit router need beyond the frame."""
    plan_year: int
    states: list = field(default_factory=list)       # frame order, sans TOTAL
    cnt: dict = field(default_factory=dict)          # state -> combos in view
    netcnt: dict = field(default_factory=dict)
    problems: dict = field(default_factory=dict)
    view: dict = field(default_factory=dict)         # state -> [(lob, scn, key, bu, row, res)]
    single: dict = field(default_factory=dict)       # state -> (lob, scn, key) when cnt==1
    slot_map: dict = field(default_factory=dict)     # (state, j) -> (lob, rate_row dict)
    modstep_map: dict = field(default_factory=dict)  # state -> (lob, mod_row dict)
    achnext_map: dict = field(default_factory=dict)  # state -> (lob, rate_row dict)


def _combo_log_rows(scn, bu: str, state: str) -> list:
    return [r for r in scn.rate_rows
            if r.get("bu") == bu and r.get("state") == state
            and r.get("eff") is not None]


def _slot_cells(scn, bu: str, st: str, meta: SsMeta) -> dict:
    """The D95 chronology: the LAST four filings, newest last. Returns the
    12 pre-rendered slot strings and registers slot_map entries."""
    rows = sorted(_combo_log_rows(scn, bu, st),
                  key=lambda r: r["eff"])            # stable — ties keep order
    cells: dict = {}
    last4 = rows[-4:]
    for j in range(1, 5):
        if j <= len(last4):
            r = last4[j - 1]
            planned = (r.get("status") == "planned")
            ach = _num(r.get("achievement"))
            eff_pct = (_num(r.get("filed")) or 0.0) * \
                ((ach if ach is not None else 1.0) if planned else 1.0)
            tok = ("P" if planned else "T") + \
                ("" if r.get("considered") == "Y" else "*")
            cells[f"chg{j}_date"] = str(r["eff"])
            cells[f"chg{j}_pct"] = _fmt_slot_pct(eff_pct)
            cells[f"chg{j}_tok"] = tok
            meta.slot_map[(st, j)] = r
        else:
            cells[f"chg{j}_date"] = ""
            cells[f"chg{j}_pct"] = ""
            cells[f"chg{j}_tok"] = ""
    return cells


def state_frame(session, lines=None, bus=None) -> tuple:
    """One row per state + a TOTAL row, honoring Line + BU filters (never a
    state filter on itself — the workbook's own doctrine). Returns
    (DataFrame, SsMeta); the frame keeps a clean RangeIndex (styles and
    edit-row mapping depend on it)."""
    book = session.page.book
    sel_lines = [ln for ln in book if not lines or ln in lines]
    results = compute.book_results(session)
    prog = compute.program_results(session)
    plan_year = next((s.plan_year for s in book.values()), 0)
    meta = SsMeta(plan_year=plan_year)

    for lob in sel_lines:
        scn = book[lob]
        rmap = results.get(lob, {})
        for key in scn.combo_keys():
            bu, st = key.split("|", 1)
            if bus and bu not in bus:
                continue
            meta.view.setdefault(st, []).append(
                (lob, scn, key, bu, scn.row(key), rmap.get(key)))

    def build_row(st: str, combos: list, is_total: bool) -> dict:
        ok = [c for c in combos
              if c[5] is not None and not compute.is_error(c[5])]
        errs = len(combos) - len(ok)
        p = plan_year

        def w(field_fn):
            return _wtd([(field_fn(lob, scn, key, row, res),
                          _num(row.get("ep")) or 0.0)
                         for lob, scn, key, _bu, row, res in ok])

        rec: dict = {k: float("nan") for k in ALL_KEYS}
        rec["state"] = st
        rec["ep"] = sum((_num(c[4].get("ep")) or 0.0) for c in combos)
        rec["ntaken"] = sum(
            1 for lob, scn, key, bu2, _row, _res in combos
            for r in _combo_log_rows(scn, bu2, st)
            if r.get("status") == "taken")
        rec["nplanned"] = sum(
            1 for lob, scn, key, bu2, _row, _res in combos
            for r in _combo_log_rows(scn, bu2, st)
            if r.get("status") == "planned")

        rec["mind"] = w(lambda l, s, k, row, r: _num(row.get("m_ind")))
        rec["m0"] = w(lambda l, s, k, row, r: _num(row.get("m0")))
        rec["m1"] = w(lambda l, s, k, row, r: _num(row.get("m1")))
        rec["crl"] = w(lambda l, s, k, row, r: r.crl_ind)
        rec["ecy"] = w(lambda l, s, k, row, r: r.e_cy[p])
        rec["mbar"] = w(lambda l, s, k, row, r: r.m_earned[p])
        rec["lrcur"] = w(lambda l, s, k, row, r: r.lr_current)
        rec["arate"] = w(lambda l, s, k, row, r: r.a_rate_p)
        rec["amod"] = w(lambda l, s, k, row, r: r.a_mod_p)
        rec["aother"] = w(lambda l, s, k, row, r:
                          _num(row.get("a_other")) if row.get("a_other")
                          is not None else 1.0)
        rec["target"] = w(lambda l, s, k, row, r: _num(row.get("target")))
        trend_of = (lambda l, s, k, row, r:
                    _num(row.get("trend")) if row.get("trend") is not None
                    else float(s.trend_default or 0.0))
        rec["trend"] = w(trend_of)
        rec["arate1"] = w(lambda l, s, k, row, r: r.a_rate_p1)
        rec["amod1"] = w(lambda l, s, k, row, r: r.a_mod_p1)

        chain = rec["lrcur"] * rec["arate"] * rec["amod"] * rec["aother"]
        chain1 = (rec["lrcur"] * (1.0 + rec["trend"]) * rec["arate1"]
                  * rec["amod1"] * rec["aother"])
        wlr = w(lambda l, s, k, row, r: r.cy_lr_p)
        wlr1 = w(lambda l, s, k, row, r: r.cy_lr_p1)
        single = (not is_total) and len(combos) == 1 and len(ok) == 1
        # single-combo rows show the visible chain (== the engine value by
        # identity — tie-tested); mixed rows show the exact weighted value
        # and DISCLOSE the residual (factor averages do not compound)
        rec["planlr"] = wlr
        rec["planlr1"] = wlr1
        # mix reports CY P only (Excel parity: one Mix column); chain1 is
        # still computed so the P+1 tie test can assert the same identity
        del chain1
        if single or math.isnan(wlr) or math.isnan(chain):
            rec["mix"] = float("nan")
        else:
            rec["mix"] = (wlr - chain) * 100.0

        net = [(lob, scn, key, bu2, row, res)
               for lob, scn, key, bu2, row, res in combos
               if row.get("netp") is not None]
        nc = len(net)
        if nc:
            rec["netsel"] = (sum(_num(c[4]["netp"]) for c in net) / nc)
            n1 = [c for c in net if c[4].get("netp1") is not None]
            rec["netsel1"] = (sum(_num(c[4]["netp1"]) for c in n1) / len(n1)
                              if n1 else float("nan"))
            pb = _wtd([
                ((prog.get(lob, {}).get(key)[0]
                  if (row.get("netp") is not None
                      and prog.get(lob, {}).get(key) is not None)
                  else (res.cy_lr_p if res is not None
                        and not compute.is_error(res) else None)),
                 _num(row.get("ep")) or 0.0)
                for lob, scn, key, _bu, row, res in combos])
            rec["progbasis"] = pb
            if not (math.isnan(pb) or math.isnan(rec["planlr"])):
                rec["proggap"] = (pb - rec["planlr"]) * 100.0

        # chronology + app extensions: single-combo rows only
        if single:
            lob, scn, key, bu2, row, res = combos[0]
            meta.single[st] = (lob, scn, key)
            rec.update(_slot_cells(scn, bu2, st, meta))
            mod_planned = [r for r in scn.mod_rows
                           if r.get("bu") == bu2 and r.get("state") == st
                           and r.get("eff") is not None
                           and r.get("status") == "planned"]
            if mod_planned:
                r0 = sorted(mod_planned, key=lambda r: r["eff"])[0]
                meta.modstep_map[st] = (lob, r0)
                rec["modstep_date"] = str(r0["eff"])
                rec["modstep_pct"] = _fmt_slot_pct(_num(r0.get("chg")) or 0.0)
            else:
                rec["modstep_date"] = rec["modstep_pct"] = ""
            nxt = [r for r in _combo_log_rows(scn, bu2, st)
                   if r.get("status") == "planned"
                   and r["eff"].year == p]
            if nxt:
                r0 = sorted(nxt, key=lambda r: r["eff"])[0]
                meta.achnext_map[st] = (lob, r0)
                ach = _num(r0.get("achievement"))
                rec["ach_next"] = f"{(ach if ach is not None else 1.0):.0%}"
            else:
                rec["ach_next"] = ""
        else:
            for k in _SLOT_KEYS + ["modstep_date", "modstep_pct", "ach_next"]:
                rec[k] = "" if is_total else "—"

        rec["problems"] = errs
        if not is_total:
            meta.cnt[st] = len(combos)
            meta.netcnt[st] = nc
            meta.problems[st] = errs
        return rec

    recs = []
    for st in sorted(meta.view):
        recs.append(build_row(st, meta.view[st], is_total=False))
    meta.states = sorted(meta.view)
    all_combos = [c for st in meta.states for c in meta.view[st]]
    if all_combos:
        recs.append(build_row("TOTAL", all_combos, is_total=True))
        # TOTAL never uses the chain and never resolves single (by
        # construction: is_total short-circuits the single branch)
    df = pd.DataFrame(recs, columns=["state"] + [k for k in ALL_KEYS
                                                 if k != "state"])
    return df.reset_index(drop=True), meta


# ------------------------------------------------------------------ styles

def _lerp(a: str, b: str, t: float) -> str:
    va = [int(a[i:i + 2], 16) for i in (1, 3, 5)]
    vb = [int(b[i:i + 2], 16) for i in (1, 3, 5)]
    return "#" + "".join(f"{round(x + (y - x) * t):02x}"
                         for x, y in zip(va, vb))


def three_color_scale(s: pd.Series, lo="#D6E8D5", mid="#FFF2CC",
                      hi="#F4B8B8") -> list:
    """Excel's 3-color rule: min -> median -> max, NaN unshaded."""
    vals = s.astype(float)
    dom = vals.dropna()
    out = [""] * len(s)
    if len(dom) < 2 or dom.min() == dom.max():
        return out
    v_min, v_med, v_max = dom.min(), dom.median(), dom.max()
    for i, v in enumerate(vals):
        if pd.isna(v):
            continue
        if v <= v_med:
            t = 0.0 if v_med == v_min else (v - v_min) / (v_med - v_min)
            out[i] = f"background-color: {_lerp(lo, mid, t)}"
        else:
            t = 0.0 if v_max == v_med else (v - v_med) / (v_max - v_med)
            out[i] = f"background-color: {_lerp(mid, hi, t)}"
    return out


def ep_bar(s: pd.Series) -> list:
    """The EP data bar: a STEEL left-fill proportional to the max."""
    vals = s.astype(float)
    top = vals.max()
    out = [""] * len(s)
    if pd.isna(top) or top <= 0:
        return out
    for i, v in enumerate(vals):
        if pd.isna(v) or v <= 0:
            continue
        pct = max(2.0, 100.0 * v / top)
        out[i] = (f"background: linear-gradient(90deg, {STEEL}55 {pct:.0f}%, "
                  f"transparent {pct:.0f}%)")
    return out


def ss_styles(df: pd.DataFrame, levers: tuple = ()) -> pd.DataFrame:
    """The exhibit's whole visual layer as one css frame (D104 mirrored):
    banding, the two plan-LR color scales, the EP bar, amber-italic planned
    tokens, the STEEL_LIGHT TOTAL band — plus a blue tint on ``levers``
    columns (the workbook's blue-means-you-type color code, W2c). TOTAL is
    excluded from every scale domain and never tinted."""
    css = pd.DataFrame("", index=df.index, columns=df.columns)
    has_total = len(df) > 0 and df["state"].iloc[-1] == "TOTAL"
    body = df.iloc[:-1] if has_total else df
    css.loc[body.index[body.index % 2 == 1], :] = "background-color: #F5F7FA"
    for c in levers:
        if c in css.columns:
            css.loc[body.index, c] = "background-color: #EAF2FC"
    for col_ in ("planlr", "planlr1"):
        scale = three_color_scale(body[col_])
        for i, s in zip(body.index, scale):
            if s:
                css.loc[i, col_] = s
    for i, s in zip(body.index, ep_bar(body["ep"])):
        if s:
            css.loc[i, "ep"] = s
    for c in [k for k in df.columns if k.endswith("_tok")]:
        mask = df[c].astype(str).str.startswith("P")
        css.loc[mask, c] = (f"color: {WARN_AMBER}; font-style: italic; "
                            f"font-weight: 600")
    if has_total:
        css.iloc[-1, :] = (f"background-color: {STEEL_LIGHT}; "
                           f"font-weight: 600")
    return css
