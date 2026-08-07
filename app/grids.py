"""The input grids: schemas + DataFrame conversion + Tabulator construction.

One schema per table drives everything — grid columns, editors, client-side
formatters, paste coercion, and write-back typing — so the four consumers
cannot drift. Column ORDER mirrors the workbook's paste blocks (D107 A:Y for
tbl_LR; Rate Log A:H; Mod Log A:G; Seasonality state+12).

Grid pattern (wfmodelingworkbench's projection_editor): typed editors, the
``value`` watcher is the authoritative write-back, rendering must never
delete config — blank spare rows are rendered for growth and dropped by the
same populated-row rule the workbook and carry.py use (BU and State both
present; date too, for the logs).
"""
from __future__ import annotations

import datetime as dt

import pandas as pd

from .paste import coerce

# ---------------------------------------------------------------- schemas
# (key, title, kind) — kinds documented in app/paste.py

LR_SCHEMA = [
    ("bu", "BU", "text"), ("state", "State", "text"),
    ("ep", "Plan EP (000s)", "dollar"),
    ("lr_proj", "Projected LR", "pct"),
    ("prem_trend", "Prem trend", "pct"), ("loss_trend", "Loss trend", "pct"),
    ("expense", "Expense", "pct"), ("alae", "ALAE", "num"),
    ("ulae", "ULAE", "num"), ("combined", "Combined", "pct"),
    ("target", "Target LR", "pct"), ("cat_load", "Cat load", "pct"),
    ("large_load", "Large loss", "pct"),
    ("m_ind", "M_ind", "mod"), ("m_prior", "M_prior", "mod"),
    ("m0", "M_0", "mod"), ("m0_asof", "M_0 as-of", "date"),
    ("m_end_prior", "M_endPrior", "mod"), ("m1", "M_1", "mod"),
    ("m2", "M_2", "mod"),
    ("trend", "Net trend", "pct"), ("a_other", "A_other", "idx"),
    ("modadj", "Mod adj", "choice:,ON,OFF"),
    ("netp", "Net rate P", "pct"), ("netp1", "Net rate P+1", "pct"),
]

RL_SCHEMA = [
    ("bu", "BU", "text"), ("state", "State", "text"),
    ("eff", "Effective", "date"), ("filed", "Filed %", "pct"),
    ("status", "Status", "choice:taken,planned"),
    ("considered", "Considered", "choice:,Y"),
    ("achievement", "Achievement", "pct"), ("comment", "Comment", "text"),
]

ML_SCHEMA = [
    ("bu", "BU", "text"), ("state", "State", "text"),
    ("eff", "Effective", "date"), ("chg", "Mod change %", "pct"),
    ("status", "Status", "choice:taken,planned"),
    ("achievement", "Achievement", "pct"), ("comment", "Comment", "text"),
]

SE_SCHEMA = ([("state", "State", "text")]
             + [(f"m{i}", m, "num") for i, m in enumerate(
                 ("Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug",
                  "Sep", "Oct", "Nov", "Dec"), start=1)])

SPARE_ROWS = 10          # blank growth rows rendered under the data

_FMT_BY_KIND = {"pct": "0.0[00]%", "mod": "0.000", "idx": "0.0000",
                "num": "0.[0000]", "dollar": "0,0"}


def _populated(schema, row: dict) -> bool:
    """The workbook's own populated-row rule: BU+State (and, for the logs,
    an effective date) or the row is spare capacity."""
    keys = [k for k, _t, _k in schema]
    if "state" in keys and "bu" not in keys:          # seasonality
        return bool(row.get("state"))
    ok = bool(row.get("bu")) and bool(row.get("state"))
    if ok and "eff" in keys:
        ok = row.get("eff") is not None
    return ok


# ------------------------------------------------------------ df <-> rows

def rows_to_df(schema, rows: list, spares: int = SPARE_ROWS) -> pd.DataFrame:
    """Scenario rows -> the grid frame (object dtype throughout so None
    survives; dates as ISO strings for the date editor)."""
    cols = [k for k, _t, _k in schema]
    recs = []
    for r in rows:
        rec = {}
        for k, _t, kind in schema:
            v = r.get(k)
            if kind == "date" and isinstance(v, (dt.date, dt.datetime)):
                v = v.isoformat()[:10]
            rec[k] = v
        recs.append(rec)
    for _ in range(spares):
        recs.append({k: None for k in cols})
    df = pd.DataFrame(recs, columns=cols, dtype=object)
    return df.where(pd.notna(df), None)


def df_to_rows(schema, df: pd.DataFrame) -> list:
    """Grid frame -> typed scenario rows, spare blanks dropped. Values a
    user typed arrive as strings/numbers from Tabulator; coerce() gives the
    same typing the paste path gets, so the two entry routes cannot
    diverge. A cell that will not coerce keeps the row out with a raised
    ValueError (the page catches and toasts it)."""
    out = []
    for _, rec in df.iterrows():
        row = {}
        for k, title, kind in schema:
            v = rec.get(k)
            if v is not None and pd.isna(v):
                v = None
            row[k] = coerce(kind, v)
        if _populated(schema, row):
            out.append(row)
    return out


def rows_equal(a: list, b: list) -> bool:
    return a == b


# ----------------------------------------------------- seasonality adapters
# Scenario stores {"state", "weights"[12]}; the grid shows 12 columns.

def season_to_grid(season_rows: list) -> list:
    out = []
    for r in season_rows:
        rec = {"state": r.get("state")}
        w = list(r.get("weights") or [])
        for i in range(12):
            rec[f"m{i + 1}"] = w[i] if i < len(w) else None
        out.append(rec)
    return out


def grid_to_season(rows: list) -> list:
    out = []
    for r in rows:
        if not r.get("state"):
            continue
        out.append({"state": r["state"],
                    "weights": [float(r.get(f"m{i + 1}") or 0.0)
                                for i in range(12)]})
    return out


# ------------------------------------------------------------- tabulator

def _editors(schema) -> dict:
    ed = {}
    for k, _t, kind in schema:
        if kind == "date":
            ed[k] = {"type": "input"}          # ISO text; coerce() parses
        elif kind.startswith("choice:"):
            vals = kind.split(":", 1)[1].split(",")
            ed[k] = {"type": "list", "values": vals}
        elif kind == "text":
            ed[k] = {"type": "input"}
        else:
            ed[k] = {"type": "number", "step": 0.001}
    return ed


def _formatters(schema) -> dict:
    from bokeh.models.widgets.tables import NumberFormatter
    out = {}
    for k, _t, kind in schema:
        spec = _FMT_BY_KIND.get(kind)
        if spec:
            out[k] = NumberFormatter(format=spec)
    return out


def make_grid(schema, rows: list, *, height: int = 440):
    """An editable Tabulator over ``rows``. Caller wires the ``value``
    watcher (the authoritative write-back). Scars honored: freeze by NAME;
    no button columns; clipboard enabled as a bonus path (the paste box is
    the guaranteed one)."""
    import panel as pn

    keys = [k for k, _t, _k in schema]
    grid = pn.widgets.Tabulator(
        rows_to_df(schema, rows),
        show_index=False,
        titles={k: t for k, t, _ in schema},
        editors=_editors(schema),
        formatters=_formatters(schema),
        frozen_columns=[k for k in ("bu", "state") if k in keys],
        text_align={k: "right" for k, _t, kind in schema
                    if kind not in ("text",) and not kind.startswith("choice")},
        layout="fit_data_table",
        height=height,
        sizing_mode="stretch_width",
        # copy OUT works (Ctrl+C a selection); paste-IN deliberately does
        # not: Tabulator's client-side paste never reaches Panel's model, so
        # advertising it would eat user edits — the paste card is the path
        configuration={"clipboard": "copy"},
    )
    return grid
