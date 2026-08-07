"""ONE number-formatting policy for every user-facing surface — lifted from
wfmodelingworkbench (mofo/format.py), plus this domain's buckets: loss
ratios render like the workbook's FMT_PCT, mods like FMT_MOD (three
decimals), factors like FMT_IDX (four), point contributions signed to one.

John's rule, quoted in the original: "there isn't a ton of use to showing
numbers out to 30 decimal places. Maybe 4 max, but be consistent."

Three consumers, one vocabulary: f-string surfaces call fmt_*; pandas
Stylers default unspecified numeric columns to ``fmt_cell``; Tabulator
widgets use ``tab_formatters(col=KIND, ...)`` with the numeral.js twins.
"""
from __future__ import annotations

import math
from typing import Optional

import numpy as np

DASH = "—"


def _num(x) -> Optional[float]:
    if x is None:
        return None
    try:
        v = float(x)
    except (TypeError, ValueError):
        return None
    if math.isnan(v) or math.isinf(v):
        return None
    return v


def fmt_metric(x) -> str:
    """The magnitude ladder — never scientific notation."""
    v = _num(x)
    if v is None:
        return DASH
    a = abs(v)
    if a >= 1000:
        return f"{v:,.0f}"
    if a >= 10:
        return f"{v:,.2f}"
    return f"{v:.4f}"


def fmt_signed(x) -> str:
    v = _num(x)
    if v is None:
        return DASH
    a = abs(v)
    if a >= 1000:
        return f"{v:+,.0f}"
    if a >= 10:
        return f"{v:+,.2f}"
    return f"{v:+.4f}"


def fmt_pct(x, signed: bool = False, decimals: int = 1) -> str:
    """Shares / loss ratios from a FRACTION (0.634 -> '63.4%')."""
    v = _num(x)
    if v is None:
        return DASH
    spec = f"{{:{'+' if signed else ''}.{decimals}%}}"
    return spec.format(v)


def fmt_mod(x) -> str:
    """Schedule mods — the workbook's FMT_MOD (1.052)."""
    v = _num(x)
    return DASH if v is None else f"{v:.3f}"


def fmt_idx(x) -> str:
    """Rate indices / bridge factors — the workbook's FMT_IDX (0.9597)."""
    v = _num(x)
    return DASH if v is None else f"{v:.4f}"


def fmt_pts(x) -> str:
    """Point contributions — signed, one decimal (-1.2)."""
    v = _num(x)
    return DASH if v is None else f"{v:+.1f}"


def fmt_dollar(x, signed: bool = False) -> str:
    v = _num(x)
    if v is None:
        return DASH
    return f"{v:+,.0f}" if signed else f"{v:,.0f}"


def fmt_count(x) -> str:
    v = _num(x)
    if v is None:
        return DASH
    return f"{int(round(v)):,}"


def fmt_cell(v) -> str:
    """Styler default for columns without an explicit spec."""
    if v is None:
        return DASH
    if isinstance(v, (bool, str)):
        return str(v)
    if isinstance(v, (int, np.integer)):
        return f"{int(v):,}"
    if isinstance(v, (float, np.floating)):
        return fmt_metric(v)
    return str(v)


def round_floats(obj, ndigits: int = 6):
    """Display copy of a dict/JSON payload — full-precision float64 reprs
    (0.30000000000000004) never reach the user; artifacts keep exact
    values."""
    if isinstance(obj, float):
        return round(obj, ndigits)
    if isinstance(obj, dict):
        return {k: round_floats(v, ndigits) for k, v in obj.items()}
    if isinstance(obj, (list, tuple)):
        return [round_floats(v, ndigits) for v in obj]
    return obj


# ---------------------------------------------------------------- tabulator
# numeral.js twins of the buckets above (Tabulator NumberFormatter).
TAB_SPECS = {
    "metric": "0,0.0000",
    "num": "0,0.[0000]",        # up to 4 — ragged trailing zeros trimmed
    "float2": "0,0.00",
    "pct": "0.0%",              # from a FRACTION column
    "pct0": "0%",
    "pct_signed": "+0.0%",
    "mod": "0.000",
    "idx": "0.0000",
    "pts": "+0.0",
    "dollar": "0,0",
    "dollar_signed": "+0,0",
    "count": "0,0",
}


def tab_formatters(**cols) -> dict:
    """``tab_formatters(planlr="pct", ep="dollar")`` → {col:
    NumberFormatter} for pn.widgets.Tabulator(formatters=...). Values may be
    TAB_SPECS keys or raw numeral.js format strings."""
    from bokeh.models.widgets.tables import NumberFormatter

    return {col: NumberFormatter(format=TAB_SPECS.get(kind, kind))
            for col, kind in cols.items()}


def md_safe(text) -> str:
    """Data-derived text made safe for Markdown/Alert panes. Panel's
    Markdown pane renders RAW HTML, so a BU or state name like
    ``<img onerror=...>`` pasted into an input grid would execute script.
    Escaping ``&`` then ``<`` neutralizes HTML while leaving markdown
    formatting fully intact. Route any f-string that interpolates USER DATA
    into a Markdown/Alert pane through this."""
    return str(text).replace("&", "&amp;").replace("<", "&lt;")
