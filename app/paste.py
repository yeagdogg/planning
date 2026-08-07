"""Excel-block paste: TSV text -> typed scenario rows. UI-free and tested on
both interpreters.

The column ORDER of each schema matches the workbook's paste blocks exactly
(tbl_LR A:Y per D107, Rate Log A:H, Mod Log A:G, Seasonality state+12), so
the same rectangle a user copies out of the real workbook drops straight
into the app — the D40 paste-parity property carried over.

Excel's clipboard is TSV of DISPLAYED text, so coercion accepts what Excel
shows: '10.3%' -> 0.103, '36,000' -> 36000.0, '3/31/2026' and ISO dates,
blanks -> None.
"""
from __future__ import annotations

import datetime as dt
from typing import Any, Optional

# kinds: how a column edits, formats, coerces and writes back
#   pct    fraction (edited raw, pasted with or without %)
#   num    plain float          mod/idx  float shown 3/4 dp
#   dollar thousands float      date     datetime.date
#   text   string               choice:VALUES  constrained string
_DATE_FORMATS = ("%m/%d/%Y", "%m/%d/%y", "%Y-%m-%d")


def parse_block(text: str) -> list:
    """Clipboard text -> rows of cells (None for blank). Tolerates trailing
    newline and \\r; skips fully blank lines."""
    rows = []
    for line in (text or "").replace("\r\n", "\n").replace("\r", "\n").split("\n"):
        if line.strip() == "":
            continue
        rows.append([c.strip() if c.strip() != "" else None
                     for c in line.split("\t")])
    return rows


def coerce(kind: str, raw: Any) -> Optional[Any]:
    """One displayed cell -> its typed value (None passes through)."""
    if raw is None:
        return None
    if isinstance(raw, (int, float)) and not isinstance(raw, bool):
        if kind == "date":
            return raw          # a serial has no business here; keep visible
        return float(raw)
    s = str(raw).strip()
    if s == "" or s == "—":
        return None
    if kind == "date":
        if isinstance(raw, (dt.date, dt.datetime)):
            return raw.date() if isinstance(raw, dt.datetime) else raw
        for f in _DATE_FORMATS:
            try:
                return dt.datetime.strptime(s, f).date()
            except ValueError:
                pass
        try:
            return dt.date.fromisoformat(s[:10])
        except ValueError:
            raise ValueError(f"cannot read date: {s!r}")
    if kind == "text" or kind.startswith("choice:"):
        return s
    # numeric kinds
    neg = s.startswith("(") and s.endswith(")")       # accounting negatives
    if neg:
        s = s[1:-1]
    pct = s.endswith("%")
    s = s.rstrip("%").replace(",", "").replace("$", "").strip()
    try:
        v = float(s)
    except ValueError:
        raise ValueError(f"cannot read number: {raw!r}")
    if pct:
        v /= 100.0
    if neg:
        v = -v
    return v


def apply_block(schema: list, text: str, skip_header: bool = True) -> list:
    """Pasted rectangle -> row dicts in ``schema`` order. ``schema`` is
    [(key, title, kind), ...]; extra pasted columns are ignored, short rows
    pad with None. Raises ValueError naming the cell on a bad value.

    ``skip_header``: people copy the header row along with the block more
    often than not — a first line whose cells match the column titles is
    dropped rather than reported as 'cannot read number: BU'."""
    lines = parse_block(text)
    if skip_header and lines:
        titles = {t.strip().lower() for _k, t, _kind in schema}
        first = [c.strip().lower() for c in lines[0] if isinstance(c, str)]
        hits = sum(1 for c in first if c in titles)
        if first and hits >= max(2, len(first) // 2):
            lines = lines[1:]
    out = []
    for i, cells in enumerate(lines):
        row = {}
        for j, (key, title, kind) in enumerate(schema):
            raw = cells[j] if j < len(cells) else None
            try:
                row[key] = coerce(kind, raw)
            except ValueError as e:
                raise ValueError(f"row {i + 1}, column '{title}': {e}") from e
        out.append(row)
    return out
