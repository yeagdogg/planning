"""Master tables — the hub's paste-first on-ramp (P5).

John's process starts BEFORE any workbook exists: one master table each
for tbl_LR, the Rate Log and the Mod Log, spanning every line. A BU can
write several lines (confirmed), so the line is an explicit leading
column — never inferred. ``apply_master`` splits a pasted master by line
and upserts per-line Scenarios into the book:

- a line PRESENT in the paste gets that one table REPLACED;
- a line ABSENT from the paste keeps everything it has (partial
  re-pastes are a feature, and the caller's flash names what was touched);
- a line the config does not know is a loud error naming the row —
  silence here would strand premium in a typo.

Rows are stored WITHOUT the lob key, so they are byte-identical to what
the per-line paste blocks and ``from_workbook`` produce — one row
vocabulary everywhere. Seasonality stays per-line on the Inputs page
(optional, rarely pasted, and shaped differently).
"""
from __future__ import annotations

from .grids import LR_SCHEMA, ML_SCHEMA, RL_SCHEMA, _populated
from .paste import apply_block
from .state import Scenario

_LINE_COL = [("lob", "Line", "text")]

M_LR_SCHEMA = _LINE_COL + LR_SCHEMA
M_RL_SCHEMA = _LINE_COL + RL_SCHEMA
M_ML_SCHEMA = _LINE_COL + ML_SCHEMA

# table label -> (master schema, Scenario attribute it replaces)
MASTERS = {
    "tbl_LR": (M_LR_SCHEMA, "lr_rows"),
    "Rate Log": (M_RL_SCHEMA, "rate_rows"),
    "Mod Log": (M_ML_SCHEMA, "mod_rows"),
}


def split_master(table: str, text: str, known_lines: list,
                 notes: list | None = None) -> dict:
    """Pasted master -> {line name: [row dicts]} in paste order.

    Line matching is case-insensitive against the config roster; rows that
    fail the populated-row rule drop out exactly as the grids drop them
    (blank Line included — a fully blank row is spare capacity, not an
    error). A populated row with a BLANK Line carries the line above it
    down (W3h): a real Excel master names the line once per block —
    merged cells or typed once — and that paste arrives with blank Line
    on every continuation row; erroring there walled off the most common
    master shape. A blank Line with nothing above it, or an unknown
    line, still raises loudly, naming the row and the configured lines.
    """
    schema, _attr = MASTERS[table]
    base = schema[1:]
    canon = {ln.strip().lower(): ln for ln in known_lines}
    by: dict = {}
    carried = None
    for i, r in enumerate(apply_block(schema, text, notes=notes)):
        raw = (r.get("lob") or "").strip()
        body = {k: v for k, v in r.items() if k != "lob"}
        populated = _populated(base, body)
        if not populated:
            continue                           # spare row, line or not
        if not raw:
            if carried is None:
                raise ValueError(
                    f"row {i + 1}: the Line column is blank and no line "
                    f"appears above it — each block's first row must name "
                    f"its line ({', '.join(known_lines)})")
            line = carried
        else:
            line = canon.get(raw.lower())
            if line is None:
                raise ValueError(
                    f"row {i + 1}: unknown line {raw!r} — configured lines "
                    f"are {', '.join(known_lines)}")
        carried = line
        by.setdefault(line, []).append(body)
    return by


def _render_cell(kind: str, v) -> str:
    """One typed value -> the text a master block carries. Inverse of
    paste.coerce for every kind this app stores: dates go ISO, numerics go
    plain (no % signs — coerce reads bare fractions), blanks go empty."""
    if v is None or v == "":
        return ""
    if kind == "date":
        return v.isoformat() if hasattr(v, "isoformat") else str(v)
    if kind in ("pct", "num", "mod", "idx", "dollar"):
        try:
            return f"{float(v):.12g}"
        except (TypeError, ValueError):
            return str(v)
    return str(v)


def master_text(book: dict, table: str) -> str:
    """The copy-out twin of ``apply_master``: the whole book's rows for one
    table as a paste-ready master block (title header + Line column).

    The round-trip law, pinned by the tests:
    ``split_master(table, master_text(book, table))`` reproduces every
    populated line's rows exactly — so a generated master IS the book.
    """
    schema, attr = MASTERS[table]
    out = ["\t".join(title for _k, title, _kind in schema)]
    for lob, scn in book.items():
        for r in getattr(scn, attr):
            cells = [lob] + [_render_cell(kind, r.get(key))
                             for key, _t, kind in schema[1:]]
            out.append("\t".join(cells))
    return "\n".join(out)


def apply_master(session, table: str, text: str,
                 notes: list | None = None) -> tuple:
    """Apply one pasted master into ``session.page.book``.

    Returns ``(by_line, created)``: rows applied per line, and which lines
    were newly created (config defaults: the line's term, seasonality OFF,
    mod master ON). Ends with one ``activate`` swap so every surface —
    grids, bridge, book — wakes on the same signal it always uses.
    """
    from . import importers
    cfg = importers.app_config()
    known = [lob.name for lob in cfg.lobs]
    by = split_master(table, text, known, notes=notes)
    if not by:
        return {}, []
    _schema, attr = MASTERS[table]
    created = []
    for line, rows in by.items():
        scn = session.page.book.get(line)
        if scn is None:
            scn = Scenario(lob=line, plan_year=cfg.plan_year,
                           term_months=cfg.lob(line).term_months,
                           name=f"{line} — master paste",
                           source="master paste")
            session.page.book[line] = scn
            created.append(line)
        setattr(scn, attr, rows)
    active = session.page.config
    keep = (active.lob if active is not None
            and active.lob in session.page.book else next(iter(by)))
    session.activate(keep)
    return by, created
