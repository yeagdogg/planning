"""Guide — how the workbench, the workbooks and the engine are one thing
(P6). Static prose, but every column list renders FROM the schemas, so
the documentation cannot drift from the code that enforces it.
"""
from __future__ import annotations

import panel as pn

from app.grids import LR_SCHEMA, ML_SCHEMA, RL_SCHEMA
from app.masters import MASTERS


def _cols(schema) -> str:
    return " · ".join(f"`{title}`" for _k, title, _kind in schema)


def _body() -> str:
    masters_fmt = "\n".join(
        f"- **{label}**: {_cols(schema)}"
        for label, (schema, _attr) in MASTERS.items())
    return f"""
## What this is

One calculation, three surfaces. The Python engine (`src/engine.py`) is
the oracle; the Excel workbooks are generated from it and verified
against it cell by cell; this app calls the SAME engine through the SAME
input mapping the release harness checks. Nothing here is a second
implementation — when the app, a workbook and the engine disagree,
something is broken, and the tests treat that as the definition of
broken.

## The hub flow

1. **Start** — Book page. Either **Load every line** (pull the current
   fleet workbooks in) or expand **Master paste** and paste the three
   master tables. Masters span every line with a leading **Line** column
   (a BU can write several lines, so the line is always explicit).
2. **Work** — the Book shows every combo, EP-weighted the way the Book
   workbook aggregates. Click a row to open its Combo deep-dive
   (bridge, waterfalls, earning parallelogram, monthly race, delivery
   band). Edit on the Inputs page — grids or paste blocks — and every
   surface moves through one debounced signal, with the Checks rules
   running live in the sidebar.
3. **Remember** — Scenarios page. Save the whole book to one yaml in
   `scenarios/`; each save appends a changelog entry diffed against the
   file it replaces ("Property · ABD|AZ: Net rate P 2.0% → 2.5%").
   Commit the folder to git for version control. Load restores a saved
   book (two-step confirm); Diff compares without touching anything.
4. **Publish** — generate real fleet workbooks from the current book:
   per-line files tagged `_APP` beside the fleet, or the whole fleet +
   summary Book into `output/app/` (recalculated in Excel end to end).
   The verified fleet in `output/` is never overwritten.

## Paste formats

Per-line blocks (Inputs page) use the workbook's own column order, so
the same rectangle copies both ways. Header rows are detected by
CONTENT and skipped, so copy them or don't.

- **tbl_LR**: {_cols(LR_SCHEMA)}
- **Rate Log**: {_cols(RL_SCHEMA)}
- **Mod Log**: {_cols(ML_SCHEMA)}

Masters (Book page) are the same blocks with a leading `Line` column:

{masters_fmt}

A master REPLACES that table for every line present in the paste; lines
absent from the paste keep theirs. What masters deliberately do not
carry: per-line toggles (term, net-trend default, seasonality on/off,
mod master) and the Seasonality table — those live on the Inputs page
and ride in scenario files. `python -m app.make_masters` writes
paste-ready masters from the current fleet into `output/masters/`.

## Numbers policy

Percent columns hold FRACTIONS (0.103 = 10.3%), exactly what the
workbook's cells hold — pasting `10.3%` or `0.103` both land as 0.103.
Every book-level mean is adjusted-plan-EP weighted over what's in view;
weighted factors do **not** multiply to the weighted plan LR — the gap
is mix, disclosed rather than normalised away. A month that earns
nothing is a gap, never a zero. Rows that fail to parse error loudly,
naming the row and column; nothing is ever silently dropped except
blank spare rows.

## Where things live

- `scenarios/*.yaml` — saved books, each carrying its own changelog.
- `output/*_APP.xlsx` — per-line exports (gitignored).
- `output/app/` — the app-generated fleet + summary Book (gitignored).
- `output/masters/` — generated master-paste blocks (gitignored).
- `app/serve.bat <port>` — serve; the venv in `app/.venv` is the app's
  own and the Excel pipeline never uses it.
"""


def build(session):
    toc = pn.pane.Markdown(
        "**Guide**\n\n1. What this is\n2. The hub flow\n3. Paste formats\n"
        "4. Numbers policy\n5. Where things live",
        sizing_mode="stretch_width")
    return {"main": pn.Column(pn.pane.Markdown(_body(),
                                               sizing_mode="stretch_width"),
                              sizing_mode="stretch_both"),
            "sidebar": toc, "bag": None}
