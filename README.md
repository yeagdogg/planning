# Calendar-Year Plan Loss Ratio Workbook

A config-driven Python generator that produces premium, fully-audited Excel workbooks
converting a projected loss ratio from a rate level indication (policy-year basis, rate fully
earned) into a **calendar-year plan loss ratio** (average earned rate level basis) — with
dashboards for rate and price flow through the plan year and the following year.

**Dimensioning:** one workbook per **line of business**; within each workbook, every input row
(projected LR, mods, rate changes) is one **business unit × state** combination. The default
config seeds 3 BUs × 21 states = 63 combos per workbook (plus spare rows), all computed
simultaneously on the Portfolio sheet. The policy term is a workbook-level parameter (Inland
Marine ships as the 6-month-term example).

**Key exhibits and options:**
- **State Summary** — the leadership exhibit: one row per state (adjusted EP, mods, chronological
  rate-change history, engine results, CY P and P+1 plan LR) with a BU filter whose "All" view
  combines business units on **adjusted-plan-EP weights** (the EP column is the weight behind
  every aggregate in the book — enter adjusted EP there).
- **Default net trend for P+1** on Control, inherited wherever a combo's trend cell is blank;
  per-state entries override.
- **Optional net rate selection** per BU × state (DECISIONS.md D39): declare the year-over-year
  combined rate + price target achieved from 1/1/P and leave the specifics for later. History
  earns as modeled; from 1/1/P each written cohort renews the target above its year-ago cohort,
  superseding planned rows and the mod projection. The bridge then shows one combined factor
  `A_net` with `A_mod = 1`. Blank = classic explicit-program behavior, bit for bit.
- **Same-row live results** (v2.1): the Rate Log and tbl_LR carry live columns showing each
  row's combo plan LR (and the delta, factors, and % of the plan year an action earns) — type
  a date or a change and the answer moves on the same line.
- **Walkthrough** — the fully worked example, live for the selected combo, from inputs through
  three specimen cohorts to the assembled plan LR, with a reconciliation strip tying it to the
  Bridge, the hidden engine, and the baked oracle.
- **Decision Board** on Portfolio (top movers, EP-weighted contribution, portfolio bridge),
  **One-Pager** (a print-ready brief per combo), and **Compare** (any two combos side by side).
- The **Solver** can follow the Control selection or solve any combo via its own override.

Method: standard on-leveling / current-rate-level methodology (Werner & Modlin, *Basic
Ratemaking*, Ch. 5), implemented as a monthly written-index × earning-matrix engine and extended
with a schedule-mod adjustment. The full methodology writeup ships inside the workbook
(`Methodology` sheet); every judgment call made during the build is logged in
[DECISIONS.md](DECISIONS.md).

## Repository layout

```
config/config.yaml         structure config: plan year, BUs, states, LOBs/terms, capacities
src/engine.py              the ORACLE: monthly engine + continuous parallelogram cross-check
src/build_workbook.py      generator CLI + config + sample data + layout + assembly
src/xlstyle.py             theme and cell-writing helpers
src/sheets_inputs.py       _lists, Inputs, Control builders
src/sheets_ratelog.py      Rate Log sheet (paste block + same-row live results)
src/sheets_engine.py       Rate Engine, Mod Engine + the shared cohort-block writer
src/sheets_calc.py         hidden _calc engine blocks (portfolio/scenario/attribution/solver)
src/sheets_main.py         Bridge, Portfolio (+ Decision Board), State Summary, Scenarios,
                           Solver, Attribution builders
src/sheets_netdelivery.py  Net Delivery tab + hidden _netcalc blocks (D57)
src/sheets_programflow.py  Program Flow tab: state x month delivered-flow grids (D59)
src/sheets_walkthrough.py  Walkthrough: the fully worked example, live for the selection
src/sheets_briefs.py       One-Pager (print-ready brief) and Compare builders
src/sheets_report.py       Flow Dashboard, _oracle, Checks, Methodology, Read Me builders
tests/test_engine.py       pytest suite: worked example, property tests, identities
tests/test_net_delivery.py net delivery closed forms vs brute force (D57/D58)
tests/test_program_flow.py program-flow legs, locked leg, ratio averages (D59)
tests/test_layout.py       Layout geometry (incl. the D56 dual-module guard)
tests/test_style.py        prose row-height calibration (nothing may clip)
tools/recalc.py            headless recalculation (Excel COM, LibreOffice fallback)
tools/verify_workbook.py   verification harness: static scans, oracle ties, toggle exercises
output/Plan_LR_Workbook_2027_<LOB>.xlsx   one generated workbook per LOB (values cached)
DECISIONS.md               every judgment call and why
```

## Setup

Python 3.11+ with `openpyxl`, `pyyaml`, `pytest`, and (Windows, for headless recalculation
through desktop Excel) `pywin32`:

```bash
pip install openpyxl pyyaml pytest pywin32
```

## Generate

```bash
python -m src.build_workbook --config config/config.yaml --out output/
```

This emits one workbook per configured LOB (`output.lobs: all`, or list specific LOBs there);
`--lob "Property"` generates a single workbook.

**Important:** `openpyxl` writes formulas without cached values. Either open each file once in
Excel (it calculates automatically), or run the recalculation tool so the shipped files carry
values:

```bash
python tools/recalc.py output/Plan_LR_Workbook_2027_Property.xlsx
```

The recalc tool uses desktop Excel through COM automation when available (the authoritative
calculation engine), falling back to LibreOffice headless (`soffice --convert-to xlsx`).
LibreOffice was not installed on the build machine; Excel 16.0 via COM was used for all
verification. Classic formula mode (Excel-2007-era functions only) is used throughout, so the
workbook also recalculates in LibreOffice and any corporate Excel build.

## Verify

```bash
python -m pytest tests -q                       # oracle suite: worked example + property tests
python tools/verify_workbook.py --lob Property  # full harness (--quick skips toggle exercises)
```

The harness proves, against a freshly recalculated file: zero formula errors in every sheet
under every exercised state; no volatile / dynamic-array / prohibited functions; no merged
cells; no external links; the §9 worked example reproduced exactly (Bridge = oracle monthly
convention to 1e-9, and 63.36% at 4 decimals on annual-term books); **every BU×state combo**
tied to a per-combo oracle run (9 metrics each); solver round-trip (+5.0% at 4/1); chart axes
intact after the Excel resave; scenario, attribution, seasonality, basis, mod-toggle,
degenerate-input, and plan-year-change exercises each tied to fresh oracle runs. At last run:
Property **190 checks / 0 failed** (full), the other five LOB files 57 (or 55 for the
6-month-term Inland Marine) / 0 failed each (phases A–C), pytest 124/124. The business
units and states in `config/config.yaml` are addressed positionally, so renaming them to
your own book keeps every seeded example working (a repeated BU or state is rejected at
load time).

## Regeneration workflow (new plan year)

1. Edit `config/config.yaml` (bump `plan_year`; adjust BUs, states, LOBs/terms, or
   `table_capacity` if needed — capacities scale from the dimensions automatically).
2. Rerun the generator and the recalc tool.
3. Paste your team's inputs into the new workbooks' tables — each dataset is one contiguous
   paste block (`tbl_LR` columns A:R, `tbl_RateLog` columns A:H; keys and engine helpers sit
   to the right and recalculate on their own), and schemas are stable across regenerations.
   Default capacities per workbook: 69 `tbl_LR` rows (63 BU×state + 6 spare), 240 rate-log
   rows, 24 seasonality rows (fixed; do not insert or delete rows).
4. Confirm each `Checks` sheet shows **ALL CHECKS PASS**.

Note: oracle-tie rows on `Checks` compare against constants baked for the *seeded sample*; they
report `N/A — INPUTS CHANGED` once you replace the sample data. Structural, identity, and
sanity checks stay live forever. Regenerating re-bakes the constants.

## First-open checklist

1. Open a generated workbook (e.g., `output/Plan_LR_Workbook_2027_Property.xlsx`) and let
   Excel calculate (automatic).
2. `Control` must show **ALL CHECKS PASS** in the Checks KPI card.
3. With sample data intact, the `Bridge` shows **63.36%** for `BU-A | AZ` in annual-term
   workbooks (the §9 worked example; each workbook's Read Me states its own expected value),
   and the `Solver` returns **+5.0%** at 4/1/2027 for its seeded target.
4. Replace the SAMPLE rows on `Inputs` with your book.

## Design notes and limitations

- All actuarial inputs live **in the workbook** (Inputs sheet); the config defines structure
  only. Blue font = inputs, black = same-sheet formulas, green = cross-sheet links, yellow
  fill = required.
- No VBA, no merged cells, no external links, no volatile functions, no `IFERROR`. Sheets are
  not protected (stated on `Read Me`); the `Checks` panel is the integrity mechanism.
- `formula_mode: modern` is reserved but intentionally not implemented (DECISIONS.md D18).
- v1 limitations (stated on `Methodology`): no new/renewal split, no exposure/audit premium
  effects, no retention response, premium-volume planning out of scope, inputs entered/pasted
  manually (Snowflake / Power Query refresh is a future hook). All four future hooks from the
  brief are documented on `Methodology`.
- BU, state, and LOB names may not contain `*`, `?`, or `|` (they feed COUNTIFS criteria and
  the concatenated `BU|State` key).
