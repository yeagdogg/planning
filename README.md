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
src/sheets_modlog.py       Mod Log sheet: dated schedule-mod actions (D70)
src/carry.py               Reads a built workbook's inputs back out, by defined
                           name, so a rebuild keeps your book (D82)
src/sheets_netdelivery.py  Net Delivery tab + hidden _netcalc blocks (D57/D75)
src/sheets_programflow.py  Program Flow tab: state x month delivered-flow grids (D59/D68)
src/sheets_book.py         the BOOK: combined roll-up across every LOB (D66)
src/sheets_walkthrough.py  Walkthrough: the fully worked example, live for the selection
src/sheets_briefs.py       One-Pager (print-ready brief) and Compare builders
src/sheets_report.py       Flow Dashboard, _oracle, Checks, Methodology, Read Me builders
tests/test_engine.py       pytest suite: worked example, property tests, identities
tests/test_net_delivery.py net delivery closed forms vs brute force (D57/D58)
tests/test_program_flow.py program-flow legs, locked/planned split both legs (D59/D77)
tests/test_mod_steps.py    stepped mod path, the endpoint-average identity (D70/D76)
tests/test_mod_solve.py    inverting the mod step for a target plan LR (D73)
tests/test_harvest.py      harvester refusals and provenance parsing (D66)
tests/test_layout.py       Layout geometry (incl. the D56 dual-module guard)
tests/test_style.py        prose row-height calibration (nothing may clip)
tools/recalc.py            headless recalculation (Excel COM, LibreOffice fallback)
tools/verify_workbook.py   verification harness: static scans, oracle ties, toggle exercises
tools/harvest.py           reads the six recalculated workbooks' published per-combo rows
tools/build_book.py        harvest -> the combined book workbook
tools/verify_book.py       book harness: harvest ties, aggregation ties, filter exercises
output/Plan_LR_Workbook_2027_<LOB>.xlsx   one generated workbook per LOB (values cached)
output/Plan_LR_Book_2027.xlsx             the combined book, harvested from all six
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
intact and no legend overlapping its plot after the Excel resave; scenario, attribution,
seasonality, basis, mod-toggle,
degenerate-input, stepped-mod, and plan-year-change exercises each tied to fresh oracle runs. At last run:
Property **277 checks / 0 failed** (full), the other five LOB files 94 (or 92 for the
6-month-term Inland Marine) / 0 failed each (phases A–C), the combined book **56 / 0**
(`tools/verify_book.py` — including the source-freshness phase), pytest 274/274. Each verify
run now works in its own temporary scratch directory, so LOBs can be verified in parallel.

**Recalculation cost, measured** (Excel `CalculateFullRebuild`, whole workbook, median of 3):
**0.36 s**. Worth stating because the shape of the file invites the opposite assumption — the
log-scan ranges are shared, so editing one Rate Log cell dirties nearly every one of the
~171k formulas. It rebuilds in a third of a second anyway. Optimising the formulas further
buys milliseconds; the thing actually worth waiting on is the verification harness, not the
workbook.

The business
units and states in `config/config.yaml` are addressed positionally, so renaming them to
your own book keeps every seeded example working (a repeated BU or state is rejected at
load time).

## Two years on the flow tabs

Program Flow and Net Delivery show **the year now flowing beside the plan year** — steel
band on the left (Jan P−1 … Dec P−1, measured against P−2), navy on the right (Jan P …
Dec P). Reading across the seam is the point: a filing shows up in the left band while it
is still earning through, and in the right band only for the months before its anniversary,
so the two blocks together say how much of today's momentum survives into the plan.

On Program Flow the prior-year columns are an **outline group** — the `[−]` button above
column N collapses the whole left band, summary columns included, leaving exactly the
plan-year-only tab. Net Delivery has no group (its summary keeps input columns under those
letters) and carries the prior year on the rate leg only: there is no target in P−1, so a
required pricing walk is undefined there and that block is deliberately empty.

Nothing new is modelled — the cohort blocks always spanned Jan P−2 … Dec P+1, so this is
data that was computed and never shown.

## Schedule mods you can act on

The mod path used to be a drift line between two planner-supplied anchors: the
current average written mod, and a projected mod at 12/31 of the plan year.
That is how a mod is **forecast**, not how it is **executed**. A schedule-mod
change is a dated percent that stays in force until superseded — the same
mechanic as a rate filing — so it has its own log (D70).

Drift owns history through 12/31 of the current year; the **Mod Log** owns the
plan year onward; the two meet at `M_endPrior`, so the path is continuous and
no action is counted twice. The day-blend that splits a mid-month filing is
the *same function* for both legs, deliberately: two copies would drift apart.
**Leave the sheet empty and nothing changes** — with no actions the engine is
bit-identical to the previous release.

What that unlocks:

- **Solver Mode C** answers the plan-LR question on the pricing lever: the mod
  step required by effective month, what to direct at your achievement
  assumption, the share of the year still reachable, and where the feasibility
  cliff falls. On the sample book, hitting 64.0% costs **+2.02% in March and
  +3.16% in May**, and is infeasible from October (D73).
- **Net Delivery** publishes its pricing ask twice: as a year-end mod *level*,
  and — once the log pins the path, where a level can no longer answer — as a
  dated **action** at the filing's own date, grossed up by achievement (D75).
- **Mod Engine** reconciles the shop's current process to the model. Averaging
  the two year-end mods is *exact* for one action on 1/1, a twelve-month term,
  and a flat prior year, so every departure is attributable: term, timing, and
  history, tied to the gap at 1e-9. On the sample book the current process
  reads about a **point lower** than the plan year earns — it credits more mod
  relief than arrives (D76).
- **Flow Dashboard** splits the pricing leg into locked and planned, the way
  the rate leg always was — how much of each month's delivery is already done,
  and how much still rides on actions not yet taken (D77).

## The combined book

One workbook per LOB stays the working artifact; the **book** stacks all of them into a
single filterable roll-up (line / business unit / state), with the same State Summary,
Portfolio, roll-up and Program Flow exhibits. Its per-combo figures are harvested VALUES —
the engines are frozen — while every filter and total is live, so it opens in seconds.

```bash
python tools/build_book.py
```

The harvest reads the six *recalculated* workbooks and refuses two ways: a workbook that was
never recalculated (or was built by an older generator), and one whose **own Checks panel is
failing** — a red line of business must not roll up into a green book (D83). What each
source reported when it was read is stamped in the `Source checks` column on the book's
`Control` sheet and counted in a hard Book Checks row; `--allow-failing-sources` is the
escape hatch, and it still marks the bad sources.

The weekly loop is short: **save the LOB workbook in Excel, then rerun `build_book.py`.**
Regeneration is not part of it — that is for structural changes only, and it needs
`--carry-forward` (see below). `tools/verify_book.py` compares each source's file time
against the harvest stamp and fails if one moved on, since external links are banned and
nothing inside the book can notice on its own.

## Regenerating without losing your book

**Regeneration replaces a workbook.** It builds a fresh copy from the sample seeds and
writes it over the file — it does not read, merge, or preserve anything you pasted. That
used to happen silently; now the generator refuses (D82):

```bash
python -m src.build_workbook --lob "Property" --carry-forward
```

`--carry-forward` reads `tbl_LR`, the Rate Log, the Mod Log, seasonality and the Control
scalars out of the existing workbook and rebuilds around **your** inputs, so a structural
change (new state, larger log, new generator version) costs nothing. It reads by defined
name, not cell address, so it also works on files built by an older generator, and it needs
no recalculation first — a file saved from Excel seconds ago carries forward fine.

Without it, a workbook holding real inputs is refused unless you pass `--force`. Either way
a timestamped `.bak.xlsx` is written before anything is overwritten.

### New plan year, step by step

1. Edit `config/config.yaml` (bump `plan_year`; adjust BUs, states, LOBs/terms, or
   `table_capacity` if needed — capacities scale from the dimensions automatically).
2. Rebuild, carrying last year's inputs forward: `python -m src.build_workbook
   --carry-forward` (bare, so each LOB carries from its own file), then run the recalc tool.
3. Update what actually changed for the new year — projected LRs, the rate program, mod
   actions — rather than re-keying the whole book. Each dataset is one contiguous paste
   block (`tbl_LR` columns A:S, `tbl_RateLog` A:H, `tbl_ModLog` A:G; keys and engine helpers
   sit to the right and recalculate on their own). Default capacities per workbook: 69
   `tbl_LR` rows (63 BU×state + 6 spare), 240 rate-log rows, 120 mod-log rows, 24
   seasonality rows (fixed; do not insert or delete rows).
4. **Clear the seeded SAMPLE mod actions** if you started from a fresh build rather than a
   carry-forward. They attach positionally, so on a real roster they land on real combos and
   move real plan loss ratios. A Checks advisory flags them.
5. Confirm each `Checks` sheet reports no FAILing row, then rebuild the book
   (`python tools/build_book.py`).

Note: oracle-tie rows on `Checks` compare against constants baked at build time. A seeded
build bakes them for the sample and they report `N/A — INPUTS CHANGED` once you paste over
it; a `--carry-forward` build bakes them for **your own** first combo, so they stay live.
Structural, identity, and sanity checks stay live either way.

### What the status banner says

`Checks!C3` (mirrored on `Control`) has three states (D80):

| Banner | Meaning |
|---|---|
| `ALL CHECKS PASS` | green — nothing failing, nothing advisory |
| `PASS WITH n WARNING(S)` | amber — nothing failing, but read the advisories |
| `CHECKS FAILING: n` | red — do not use the results |

A freshly generated sample workbook reads amber: it ships with one standing advisory, which
is honest and shows you what a warning looks like.

## First-open checklist

1. Open a generated workbook (e.g., `output/Plan_LR_Workbook_2027_Property.xlsx`) and let
   Excel calculate (automatic).
2. The Checks KPI card on `Control` must not be red — a fresh sample workbook reads
   `PASS WITH 1 WARNING(S)` in amber.
3. With sample data intact, the `Bridge` shows **63.4%** for `BU-A | AZ` in annual-term
   workbooks (the §9 worked example at the precision the Bridge actually displays; each
   workbook's Read Me states its own expected value), and the `Solver` returns **+5.0%** at
   4/1/2027 for its seeded target.
4. Replace the SAMPLE rows on **`Inputs`, `Rate Log` AND `Mod Log`** with your book. The Mod
   Log is easy to miss and ships with live sample actions.

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
