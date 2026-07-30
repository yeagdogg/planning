# BUILD BRIEF — Calendar-Year Plan Loss Ratio Workbook (Premium Excel + Python Generator)

You are building a flagship internal actuarial planning tool for a commercial lines pricing actuary. Correctness comes first, auditability second, polish third — but all three are required. Read this entire brief before writing any code. Where this brief is silent, choose the option that is easiest to audit, and log the decision in `DECISIONS.md`.

---

## 1. Mission

Build a **config-driven Python generator** that produces a **premium Excel workbook** which converts a projected loss ratio from a rate level indication (policy-year basis, rate fully earned) into a **calendar-year plan loss ratio** (average earned rate level basis), for **multiple lines of business across multiple business units**, with dashboards showing rate and price flow through the plan year and the following year.

Deliverables (see §10): a reference engine in Python (the "oracle"), a pytest suite, the workbook generator, the generated workbook(s), a config file, and documentation.

---

## 2. Audience and quality bar

- Primary user: an actuary doing pricing/planning work. Fluent in Excel, Python-literate, works under regulatory defensibility constraints — every number must be traceable to inputs via visible, auditable formulas.
- The workbook will be shared with leadership and planning teams. It must look professionally designed, not like a default Excel file.
- The workbook must be maintainable by actuaries who did not build it: no clever tricks that trade auditability for brevity.

---

## 3. Actuarial methodology (authoritative specification)

This is the method. Do not redesign it. It is standard on-leveling / current-rate-level methodology (cf. Werner & Modlin, *Basic Ratemaking*, Ch. 5), extended with a schedule mod adjustment.

### 3.1 Definitions and timeline conventions

- **Plan year `P`**: the calendar year being planned (e.g., 2027).
- **Indication**: a rate level indication with assumed effective date **7/1/(P−1)** (e.g., 7/1/2026 indication plans CY 2027). Its projected loss ratio applies to the policy year effective 7/1/(P−1), whose average loss date is 7/1/P — the midpoint of the plan year. **Therefore no additional loss trending is applied for the plan-year result.** Document this alignment explicitly on the Methodology sheet.
- **`LR_proj`**: the projected loss ratio from the indication. May be stated at *current* rate level or at *proposed* rate level — this is a required input toggle (see §3.4).
- **`CRL_ind`** (indication rate level): the cumulative rate level embedded in the indication's premium base, expressed as an index relative to a common base. The indication assumes this level is **fully earned**.
- **Rate change log**: each rate change has an effective date, a percentage, a status (`taken` or `planned`), a **`considered_in_indication` flag** (Y/N), and an **achievement %** (applies to `planned` rows only; default 100%; effective change = filed % × achievement %).
- **Schedule mods**: multiplicative premium modifiers applied at policy issuance. Average mod drift is economically a price change that earns in on written business exactly like rate. Inputs: `M_ind` (avg mod assumed in the indication), `M_0` (current avg written mod, with an as-of date), `M_1` (projected avg written mod at 12/31/P), optional `M_prior` (avg mod ~12 months before `M_0`, see §3.3), optional `M_2` (projected at 12/31/(P+1), default = `M_1`).

### 3.2 Rate engine — monthly written-index × earning-matrix

Do **not** implement closed-form parallelogram geometry as the production engine. Implement a discretized engine that reproduces it and generalizes it:

1. **Window**: monthly writing cohorts `k` spanning **Jan (P−1) through Dec (P+1)** (36 months). Earned results are reported for **Jan P through Dec (P+1)** (24 months). Rate changes effective before the window multiply into the base index for all months (and into `CRL_ind` if flagged considered).
2. **Written rate index `W_k`**: cumulative product of `(1 + effective change)` for all changes effective on or before cohort `k`, relative to base = 1.000. **Mid-month effective dates**: split that month's written weight pro-rata by day — weight `p = (days in month on/after effective date) / (days in month)` at the post-change index, `(1−p)` at the pre-change index. Writings within a month are otherwise treated as occurring at mid-month.
3. **Earning profile** for policy term `T` months (default `T = 12`, per-line override): a cohort written in month `k` earns `1/(2T)` in month `k`, `1/T` in months `k+1 … k+T−1`, and `1/(2T)` in month `k+T` (the standard mid-month convention).
4. **Written weights `w_k`**: uniform by default. Optional per-line **seasonality vector** (12 monthly weights, normalized) to reflect renewal humps (e.g., 1/1 and 7/1 concentrations). This is a toggle; uniform is the default and reproduces the classical parallelogram.
5. **Earned index by month**: `E_m = Σ_k w_k·e(k,m)·W_k / Σ_k w_k·e(k,m)`.
6. **Calendar-year earned index** (aggregate ratio, not average of monthly ratios):
   `E_CY(Y) = Σ_{m∈Y} Σ_k w_k·e(k,m)·W_k / Σ_{m∈Y} Σ_k w_k·e(k,m)`.
7. **Indication rate level**: `CRL_ind = ∏ (1 + r_i)` over rows with `considered_in_indication = Y` (these are always `taken` rows; validate and warn if a `planned` row is flagged considered).
8. **Rate adjustment factor**: `A_rate(P) = CRL_ind / E_CY(P)`. Similarly compute `A_rate(P+1)` for the following-year view.

**Directionality checks (must hold; encode as oracle property tests):** a planned rate *increase* not considered in the indication ⇒ `A_rate < 1` (CY LR below projected). A change considered in the indication but only partially earned in CY P ⇒ pushes `A_rate > 1`. If mods rise above `M_ind` ⇒ `A_mod < 1`.

### 3.3 Schedule mod engine

1. **Written mod path `M_w(k)`**: piecewise-linear through the anchors `(as-of date, M_0)` and `(12/31/P, M_1)`, extended to `(12/31/(P+1), M_2)`. **Backward extension** (needed because earned mod in early plan-year months draws on cohorts written up to `T` months before the plan year): if `M_prior` is provided, interpolate linearly from `(as-of − 12 months, M_prior)`; otherwise extrapolate the `M_0 → M_1` line backward. Document whichever applies in-sheet.
2. **Earned mod**: `M̄_E(Y) = Σ_{m∈Y} Σ_k w_k·e(k,m)·M_w(k) / Σ_{m∈Y} Σ_k w_k·e(k,m)` — same matrix, exposure-weighted average (not a product; mods are levels, not changes).
3. **Mod adjustment factor**: `A_mod(P) = M_ind / M̄_E(P)`; similarly for P+1.
4. **Double-count guard**: a toggle `mod_adjustment_enabled` (default ON). Display a visible warning near it: *"Set OFF if the indication's premium trend already reflects schedule mod drift — otherwise the drift is double-counted."*
5. **QA identity** (encode as an oracle test under default settings): with uniform writings, annual term, and a globally linear written mod path, the CY earned mod equals the written mod at **January 1 of that CY**. Excel must match within 0.0002.

### 3.4 The bridge

1. **Basis normalization**: if the input LR basis is `proposed`, convert first: `LR_current = LR_input × (1 + s)` where `s` is the selected/indicated change from the indication (an input). If basis is `current`, `LR_current = LR_input`.
2. **Plan-year result**:
   `CY_LR(P) = LR_current × A_rate(P) × A_mod(P) × A_other`
   where `A_other` is a manual adjustment factor (default 1.000) with a required text label when ≠ 1 (e.g., commission timing, tier drift).
3. **Following-year indicative result**:
   `CY_LR(P+1) = LR_current × (1 + net_trend)^1 × A_rate(P+1) × A_mod(P+1) × A_other`
   where `net_trend` is an input (annual net loss-over-premium trend, default 0.0% with a visible caveat that the P+1 view requires it). This is labeled *indicative* — it assumes no new indication.
4. Also report, for communication: **CY average earned rate change vs. indication level** = `E_CY(P)/CRL_ind − 1`, and **earned rate change year-over-year** = `E_CY(P)/E_CY(P−1) − 1` and `E_CY(P+1)/E_CY(P) − 1` (the "carryover + new actions" number planning teams quote).

### 3.5 Conventions, edge cases, and stated assumptions

- Rate changes compound multiplicatively; enter as decimals or percents consistently (pick one, validate, and label).
- Rate changes are assumed to apply to all written premium (no new/renewal split in v1 — note as a limitation with a future hook).
- Exposure/audit premium effects (e.g., WC payroll audits) are out of scope; note as a limitation.
- Zero planned changes, zero rate history, flat mods must all compute cleanly (factors = 1.000-ish), with no `#DIV/0!`, `#REF!`, `#N/A`, or `#NAME?` anywhere under any toggle state.
- All ratios are currency-agnostic.

---

## 4. Inputs (live in the workbook, not in config)

All actuarial inputs are maintained **in the workbook** by the team. The config file defines structure only (§6). Structure inputs as proper Excel Tables:

1. **`tbl_LR`** — one row per (BU, LOB): projected LR, LR basis (`current`/`proposed`), selected indication change `s`, `M_ind`, `M_0`, `M_0` as-of date, `M_1`, optional `M_prior`, optional `M_2`, optional plan earned premium (used to weight the Portfolio total), net trend for P+1, `A_other` + label, mod adjustment toggle.
2. **`tbl_RateLog`** — one row per rate change per (BU, LOB): effective date, filed %, status (`taken`/`planned`), considered-in-indication (Y/N), achievement % (planned only), comment. Include a convenience helper: a button-free "seed" row pattern showing how to enter the indication's selected change as a planned row.
3. **`tbl_Seasonality`** (optional) — per LOB: 12 monthly written weights; blank = uniform.
4. **Control inputs** — plan year `P`, active BU, active LOB (data-validated dropdowns), scenario selector.

Every input cell: blue font, light-yellow fill for required cells, data validation where applicable, and one realistic example row per table (clearly marked as sample data to replace).

---

## 5. Workbook design — sheet by sheet

Order and names exactly as follows (visible sheets):

1. **`Read Me`** — purpose, one-paragraph method summary, how-to steps, color legend, version/change log.
2. **`Control`** — BU/LOB/plan-year selectors, toggles (LR basis display, seasonality on/off, mod adjustment), and a KPI card row: Projected LR → CY Plan LR(P), CY earned rate change, carryover into P+1, current vs. projected mod. This is the executive landing page.
3. **`Inputs`** — the tables in §4.
4. **`Rate Engine`** — the visible 36-month engine for the selected (BU, LOB): written weights, written index `W_k`, earning matrix aggregation, monthly earned index `E_m`, CY aggregates, `CRL_ind`, `A_rate(P)` and `A_rate(P+1)`. Every column labeled; helper columns preferred over mega-formulas.
5. **`Mod Engine`** — same layout for the mod path and `A_mod`.
6. **`Bridge`** — the waterfall: Projected LR (as input) → basis normalization → rate factor → mod factor → other → **CY Plan LR(P)**, each step shown in loss-ratio points; a second column for the indicative P+1 bridge. Native waterfall chart if compatible, stacked-column fallback otherwise.
7. **`Flow Dashboard`** — charts: (a) cumulative written vs. earned rate index over 36 months with the plan year shaded; (b) YoY earned rate change % by quarter for P and P+1; (c) written vs. earned mod path; (d) combined **price index** = rate index × (mod ÷ `M_ind`); (e) quarterly table of earned rate change and "unearned runway" carrying into P+1.
8. **`Portfolio`** — all BU × LOB combinations computed simultaneously: CY Plan LR(P), `A_rate`, `A_mod`, CY earned rate change, P+1 carryover; conditional-format heatmap; a total row weighted by plan earned premium where provided (simple average with a footnote otherwise). Backed by a hidden `_calc` sheet with one compact engine block per combination, generated by the script.
9. **`Scenarios`** — up to 4 side-by-side scenarios for the selected combo, each varying planned change magnitude, effective date, achievement, and mod path; delta table vs. base and a comparison chart.
10. **`Solver`** — Enhancement E1 (§7).
11. **`Attribution`** — Enhancement E2 (§7).
12. **`Checks`** — validation panel (§9): PASS/FAIL rows, green/red conditional formatting, a bold overall status cell referenced on `Control`.
13. **`Methodology`** — the regulatory-defensibility writeup: formulas as implemented, assumptions (7/1 midpoint alignment, uniform writings unless seasonality on, mod path construction, achievement, limitations), definitions of every named range, and the citation to standard on-leveling methodology.

Hidden: **`_lists`** (validation lists), **`_calc`** (portfolio engines), **`_oracle`** (baked-in oracle constants for Checks).

---

## 6. Multi-BU / multi-LOB architecture and the generator

- **`config/config.yaml`** defines: plan year, list of business units, list of LOBs (with per-LOB overrides: `term_months`, seasonality profile name), output mode (`master` | `per_bu` | `both`), formula mode (`classic` default | `modern`), and theme options.
- **`src/build_workbook.py`** — CLI: `python -m src.build_workbook --config config/config.yaml --out output/`. Generates the master workbook (all combos in the input tables, selector-driven engines, Portfolio across everything) and, if configured, per-BU copies (filtered tables, same everything else). Regenerating for a new plan year = edit config, rerun; inputs are then re-entered/pasted into the new workbook's tables (note this workflow in the README, and keep table schemas stable so paste-over works).
- **`src/engine.py`** — the oracle: a clean, documented Python implementation of §3 (both the monthly-convention engine and the continuous closed-form parallelogram for cross-checking). This is the source of truth.
- Selector mechanics: `INDEX`/`MATCH` against the tables keyed on a concatenated `BU|LOB` key column. No volatile functions (`OFFSET`, `INDIRECT`, `TODAY` in calc paths).

---

## 7. Enhancement modules (build both)

**E1 — Inverse Plan Solver (`Solver` sheet).** Because the CY earned index is **linear** in a single unknown planned change `r` (with everything else fixed), solve in closed form — no Goal Seek, no iteration:
- Compute `C_pre = Σ w·e·W⁰` over cohorts written before the chosen effective date and `C_post = Σ w·e·W⁰` over cohorts on/after it (within CY P sums), where `W⁰` is the index with the unknown change excluded. Then `E_CY = C_pre + (1+r)·C_post`.
- **Mode A (rate for target)**: input target CY LR and effective date → required `A_rate = target / (LR_current × A_mod × A_other)` → `E_needed = CRL_ind / A_rate` → `r = (E_needed − C_pre)/C_post − 1`. Display filed-rate equivalent given achievement %.
- **Mode B (timing for rate)**: input `r` → tabulate resulting CY LR for effective dates on the 1st of each month of P → chart "CY LR by effective month" and report the latest month that still meets target.
- Guard rails: warn when `C_post` is tiny (late-year dates make targets mathematically absurd) and when required `r` exceeds a configurable reasonability bound.

**E2 — Plan-vs-Actual Attribution (`Attribution` sheet).** Inputs: actual achieved change % and actual effective date per planned row, actual mod path anchors, actual CY LR. Compute a sequential multiplicative decomposition of (actual − plan) CY LR, in points: (1) **rate magnitude** (achieved % at planned dates), (2) **rate timing** (achieved % at actual dates), (3) **mod drift** (actual vs. projected path), (4) **loss-side residual** (whatever remains). Waterfall chart. Document that the decomposition is order-dependent and state the order used.

Future hooks to note (do not build): new/renewal split with differential rate, retention response to rate, premium-volume plan (EP × LR → plan losses), Snowflake/Power Query input refresh.

---

## 8. Engineering and formatting standards

**Formulas**
- Formulas, never hardcoded results: the workbook must recalculate when inputs change. No number appears in a formula that isn't a named cell or table reference.
- **Formula mode `classic` (default)**: only Excel-2007-era functions — `SUMPRODUCT`, `SUMIFS`, `INDEX`, `MATCH`, `IF`, `IFERROR` (sparingly, with a cell comment justifying each). No `XLOOKUP`, `XMATCH`, `FILTER`, `SORT`, `UNIQUE`, `SEQUENCE`, `LET`, `LAMBDA` — this guarantees LibreOffice recalculation for validation and compatibility with any corporate Excel build. If you use `TEXTJOIN`/`IFS`/`SWITCH`/`CONCAT`/`MAXIFS`/`MINIFS`, write them with the `_xlfn.` prefix. Mode `modern` may be offered behind the config flag but `classic` ships as default.
- Prefer helper columns over nested mega-formulas; target ≤ ~120 characters per formula. Auditability beats cleverness.
- Named ranges for every scalar the Bridge uses; all documented on `Methodology`. Quote sheet names containing spaces in references. No merged cells in calculation areas. No external workbook links. No circular references. Automatic calculation.

**Formatting (financial-model conventions)**
- Blue font = hardcoded inputs/levers; black = formulas; green = links from another sheet; yellow fill = cells the user must fill.
- Percentages stored as fractions, formatted `0.0%`; loss ratios `0.0%`; indices `0.0000`; factors `0.0000`; points shown as `0.00 "pts"`; dates as real dates `mm/dd/yyyy`; years as text.
- One professional font family throughout (Arial or Calibri), a restrained palette (e.g., navy `#1F3864` headers, steel `#8497B0` accents, white/near-white panels), consistent borders, freeze panes on every data sheet, gridlines off on `Control`, `Bridge`, `Flow Dashboard`, `Portfolio`; print-ready page setup with headers/footers on presentation sheets.
- Charts share one style: titled, axis-labeled, no default chart junk, plan year visually distinguished.
- No VBA/macros. Sheet protection optional and, if applied, without passwords (state so in Read Me).

---

## 9. Validation and QA (mandatory)

**Oracle-first development.** Build `src/engine.py` and `tests/test_engine.py` before any Excel. Tests must include the worked example below, the directionality properties (§3.2), the linear-mod identity (§3.3), monthly-vs-continuous convergence (agreement within tolerance under uniform/annual settings), and degenerate inputs (no changes, flat mods).

**Worked example (must reproduce).** Plan year P = 2027; indication effective 7/1/2026; `LR_proj = 65.0%` at **current** rate level; `M_ind = 0.850`. Rate log: +4.0% eff 2025-07-01 (taken, considered = Y); +6.0% eff 2026-07-01 (taken, considered = Y); +5.0% eff 2027-04-01 (planned, considered = N, achievement 100%). Mods: `M_0 = 0.860` as of 2026-09-30; `M_1 = 0.890` at 2027-12-31; linear path extrapolated backward; mod adjustment ON; `A_other = 1.000`; uniform writings; annual term.

Continuous closed-form expected values (oracle must match to 1e-6; the monthly-convention engine and the Excel workbook must match the oracle's monthly values exactly, and the continuous values within ±0.10 loss-ratio points):
- `CRL_ind = 1.1024`
- CY 2027 earned rate index `E_CY = 1.1101025` (areas: 0.125 @ 1.0400, 0.59375 @ 1.1024, 0.28125 @ 1.157520)
- `A_rate = 1.1024 / 1.1101025 = 0.99306`
- CY 2027 earned mod `M̄_E = 0.8660` (linear-mod identity: written mod at 1/1/2027) → `A_mod = 0.850/0.866 = 0.98152`
- **CY 2027 plan LR = 65.0% × 0.99306 × 0.98152 = 63.36%**

**Workbook verification.** After generating: recalculate the file headlessly with LibreOffice (`soffice --headless --convert-to xlsx` round-trip, or an equivalent recalc script) and re-read cached values with `openpyxl` (`data_only=True` — note openpyxl-written formulas read back as `None` until recalculated). Zero formula errors anywhere, under every toggle state you can exercise by writing test inputs. Verify the Bridge output for the worked example matches the oracle. If LibreOffice is unavailable in the environment, say so explicitly and ship a "first-open checklist" in the README instructing the user to open the file and confirm the `Checks` sheet shows all PASS.

**In-workbook `Checks` panel** (live formulas vs. `_oracle` constants baked at build time): worked-example Bridge output ties to oracle; earning-matrix column sums equal 1 per cohort within the reporting window; earned index between min and max written index each month; `CRL_ind` ties to the product of flagged rows; mod identity check under default settings; input sanity (dates within window, achievement ∈ [0%,150%], mods ∈ [0.5, 1.5], considered flags only on taken rows); overall PASS/FAIL surfaced on `Control`.

---

## 10. Deliverables

```
config/config.yaml
src/engine.py            # oracle: monthly engine + continuous parallelogram cross-check
src/build_workbook.py    # generator CLI
tests/test_engine.py     # pytest incl. worked example + property tests
output/Plan_LR_Workbook_<P>.xlsx   (+ per-BU files if configured)
README.md                # setup, regeneration workflow, first-open checklist
DECISIONS.md             # every judgment call you made and why
```

Seed the config with: business units `BU-A`, `BU-B`, `BU-C`; LOBs `Property`, `General Liability`, `Commercial Auto`, `Workers Comp`, `Umbrella`, `Inland Marine` (annual term; make one LOB demonstrate a 6-month term override to prove the parameter works); plan year 2027. The user will rename to real values.

## 11. Acceptance criteria (verify each; report status at the end)

1. Oracle pytest suite passes, including the worked example to stated precision.
2. Generated workbook recalculates with **zero** formula errors (LibreOffice-verified, or checklist shipped with an explicit note).
3. Worked example entered in the workbook reproduces CY LR = 63.36% within ±0.10 pts of continuous (exact vs. oracle monthly values).
4. All toggles (LR basis, seasonality, mod adjustment, term override) change results in the correct direction, with no errors in any state.
5. Selector switches BU/LOB and every visible engine, bridge, and chart follows.
6. Portfolio sheet shows all combos, heatmapped, with an EP-weighted total when EP is provided.
7. Solver Mode A reproduces: solving for the rate that yields the worked example's own CY LR returns +5.0% at 4/1/2027.
8. Attribution waterfall reconciles: factors multiply back to the actual CY LR entered.
9. Flow Dashboard renders all five visuals for P and P+1.
10. No volatile functions, no dynamic-array functions in classic mode, no external links, no VBA, no merged cells in calc areas.
11. Formatting conventions of §8 applied throughout; `Read Me`, `Methodology`, and legends complete.
12. `DECISIONS.md` lists every deviation or judgment call.

## 12. Defaults (proceed without asking; log in DECISIONS.md)

| Question | Default |
|---|---|
| Excel compatibility | `classic` formula mode (works everywhere; LibreOffice-verifiable) |
| VBA | None |
| Output mode | `master` workbook (per-BU emit available via config) |
| Policy term | 12 months (per-LOB override; one seeded 6-month example) |
| Writing pattern | Uniform (per-LOB seasonality vector optional) |
| Achievement % | 100% on planned rows |
| Mod adjustment | ON, with double-count warning |
| P+1 net trend | 0.0% with visible caveat |
| Rate entry format | Percent-formatted fractions (e.g., 5.0% stored as 0.05) |

Ask the user only if something is genuinely blocking; otherwise build with these defaults.

## 13. Build order

1. **Plan**: restate the method in your own words in `DECISIONS.md`; confirm the worked example by hand/oracle before touching Excel.
2. **Oracle + tests**: `engine.py`, full pytest green.
3. **Minimal workbook**: one BU/LOB, Inputs → Rate Engine → Mod Engine → Bridge; tie to oracle via recalc; fix until exact.
4. **Scale out**: selector, Portfolio `_calc` blocks, Scenarios, Solver, Attribution, Flow Dashboard.
5. **Checks + Methodology + Read Me.**
6. **Format** to §8, regenerate, re-verify (recalc + re-read), run the acceptance list, and summarize results with any deviations.
