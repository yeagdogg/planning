"""W4a: the targets file — the formula must BE the engine, not resemble it.

The load-bearing test is the first one: build the file, evaluate the model
the Excel cells encode at random target pairs, and tie every value to a
fresh ``engine.run_bridge``. Everything else guards the contract around
that — the structural residual that must kill a build, the blank/carry
semantics the field user relies on, the banned-function law, the named
ranges the collect step reads, and the non-interference rule that keeps a
targets file from ever being imported as a whole line.

All of it is openpyxl + engine, so both interpreters run it; one live
Excel test rides in tests/test_live.py, the only file that starts Excel.
"""
from __future__ import annotations

import random
import sys
from dataclasses import replace
from pathlib import Path

import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"


@pytest.fixture(scope="module")
def scn():
    from app import importers
    return importers.from_workbook(WB_PROP)


@pytest.fixture(scope="module")
def built(scn):
    import datetime as dt

    from app import targetbook
    wb, bakes = targetbook.build_targets_wb(
        scn, now=dt.datetime(2026, 8, 8, 9, 30))
    return wb, bakes


def _tabs(wb):
    from app import targetbook
    return [s for s in wb.sheetnames
            if s not in (targetbook.COVER, targetbook.FACTORS)]


# ------------------------------------------------------ the tie that matters

def test_fit_ties_engine_at_random_targets(scn, built):
    """The whole promise: whatever a field actuary types, the sheet shows
    the number the engine would produce. 20 seeded pairs per combo across
    the whole line, including the blank-P+1 carry."""
    from app import compute, targetbook
    from src import engine
    _wb, bakes = built
    rng = random.Random(20260808)
    worst = 0.0
    for b in bakes:
        assert b.ok, f"{b.key}: {b.error}"
        ci = compute.combo_inputs(scn, b.row)
        for _ in range(20):
            x = rng.uniform(-0.25, 0.35)
            x1 = rng.choice([None, rng.uniform(-0.25, 0.35)])
            res = engine.run_bridge(
                scn.plan_year, replace(ci, net_sel_p=x, net_sel_p1=x1),
                "monthly")
            worst = max(worst,
                        abs(b.fit.lr_p(x) - res.cy_lr_p),
                        abs(b.fit.lr_p1(x, x1) - res.cy_lr_p1))
    assert worst < 1e-9, f"model drifted from the engine by {worst:.3e}"
    assert isinstance(targetbook.FORMAT_VERSION, int)


def test_blank_target_falls_back_to_the_logged_program(scn, built):
    """Blank is not zero: it means 'no net selection', and the sheet must
    show what the logged program delivers on its own."""
    from app import compute, targetbook
    from src import engine
    _wb, bakes = built
    for b in bakes[:8]:
        ci = compute.combo_inputs(scn, b.row)
        want = engine.program_basis_plan_lr(scn.plan_year, ci)
        assert b.prog_p == pytest.approx(want[0], abs=1e-12)
        assert b.prog_p1 == pytest.approx(want[1], abs=1e-12)
        # ...and a 0% target is a DIFFERENT answer (net mode ON, flat)
        flat = engine.run_bridge(
            scn.plan_year, replace(ci, net_sel_p=0.0, net_sel_p1=0.0),
            "monthly")
        assert b.fit.lr_p(0.0) == pytest.approx(flat.cy_lr_p, abs=1e-9)
    assert targetbook.FACTORS.startswith("_")


def test_structural_break_fails_the_build_loudly(scn, monkeypatch):
    """A residual at the check point is not rounding — it means the
    engine's net path changed shape. The build must refuse to ship a
    workbook whose formulas would then be wrong."""
    from app import targetbook
    from src import engine as eng

    real = eng.run_bridge
    check_x = targetbook._CHECK[0]

    def bent(plan_year, ci, mode="monthly", *a, **k):
        res = real(plan_year, ci, mode, *a, **k)
        if ci.net_sel_p == check_x:      # the falsification point only —
            # bend it and the fitted line no longer passes through it,
            # which is exactly what a changed engine would look like
            return replace(res, cy_lr_p=res.cy_lr_p * 1.01)
        return res

    monkeypatch.setattr(targetbook.engine, "run_bridge", bent)
    with pytest.raises(targetbook.TargetsBuildError, match="no longer an exact"):
        targetbook.build_targets_wb(scn)


def test_formula_text_carries_the_blank_and_carry_rules(built):
    """The two semantics a field user leans on live in the cell text: a
    blank P target gates BOTH years to the program constants, and a blank
    P+1 carries P — the engine's own net_x1 rule, in Excel."""
    from app import targetbook
    wb, _bakes = built
    ws = wb[_tabs(wb)[0]]
    r = targetbook.FIRST_ROW
    j, k = ws[f"J{r}"].value, ws[f"K{r}"].value
    assert j.startswith(f'=IF($H{r}="",{targetbook.FACTORS}!')
    assert k.startswith(f'=IF($H{r}="",{targetbook.FACTORS}!')
    assert f'(1+IF($I{r}="",$H{r},$I{r}))' in k       # the carry
    assert f"(1+$H{r})" in j and f"(1+$H{r})" in k
    # a +1 target alone is a no-op, and the sheet says so rather than
    # silently ignoring it
    assert "alone does nothing" in ws[f"N{r}"].value


def test_no_banned_functions_and_balanced_parens(built):
    """The fleet's formula law applies here too: nothing volatile, nothing
    dynamic-array, and above all nothing that SWALLOWS an error."""
    from tools.verify_workbook import BANNED_RE

    from src.xlstyle import assert_formulas_balanced
    wb, _bakes = built
    hits = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for c in row:
                v = c.value
                if isinstance(v, str) and v.startswith("="):
                    m = BANNED_RE.search(v)
                    if m:
                        hits.append(f"{ws.title}!{c.coordinate} {m.group(0)}")
    assert not hits, hits
    assert_formulas_balanced(wb)


def test_factors_schema_and_residuals(built):
    """The hidden machine area is read by column SCHEMA, so the tests and
    the workbook cannot drift apart."""
    from app import targetbook
    wb, bakes = built
    fx = wb[targetbook.FACTORS]
    cols = targetbook.FACTORS_COLS
    for i, b in enumerate(bakes):
        r = targetbook.FACT_FIRST + i
        assert fx.cell(row=r, column=cols["key"]).value == b.key
        assert fx.cell(row=r, column=cols["c0"]).value == pytest.approx(
            b.fit.c0, abs=1e-15)
        assert fx.cell(row=r, column=cols["prog_p"]).value == pytest.approx(
            b.prog_p, abs=1e-15)
        assert fx.cell(row=r, column=cols["status"]).value == "ok"
        assert max(b.fit.resid_p, b.fit.resid_p1) < 1e-11
    # the date tokens the MATCH resolves against
    toks = [fx.cell(row=2, column=cols["a1"] + j).value
            for j in range(len(targetbook.DATE_TOKENS))]
    assert tuple(toks) == targetbook.DATE_TOKENS


def test_helper_constants_tie_suggest_net_rate(scn, built):
    """The suggested change is the engine's own closed form, baked."""
    import datetime as dt

    from app import compute, targetbook
    from src import engine
    _wb, bakes = built
    checked = 0
    for b in bakes[:10]:
        ci = compute.combo_inputs(scn, b.row)
        for j, m in enumerate(targetbook.DATE_MONTHS):
            if not b.helper.b[j]:
                continue                       # no exposure after that date
            x = 0.06
            want, _c = engine.suggest_net_rate(
                scn.plan_year, replace(ci, net_sel_p=x, net_sel_p1=x),
                dt.date(scn.plan_year, m, 1))
            assert b.helper.suggested(x, j) == pytest.approx(want, abs=1e-9)
            checked += 1
    assert checked, "no date had exposure — the helper was never exercised"


def test_names_cover_the_read_back(built):
    """The collect step reads by defined name; the true BU string must be
    recoverable even though the tab name is sanitized."""
    from app import targetbook
    wb, bakes = built
    names = set(wb.defined_names)
    for scalar in ("tgt_Format", "tgt_LOB", "tgt_PlanYear", "tgt_Generated",
                   "tgt_Fingerprint"):
        assert scalar in names
    bus = {b.bu for b in bakes}
    found = set()
    for nm in names:
        if not nm.startswith("tgt_bu_"):
            continue
        sfx = nm[len("tgt_bu_"):]
        for part in ("state", "netp", "netp1", "date", "lrp", "lrp1", "sugg"):
            assert f"tgt_{part}_{sfx}" in names, f"missing tgt_{part}_{sfx}"
        (sheet, ref), = wb.defined_names[nm].destinations
        found.add(wb[sheet][ref.replace("$", "")].value)
    assert found == bus
    # the state column matches the roster, in order
    for nm in names:
        if not nm.startswith("tgt_state_"):
            continue
        (sheet, ref), = wb.defined_names[nm].destinations
        got = [c.value for row in wb[sheet][ref.replace("$", "")]
               for c in row]
        bu = wb[sheet][
            list(wb.defined_names[f"tgt_bu_{nm[len('tgt_state_'):]}"]
                 .destinations)[0][1].replace("$", "")].value
        assert got == [b.state for b in bakes if b.bu == bu]
    assert targetbook.FIRST_ROW == 6


def test_only_the_target_cells_are_writable(built):
    """The file's integrity mechanism: everything except the three input
    cells is locked, passwordless — guidance, not security."""
    from app import targetbook
    wb, bakes = built
    for name in _tabs(wb):
        ws = wb[name]
        assert ws.protection.sheet is True
        r = targetbook.FIRST_ROW
        for c in ("H", "I", "L"):
            assert ws[f"{c}{r}"].protection.locked is False, f"{name}!{c}{r}"
        for c in ("A", "B", "E", "J", "K", "M"):
            assert ws[f"{c}{r}"].protection.locked is not False
        # inputs read as inputs: blue font on a yellow fill (the house rule)
        from src.xlstyle import INPUT_BLUE, YELLOW_REQ
        assert INPUT_BLUE in (ws[f"H{r}"].font.color.rgb or "")
        assert YELLOW_REQ in (ws[f"H{r}"].fill.fgColor.rgb or "")
    assert wb[targetbook.FACTORS].sheet_state == "hidden"


def test_prefilled_targets_round_trip(scn, built):
    """An untouched file must diff to nothing later, so the input cells
    open holding the book's current targets — including the blanks."""
    from app import targetbook
    wb, bakes = built
    for i, b in enumerate(bakes):
        ws = wb[[s for s in _tabs(wb)
                 if wb[s]["Q1"].value == b.bu][0]]
        rows = [x for x in bakes if x.bu == b.bu]
        r = targetbook.FIRST_ROW + rows.index(b)
        assert ws[f"H{r}"].value == b.row.get("netp")
        assert ws[f"I{r}"].value == b.row.get("netp1")


def test_validation_bounds_and_the_date_list(built):
    from app import targetbook
    wb, _bakes = built
    ws = wb[_tabs(wb)[0]]
    kinds = {dv.type: dv for dv in ws.data_validations.dataValidation}
    assert "decimal" in kinds and "list" in kinds
    dec = kinds["decimal"]
    assert (dec.formula1, dec.formula2) == ("-0.5", "0.5")
    assert "fractions" in dec.error
    assert all(t in kinds["list"].formula1 for t in targetbook.DATE_TOKENS)


def test_the_date_column_is_text_formatted(built):
    """Excel parses "7/1" into a DATE the moment it is typed or picked from
    the dropdown, and the MATCH against the token row would then never hit
    — the live Excel test caught exactly that. Text format is the fix, and
    this pins it (no python-level evaluation can see the coercion)."""
    from app import targetbook
    wb, _bakes = built
    for name in _tabs(wb):
        ws = wb[name]
        assert ws[f"L{targetbook.FIRST_ROW}"].number_format == "@", name


def test_a_rejected_combo_dashes_and_is_listed(scn):
    """An engine rejection is a DATA failure: dash that row, name it on
    the cover, and keep the rest of the file useful."""
    import dataclasses

    from app import targetbook
    rows = [dict(r) for r in scn.lr_rows]
    rows[0] = dict(rows[0], lr_proj="not-a-number")
    hurt = dataclasses.replace(scn, lr_rows=rows)
    wb, bakes = targetbook.build_targets_wb(hurt)
    bad = [b for b in bakes if not b.ok]
    assert len(bad) == 1 and bad[0].key == f"{rows[0]['bu']}|{rows[0]['state']}"
    ws = wb[[s for s in _tabs(wb) if wb[s]["Q1"].value == bad[0].bu][0]]
    r = targetbook.FIRST_ROW
    assert ws[f"J{r}"].value == "—" and ws[f"K{r}"].value == "—"
    assert ws[f"H{r}"].value is None                 # not an input any more
    cover = "\n".join(str(c.value) for row in wb[targetbook.COVER].iter_rows()
                      for c in row if c.value)
    assert bad[0].key in cover


def test_fingerprint_tracks_inputs_not_targets(scn):
    """Targets are the point of the exercise, so they are NOT part of the
    fingerprint; anything that moves the baked math IS."""
    import dataclasses

    from app import targetbook
    base = targetbook.fingerprint_line(scn)
    rows = [dict(r) for r in scn.lr_rows]
    rows[0] = dict(rows[0], netp=0.25, netp1=0.05)
    assert targetbook.fingerprint_line(
        dataclasses.replace(scn, lr_rows=rows)) == base
    rows[0] = dict(rows[0], lr_proj=float(rows[0]["lr_proj"]) + 0.01)
    assert targetbook.fingerprint_line(
        dataclasses.replace(scn, lr_rows=rows)) != base
    rl = [dict(r) for r in scn.rate_rows]
    rl[0] = dict(rl[0], filed=(rl[0].get("filed") or 0.0) + 0.01)
    assert targetbook.fingerprint_line(
        dataclasses.replace(scn, rate_rows=rl)) != base


def test_build_all_writes_one_file_per_line_and_stays_unimportable(tmp_path):
    """A targets file must never be mistaken for a fleet workbook: read as
    a whole line it would replace that line with emptiness."""
    from app import importers, targetbook
    book = {"Property": importers.from_workbook(WB_PROP),
            "General Liability": importers.from_workbook(WB_GL)}
    paths = targetbook.build_all(book, out_dir=tmp_path)
    assert [p.name for p in paths] == [
        "Plan_LR_Targets_2027_Property.xlsx",
        "Plan_LR_Targets_2027_General_Liability.xlsx"]
    for p in paths:
        assert p.exists() and p.stat().st_size > 0
        with pytest.raises(ValueError):
            importers.from_workbook(p)
    assert not any(str(p).endswith(x) for p in paths
                   for x in importers.workbook_choices().values())


def test_duplicate_combo_is_a_build_error(scn):
    import dataclasses

    from app import targetbook
    rows = [dict(r) for r in scn.lr_rows]
    rows.append(dict(rows[0]))
    with pytest.raises(targetbook.TargetsBuildError, match="duplicate combo"):
        targetbook.build_targets_wb(
            dataclasses.replace(scn, lr_rows=rows))
