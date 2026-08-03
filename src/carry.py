"""Read a generated workbook's INPUTS back out, so a rebuild can keep them (D82).

Regeneration builds a brand-new workbook from sample seeds and saves it over
whatever was there. Before this module that was the whole story: a mid-season
config change — add a state, rename a BU, resize a log — silently destroyed
every pasted cell, and there was no way back except version control. The one
question this answers is "can I rebuild without losing my book?".

Everything is read BY DEFINED NAME rather than by cell address. The names are
the workbook's own stable contract (lr_bu, rl_eff, ml_status, se_block,
nr_TrendDefault...); addresses move whenever a layout decision lands, and this
reader has to work against files built by an OLDER generator than itself.

Nothing here needs the source file to be recalculated: every cell read is a
literal the user typed or pasted, so a workbook saved from Excel five seconds
ago carries forward exactly as well as one that has been through the harness.
"""

from __future__ import annotations

import datetime as dt
from dataclasses import dataclass, field
from pathlib import Path

import openpyxl


@dataclass(frozen=True)
class CarriedInputs:
    """Everything a rebuild needs to reproduce the user's book."""

    lr_rows: list[dict]
    rate_rows: list[dict]
    mod_rows: list[dict]
    season_rows: list[dict]
    term_months: int | None = None
    season_on: str | None = None
    mod_master: str | None = None
    trend_default: float | None = None
    sel_bu: str | None = None
    sel_state: str | None = None
    source: Path | None = None
    missing: tuple[str, ...] = field(default_factory=tuple)

    def summary(self) -> str:
        return (f"{len(self.lr_rows)} tbl_LR rows, {len(self.rate_rows)} rate-log rows, "
                f"{len(self.mod_rows)} mod-log rows, {len(self.season_rows)} seasonality rows")


class CarryError(RuntimeError):
    """The source file cannot be read as a generated workbook."""


def _cells(wb, name: str):
    """Flat list of cells behind a defined name, or None when it is absent."""
    if name not in wb.defined_names:
        return None
    out = []
    for sheet, ref in wb.defined_names[name].destinations:
        if sheet not in wb.sheetnames:
            return None
        got = wb[sheet][ref.replace("$", "")]
        if not isinstance(got, tuple):          # single cell
            out.append(got)
            continue
        for row in got:
            out.extend(row if isinstance(row, tuple) else (row,))
    return out


def _col(wb, name: str) -> list:
    cells = _cells(wb, name)
    return [] if cells is None else [c.value for c in cells]


def _scalar(wb, name: str):
    cells = _cells(wb, name)
    return None if not cells else cells[0].value


def _date(v):
    """Excel gives datetimes; the engine wants dates. Serials stay numbers."""
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    return v


def _clean(v):
    """Trim strings; treat whitespace-only as blank."""
    if isinstance(v, str):
        v = v.strip()
        return v or None
    return v


def read_inputs(path: Path | str) -> CarriedInputs:
    """Read every input table out of a generated workbook.

    Rows are kept when they carry a BU and a state — the same "populated"
    test the workbook's own helpers use. Blank spare rows are dropped, so a
    rebuild at a different capacity simply re-lays what is really there.
    """
    path = Path(path)
    if not path.exists():
        raise CarryError(f"carry-forward source does not exist: {path}")
    try:
        wb = openpyxl.load_workbook(path, data_only=False)
    except Exception as e:  # noqa: BLE001 - surface a usable message, not a traceback
        raise CarryError(f"cannot open {path.name}: {e}") from e

    try:
        missing = []

        def need(name: str) -> list:
            got = _col(wb, name)
            if not got:
                missing.append(name)
            return got

        # ---- tbl_LR ----------------------------------------------------
        lr_cols = {
            "bu": "lr_bu", "state": "lr_state", "lr_proj": "lr_lrproj",
            "basis": "lr_basis", "s": "lr_s", "target": "lr_target",
            "m_ind": "lr_mind",
            "m_prior": "lr_mprior", "m0": "lr_m0", "m0_asof": "lr_m0asof",
            "m_end_prior": "lr_mendprior", "m1": "lr_m1", "m2": "lr_m2",
            "ep": "lr_ep", "trend": "lr_trend", "a_other": "lr_aother",
            "a_other_label": "lr_aotherlbl", "modadj": "lr_modadj",
            "netp": "lr_netp", "netp1": "lr_netp1",
        }
        data = {k: need(n) for k, n in lr_cols.items()}
        n_lr = max((len(v) for v in data.values()), default=0)
        lr_rows = []
        for i in range(n_lr):
            row = {k: _clean(v[i]) if i < len(v) else None for k, v in data.items()}
            if not (row["bu"] and row["state"]):
                continue
            row["m0_asof"] = _date(row["m0_asof"])
            lr_rows.append(row)

        # ---- Rate Log --------------------------------------------------
        rl_cols = {"bu": "rl_bu", "state": "rl_state", "eff": "rl_eff",
                   "filed": "rl_filed", "status": "rl_status",
                   "considered": "rl_cons", "achievement": "rl_ach",
                   "comment": "rl_comment"}
        rl = {k: _col(wb, n) for k, n in rl_cols.items()}
        for k in ("bu", "state", "eff", "filed", "status"):
            if not rl[k]:
                missing.append(rl_cols[k])
        rate_rows = []
        for i in range(max((len(v) for v in rl.values()), default=0)):
            row = {k: _clean(v[i]) if i < len(v) else None for k, v in rl.items()}
            if not (row["bu"] and row["state"] and row["eff"] is not None):
                continue
            row["eff"] = _date(row["eff"])
            rate_rows.append(row)

        # ---- Mod Log (absent in pre-D70 files: carried as empty, not fatal) --
        ml_cols = {"bu": "ml_bu", "state": "ml_state", "eff": "ml_eff",
                   "chg": "ml_chg", "status": "ml_status",
                   "achievement": "ml_ach", "comment": "ml_comment"}
        ml = {k: _col(wb, n) for k, n in ml_cols.items()}
        mod_rows = []
        for i in range(max((len(v) for v in ml.values()), default=0)):
            row = {k: _clean(v[i]) if i < len(v) else None for k, v in ml.items()}
            if not (row["bu"] and row["state"] and row["eff"] is not None):
                continue
            row["eff"] = _date(row["eff"])
            mod_rows.append(row)

        # ---- seasonality ------------------------------------------------
        se_states = _col(wb, "se_state")
        se_cells = _cells(wb, "se_block") or []
        season_rows = []
        if se_states and se_cells:
            width = 12
            for i, st in enumerate(se_states):
                st = _clean(st)
                if not st:
                    continue
                weights = [c.value for c in se_cells[i * width:(i + 1) * width]]
                if len(weights) != width or all(w is None for w in weights):
                    continue
                season_rows.append({"state": st,
                                    "weights": [0.0 if w is None else float(w)
                                                for w in weights]})

        term = _scalar(wb, "nr_TermMonths")
        trend = _scalar(wb, "nr_TrendDefault")
        return CarriedInputs(
            lr_rows=lr_rows, rate_rows=rate_rows, mod_rows=mod_rows,
            season_rows=season_rows,
            term_months=int(term) if isinstance(term, (int, float)) else None,
            season_on=_clean(_scalar(wb, "nr_SeasonOn")),
            mod_master=_clean(_scalar(wb, "nr_ModAdjMaster")),
            trend_default=float(trend) if isinstance(trend, (int, float)) else None,
            sel_bu=_clean(_scalar(wb, "nr_SelBU")),
            sel_state=_clean(_scalar(wb, "nr_SelState")),
            source=path, missing=tuple(dict.fromkeys(missing)),
        )
    finally:
        wb.close()


def _sample_fingerprint(rows, rate_rows, mod_rows) -> tuple:
    """Identity of a book's inputs, coarse enough to survive formatting."""
    def num(v):
        return round(float(v), 10) if isinstance(v, (int, float)) else v

    return (
        tuple(sorted((r["bu"], r["state"], num(r["lr_proj"]), num(r["m_ind"]),
                      num(r["m0"]), num(r["m1"])) for r in rows)),
        tuple(sorted((r["bu"], r["state"], str(r["eff"]), num(r["filed"]),
                      r["status"]) for r in rate_rows)),
        tuple(sorted((r["bu"], r["state"], str(r["eff"]), num(r["chg"]),
                      r["status"]) for r in mod_rows)),
    )


def holds_sample_data(path: Path | str, cfg, lob_name: str) -> bool:
    """True when the file on disk still holds exactly the seeded sample.

    This is what makes the overwrite guard trustworthy: it compares the file's
    real input values against freshly generated seeds rather than trusting a
    baked flag. Anything the user changed — one pasted row, one edited mod —
    reads as real data and the rebuild refuses without --force.

    Reads through the same defined-name path as carry-forward, so a file it
    cannot parse is reported as NOT sample (refuse, don't guess).
    """
    from .build_workbook import (sample_lr_rows, sample_mod_rows,
                                 sample_rate_rows)

    try:
        got = read_inputs(path)
    except CarryError:
        return False
    lob = cfg.lob(lob_name)
    want = _sample_fingerprint(sample_lr_rows(cfg, lob), sample_rate_rows(cfg, lob),
                               sample_mod_rows(cfg, lob))
    mine = _sample_fingerprint(got.lr_rows, got.rate_rows, got.mod_rows)
    return want == mine
