"""Workbook -> Scenario: the day-one on-ramp.

src/carry.py::read_inputs already knows how to read every input table out of
a generated workbook BY DEFINED NAME (it is what --carry-forward uses), so
the app opens John's real books with zero new parsing. Global toggles the
workbook carries (term override, seasonality on/off, mod master, trend
default) ride along.
"""
from __future__ import annotations

import functools
import re
from pathlib import Path

from src.build_workbook import load_config as _load_config
from src.carry import read_inputs

from .state import Scenario

ROOT = Path(__file__).resolve().parents[1]
CONFIG_PATH = ROOT / "config" / "config.yaml"


@functools.lru_cache(maxsize=1)
def app_config():
    """The repo config (plan year, LOB roster, seasonality profiles) —
    loaded once per process; the app never mutates it."""
    return _load_config(CONFIG_PATH)


def _lob_for_filename(path: Path):
    """Which configured LOB produced this file? Matches the config's
    filename template, tolerating pilot tags (_MODERN etc.)."""
    cfg = app_config()
    stem = path.stem
    for lob in cfg.lobs:
        expected = cfg.filename.format(plan_year=cfg.plan_year,
                                       lob=lob.name.replace(" ", "_"))
        if stem == expected or stem.startswith(expected + "_"):
            return lob
    return None


def workbook_choices() -> dict:
    """{display label: path} for the Open-workbook picker — every generated
    workbook in output/, newest first, minus the A/B scratch twin."""
    cfg = app_config()
    out = ROOT / cfg.output_dir
    found = sorted(out.glob("*.xlsx"), key=lambda p: p.stat().st_mtime,
                   reverse=True)
    choices = {}
    for p in found:
        if p.stem.endswith("_ABBASE") or re.search(r"\.bak$", p.stem):
            continue
        lob = _lob_for_filename(p)
        if lob is None:
            continue                      # the Book, backups, foreign files
        choices[p.name] = str(p)
    return choices


def fleet_choices() -> dict:
    """{LOB name: path} — ONE workbook per configured line, in config
    order: the untagged fleet file when present, else the newest tagged
    twin (_MODERN etc.). This is the 'load every line' on-ramp."""
    cfg = app_config()
    out = ROOT / cfg.output_dir
    found = sorted(out.glob("*.xlsx"), key=lambda p: p.stat().st_mtime,
                   reverse=True)
    best: dict = {}
    for p in found:
        if p.stem.endswith("_ABBASE"):
            continue
        lob = _lob_for_filename(p)
        if lob is None:
            continue
        expected = cfg.filename.format(plan_year=cfg.plan_year,
                                       lob=lob.name.replace(" ", "_"))
        exact = p.stem == expected
        cur = best.get(lob.name)
        if cur is None or (exact and not cur[0]):
            best[lob.name] = (exact, p)
    return {lob.name: str(best[lob.name][1]) for lob in cfg.lobs
            if lob.name in best}


def from_workbook(path) -> Scenario:
    """Read a generated workbook's inputs into a Scenario."""
    path = Path(path)
    lob = _lob_for_filename(path)
    if lob is None:
        raise ValueError(
            f"{path.name} does not match any configured LOB workbook")
    cfg = app_config()
    ci = read_inputs(path)
    return Scenario(
        lob=lob.name,
        plan_year=cfg.plan_year,
        term_months=int(ci.term_months or lob.term_months),
        lr_rows=[dict(r) for r in ci.lr_rows],
        rate_rows=[dict(r) for r in ci.rate_rows],
        mod_rows=[dict(r) for r in ci.mod_rows],
        season_rows=[dict(r) for r in ci.season_rows],
        season_on=(str(ci.season_on or "OFF").upper() == "ON"),
        mod_master=(str(ci.mod_master or "ON").upper() != "OFF"),
        trend_default=float(ci.trend_default or 0.0),
        name=path.stem,
        source=str(path),
    )
