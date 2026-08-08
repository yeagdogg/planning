"""Scenario files — the whole book as ONE yaml document, carrying its own
audit trail (P5).

The hub's contract: the book (every line's Scenario) saves to a single
``scenarios/*.yaml``. Saving over an existing file DIFFS against what it
replaces and appends a changelog entry — timestamp, note, and the
sentence-level changes ("Property · ABD|AZ: Net rate P 2.0% → 2.5%") —
so the file tells its own story the way the workbook's logs do. Version
control is git's job: the scenarios/ folder is meant to be committed.

UI-free and yaml-native: dates ride as real dates (safe_dump/safe_load
round-trip ``datetime.date``), rows are the same dicts every other
surface uses.
"""
from __future__ import annotations

import datetime as dt
from pathlib import Path

import yaml

from .grids import LR_SCHEMA, ML_SCHEMA, RL_SCHEMA
from .state import Scenario

ROOT = Path(__file__).resolve().parents[1]
SCEN_DIR = ROOT / "scenarios"
FORMAT = 1

_TOGGLES = [("term_months", "term (months)"),
            ("season_on", "seasonality"),
            ("mod_master", "mod adjustment master"),
            ("trend_default", "net trend default")]


# ------------------------------------------------------------- doc <-> book

def book_to_doc(book: dict, name: str) -> dict:
    lines = {}
    for lob, scn in book.items():
        lines[lob] = dict(
            term_months=int(scn.term_months),
            season_on=bool(scn.season_on),
            mod_master=bool(scn.mod_master),
            trend_default=float(scn.trend_default or 0.0),
            lr_rows=[dict(r) for r in scn.lr_rows],
            rate_rows=[dict(r) for r in scn.rate_rows],
            mod_rows=[dict(r) for r in scn.mod_rows],
            season_rows=[dict(r) for r in scn.season_rows],
        )
    plan_year = next((s.plan_year for s in book.values()), None)
    return {"plan_workbench_scenario": FORMAT, "name": name,
            "plan_year": plan_year, "lines": lines}


def doc_to_book(doc: dict, source: str = "") -> dict:
    """Yaml doc -> {line: Scenario}. Unknown lines are a loud error — a
    typo'd line would otherwise surface later as engine failures."""
    from . import importers
    cfg = importers.app_config()
    known = {lob.name for lob in cfg.lobs}
    lines = doc.get("lines") or {}
    unknown = sorted(set(lines) - known)
    if unknown:
        raise ValueError(
            f"scenario file names lines the config does not know: "
            f"{', '.join(unknown)} (configured: {', '.join(sorted(known))})")
    plan_year = int(doc.get("plan_year") or cfg.plan_year)
    name = doc.get("name") or "scenario"
    book = {}
    for lob, d in lines.items():
        book[lob] = Scenario(
            lob=lob, plan_year=plan_year,
            term_months=int(d.get("term_months")
                            or cfg.lob(lob).term_months),
            lr_rows=[dict(r) for r in d.get("lr_rows") or []],
            rate_rows=[dict(r) for r in d.get("rate_rows") or []],
            mod_rows=[dict(r) for r in d.get("mod_rows") or []],
            season_rows=[dict(r) for r in d.get("season_rows") or []],
            season_on=bool(d.get("season_on")),
            mod_master=bool(d.get("mod_master", True)),
            trend_default=float(d.get("trend_default") or 0.0),
            name=f"{name} — {lob}", source=source)
    return book


# ------------------------------------------------------------------- diffs

def _fmt(kind: str, v) -> str:
    if v is None or v == "":
        return "—"
    try:
        if kind == "pct":
            return f"{float(v):.1%}"
        if kind == "mod":
            return f"{float(v):.3f}"
        if kind == "idx":
            return f"{float(v):.4f}"
        if kind == "dollar":
            return f"{float(v):,.0f}"
        if kind == "num":
            return f"{float(v):g}"
    except (TypeError, ValueError):
        pass
    return str(v)


def _log_line(schema, r: dict) -> str:
    """One log row as a sentence fragment: 'ABD AZ 2026-03-31 +10.3% taken'."""
    parts = []
    for key, _title, kind in schema:
        v = r.get(key)
        if v is None or v == "":
            continue
        if kind == "pct":
            parts.append(f"{float(v):+.1%}")
        else:
            parts.append(str(v))
    return " ".join(parts)


def _freeze(r: dict) -> tuple:
    return tuple(sorted((k, str(v)) for k, v in r.items()
                        if v is not None and v != ""))


def _diff_log(line: str, label: str, schema, old: list, new: list,
              out: list) -> None:
    """Multiset diff — log rows have no unique key (two changes can share a
    month legitimately), so removed/added rows are rendered whole."""
    from collections import Counter
    co, cn = Counter(map(_freeze, old)), Counter(map(_freeze, new))
    by_freeze = {_freeze(r): r for r in old + new}
    for f in sorted((co - cn)):
        out.append(f"{line} · {label}: − {_log_line(schema, by_freeze[f])}")
    for f in sorted((cn - co)):
        out.append(f"{line} · {label}: + {_log_line(schema, by_freeze[f])}")


def diff_books(old_lines: dict, new_lines: dict) -> list:
    """Human sentences describing new vs old — the changelog's payload."""
    out: list = []
    for line in sorted(set(old_lines) - set(new_lines)):
        n = len(old_lines[line].get("lr_rows") or [])
        out.append(f"− {line}: line removed ({n} combos)")
    for line in sorted(set(new_lines) - set(old_lines)):
        n = len(new_lines[line].get("lr_rows") or [])
        out.append(f"+ {line}: line added ({n} combos)")
    for line in sorted(set(old_lines) & set(new_lines)):
        o, n = old_lines[line], new_lines[line]
        for key, label in _TOGGLES:
            if o.get(key) != n.get(key):
                out.append(f"{line}: {label} {o.get(key)} → {n.get(key)}")
        # tbl_LR: keyed by combo, field-level
        ok = {f"{r.get('bu')}|{r.get('state')}": r
              for r in o.get("lr_rows") or []}
        nk = {f"{r.get('bu')}|{r.get('state')}": r
              for r in n.get("lr_rows") or []}
        for key in sorted(set(ok) - set(nk)):
            out.append(f"{line}: − combo {key}")
        for key in sorted(set(nk) - set(ok)):
            out.append(f"{line}: + combo {key}")
        for key in sorted(set(ok) & set(nk)):
            ro, rn = ok[key], nk[key]
            for fkey, title, kind in LR_SCHEMA:
                if fkey in ("bu", "state"):
                    continue
                if ro.get(fkey) != rn.get(fkey):
                    out.append(f"{line} · {key}: {title} "
                               f"{_fmt(kind, ro.get(fkey))} → "
                               f"{_fmt(kind, rn.get(fkey))}")
        _diff_log(line, "Rate Log", RL_SCHEMA,
                  o.get("rate_rows") or [], n.get("rate_rows") or [], out)
        _diff_log(line, "Mod Log", ML_SCHEMA,
                  o.get("mod_rows") or [], n.get("mod_rows") or [], out)
        so = {r.get("state"): r.get("weights") for r in o.get("season_rows") or []}
        sn = {r.get("state"): r.get("weights") for r in n.get("season_rows") or []}
        for st in sorted(set(so) | set(sn)):
            if so.get(st) != sn.get(st):
                what = ("weights changed" if st in so and st in sn
                        else ("+ row" if st in sn else "− row"))
                out.append(f"{line} · Seasonality {st}: {what}")
    return out


# ------------------------------------------------------------- save / load

def _now() -> str:
    return dt.datetime.now().isoformat(timespec="seconds")


def save_book(path, book: dict, name: str, note: str = "") -> list:
    """Write the book; append a changelog entry diffed against the file
    being replaced. Returns the changes recorded."""
    path = Path(path)
    doc = book_to_doc(book, name)
    old = None
    if path.exists():
        old = yaml.safe_load(path.read_text(encoding="utf-8")) or {}
    if old is None:
        changes = [f"initial save — {len(doc['lines'])} line(s), "
                   f"{sum(len(d['lr_rows']) for d in doc['lines'].values())}"
                   f" combos"]
        log = []
    else:
        changes = diff_books(old.get("lines") or {}, doc["lines"])
        log = list(old.get("changelog") or [])
        if not changes:
            changes = ["no input changes"]
    log.append({"at": _now(), "note": note or "", "changes": changes})
    doc["changelog"] = log
    path.parent.mkdir(parents=True, exist_ok=True)
    path.write_text(yaml.safe_dump(doc, sort_keys=False, allow_unicode=True),
                    encoding="utf-8")
    return changes


def read_doc(path) -> dict:
    return yaml.safe_load(Path(path).read_text(encoding="utf-8")) or {}


def load_book(path) -> tuple:
    """-> ({line: Scenario}, doc). The doc rides along for the changelog."""
    path = Path(path)
    doc = read_doc(path)
    return doc_to_book(doc, source=str(path)), doc


def scenario_files() -> dict:
    """{filename: path} in scenarios/, newest first."""
    if not SCEN_DIR.exists():
        return {}
    found = sorted(SCEN_DIR.glob("*.y*ml"),
                   key=lambda p: p.stat().st_mtime, reverse=True)
    return {p.name: str(p) for p in found}


def safe_name(name: str) -> str:
    keep = "".join(c if c.isalnum() or c in "-_ " else "" for c in name)
    return ("-".join(keep.split()) or "scenario") + ".yaml"
