"""Generate master-paste blocks from the current fleet workbooks (P6).

    python -m app.make_masters                # -> output/masters/
    python -m app.make_masters --out somewhere

One .txt per master (tbl_LR, Rate Log, Mod Log): tab-separated, title
header included, one leading Line column — exactly what the Book page's
master boxes accept. Open a file, Ctrl+A, Ctrl+C, paste. The blocks are
built by ``masters.master_text`` from ``importers.fleet_choices()`` (one
workbook per configured line), so they are the fleet's OWN inputs in
master dress — the round-trip the tests pin.

Note what a master deliberately does NOT carry: per-line toggles (term,
trend default, seasonality on/off, mod master) and the seasonality
table. Those live on the Inputs page and in scenario files.
"""
from __future__ import annotations

import argparse
from pathlib import Path

from . import importers
from .masters import MASTERS, master_text

ROOT = Path(__file__).resolve().parents[1]

_FILENAMES = {"tbl_LR": "master_tbl_LR.txt",
              "Rate Log": "master_rate_log.txt",
              "Mod Log": "master_mod_log.txt"}


def make_masters(out_dir=None) -> dict:
    """Write the three master blocks; -> {table: path}."""
    out = Path(out_dir) if out_dir else ROOT / "output" / "masters"
    out.mkdir(parents=True, exist_ok=True)
    book = {}
    for lob, path in importers.fleet_choices().items():
        book[lob] = importers.from_workbook(path)
    written = {}
    for table in MASTERS:
        p = out / _FILENAMES[table]
        p.write_text(master_text(book, table) + "\n", encoding="utf-8")
        written[table] = p
    return written


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Write master-paste blocks from the fleet workbooks.")
    ap.add_argument("--out", default=None,
                    help="output directory (default: output/masters/)")
    args = ap.parse_args(argv)
    for table, p in make_masters(args.out).items():
        n_rows = sum(1 for _ in p.open(encoding="utf-8")) - 1
        print(f"{table}: {p}  ({n_rows} rows)")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
