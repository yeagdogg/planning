"""Build the combined book workbook from the generated LOB files (D66).

    python tools/build_book.py                    # config/config.yaml -> output/
    python tools/build_book.py --out somewhere/

Requires the LOB workbooks to have been generated AND recalculated — the
harvest refuses otherwise, so a stale book cannot be produced by accident.
The book itself is written without cached values like any openpyxl output;
run tools/recalc.py on it before reading values back.
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # noqa: E402

import argparse

from src.build_workbook import load_config
from src.sheets_book import build_book
from tools.harvest import HarvestError, harvest


def book_path(cfg, out_dir: Path) -> Path:
    return out_dir / f"Plan_LR_Book_{cfg.plan_year}.xlsx"


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(description="Build the combined book workbook.")
    ap.add_argument("--config", default="config/config.yaml")
    ap.add_argument("--out", default=None, help="output directory (default: from config)")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    out_dir = Path(args.out or cfg.output_dir)
    try:
        data = harvest(cfg, out_dir)
    except HarvestError as e:
        print(f"cannot build the book: {e}")
        return 2
    wb = build_book(data)
    out_dir.mkdir(parents=True, exist_ok=True)
    path = book_path(cfg, out_dir)
    wb.save(path)
    print(f"wrote {path}  ({len(data.rows)} combos from {len(data.sources)} workbooks, "
          f"as of {data.as_of})")
    print("NOTE: run tools/recalc.py on it so the filters and totals carry values.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
