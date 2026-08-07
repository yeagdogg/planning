"""D113 A/B: one LOB, both dialects, identical inputs — prove the values agree.

Builds the LOB twice from the SAME carried inputs (classic and modern), runs
both through Excel, then compares every populated cell value on every sheet
except the two that intentionally differ:

  * Read Me — carries the dialect stamp, the degradation banner and the LET
    canary (modern only);
  * Checks  — the modern build has one extra row (the dialect sentinel), which
    shifts the rows below it; the harness verifies that sheet separately in
    both files.

Everything else — every exhibit, both logs, both engines, every hidden `_calc`
and `_netcalc` cell — must match exactly. This is the value half of the
dual-dialect guarantee; the structural half is tests/test_letfn.py.

Usage:  python tools/ab_compare.py [--lob Property]
Exit 0 only when zero differences survive.
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))

import openpyxl  # noqa: E402

from src.build_workbook import build, load_config, output_path  # noqa: E402
from src.carry import read_inputs  # noqa: E402
from tools.recalc import recalc  # noqa: E402

SKIP_SHEETS = {"Read Me", "Checks"}


def _values(path: Path) -> dict:
    wb = openpyxl.load_workbook(path, data_only=True)
    out = {}
    for ws in wb.worksheets:
        if ws.title in SKIP_SHEETS:
            continue
        for row in ws.iter_rows():
            for c in row:
                if c.value is not None:
                    out[f"{ws.title}!{c.coordinate}"] = c.value
    return out


def _same(a, b) -> bool:
    if isinstance(a, (int, float)) and isinstance(b, (int, float)) \
            and not isinstance(a, bool) and not isinstance(b, bool):
        return abs(a - b) <= 1e-12 * max(1.0, abs(a), abs(b))
    return a == b


def main(argv=None) -> int:
    ap = argparse.ArgumentParser()
    ap.add_argument("--lob", default="Property")
    ap.add_argument("--config", default="config/config.yaml")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    out_dir = Path(cfg.output_dir)
    shipped = output_path(cfg, out_dir, args.lob)
    carried = read_inputs(shipped) if shipped.exists() else None
    src = shipped.name if carried else "sample seeds"
    print(f"A/B inputs: {src}")

    paths = {}
    for mode, tag in (("classic", "ABBASE"), ("modern", "MODERN")):
        p = shipped.with_name(f"{shipped.stem}_{tag}{shipped.suffix}")
        wb = build(cfg, args.lob, carried=carried, formula_mode=mode)
        wb.save(p)
        used = recalc(p)
        print(f"  built + recalculated ({used}): {p.name}")
        paths[mode] = p

    va, vb = _values(paths["classic"]), _values(paths["modern"])
    only_a = sorted(set(va) - set(vb))
    only_b = sorted(set(vb) - set(va))
    diffs = [k for k in va.keys() & vb.keys() if not _same(va[k], vb[k])]
    print(f"compared {len(va.keys() & vb.keys()):,} populated cells "
          f"across {len(set(k.split('!')[0] for k in va)):,} sheets "
          f"(Read Me and Checks excluded by design)")
    for label, bad in (("value differs", diffs),
                       ("only in classic", only_a), ("only in modern", only_b)):
        for k in bad[:8]:
            print(f"  DIFF [{label}] {k}: "
                  f"classic={va.get(k)!r} modern={vb.get(k)!r}")
        if len(bad) > 8:
            print(f"  ... and {len(bad) - 8} more [{label}]")
    n_bad = len(diffs) + len(only_a) + len(only_b)
    print(f"RESULT: {'IDENTICAL' if n_bad == 0 else f'{n_bad} differences'}")
    return 1 if n_bad else 0


if __name__ == "__main__":
    raise SystemExit(main())
