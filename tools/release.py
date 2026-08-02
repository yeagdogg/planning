"""One command for the whole pipeline: build, recalculate, roll up, verify (D87).

Before this, a full regeneration was nine manual invocations — build, six
recalcs (recalc.py takes one path), build_book, one more recalc — and both
builders ended by printing a note telling you to go run the next one yourself.
Two of those steps fail badly when skipped: a forgotten recalc at least stops
the harvest, but a forgotten book rebuild is silent, and the book then shows
last month's numbers under a fresh-looking as-of stamp.

The tiers exist because verification cost is wildly uneven:

    --smoke   pytest only                            seconds
    --quick   build + recalc + phases A-C            ~1 min per line
    --full    adds phase D (mutate -> recalc -> tie)  several min per line

Lines are verified as parallel subprocesses, each owning its own Excel instance
and its own scratch directory, so the wall clock is a few lines rather than six.
That only became possible once the harness stopped sharing a single scratch
file (D84) — before, two concurrent runs deleted each other's working copy.

Width is capped at three deliberately: at six, Excel/COM came apart — two lines
raised "The interface is unknown" and one produced a set of assertion failures
whose actual values were the workbook's DEFAULT state, meaning its mutation had
silently not been applied. Verification that can quietly fail to set up its own
exercise is worse than slow verification, so the cap stays low and the harness
now re-reads every mutated cell before asserting anything.

    python tools/release.py --quick
    python tools/release.py --full --lob Property
    python tools/release.py --smoke
"""

from __future__ import annotations

import argparse
import concurrent.futures as cf
import shutil
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # noqa: E402

from src.build_workbook import load_config, output_path  # noqa: E402
from tools.build_book import book_path  # noqa: E402

PY = sys.executable

# Concurrency is capped by MEMORY and by COM, and the two tiers are very
# different loads. Phase D reads a whole recalculated workbook into openpyxl per
# exercise — 865k cells, and that dominates the footprint — on top of an Excel
# instance per line: at six, COM came apart ('The interface is unknown', and one
# line whose mutation had silently not applied); at three, one line was killed
# outright with no output at all. Phases A-C carry no such load.
MAX_JOBS_FULL = 2
MAX_JOBS_QUICK = 4
ROOT = Path(__file__).resolve().parents[1]


def _run(args: list[str], label: str) -> tuple[str, int, str]:
    """Run a subprocess, returning (label, returncode, tail of output).

    A failure with NO output is a real outcome — a killed process, an
    out-of-memory, a COM teardown that took the interpreter with it — so say so
    rather than printing an empty string and leaving the reader to guess.
    """
    r = subprocess.run([PY, *args], cwd=ROOT, capture_output=True, text=True)
    out = (r.stdout or "") + (r.stderr or "")
    tail = "\n".join(l for l in out.splitlines()
                     if "RESULT" in l or "FAIL" in l or "Error" in l)
    if not tail:
        tail = out.strip()[-400:]
    if not tail:
        tail = (f"exit code {r.returncode} with no output — the process was very "
                f"likely killed (memory, or Excel taking the interpreter down). "
                f"Re-run this line on its own: python tools/verify_workbook.py "
                f"--lob \"{label}\" --skip-recalc")
    return label, r.returncode, tail


def _locked(out_dir: Path) -> list[str]:
    """Excel lock files. Recalculating a workbook that is open in Excel either
    fails or silently binds to the open copy, so refuse rather than guess."""
    return sorted(p.name for p in out_dir.glob("~$*.xlsx"))


def _sweep_stale_scratch(out_dir: Path, older_than_s: float = 7200) -> int:
    """Remove verify scratch directories left behind by a killed run.

    The harness removes its own in a finally block, but a process that is KILLED
    never gets there — and each directory holds a copy of the workbook per
    distinct mutation state, so an abandoned one is ~100MB. Only sweep clearly
    stale ones, so a verify running in another window is never touched.
    """
    import time as _t
    n = 0
    for d in out_dir.glob("verify_*"):
        try:
            if d.is_dir() and _t.time() - d.stat().st_mtime > older_than_s:
                shutil.rmtree(d, ignore_errors=True)
                n += 1
        except OSError:
            pass
    return n


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Build, recalculate and verify the whole artifact set.")
    ap.add_argument("--config", default="config/config.yaml")
    ap.add_argument("--out", default=None)
    ap.add_argument("--lob", action="append", default=None,
                    help="limit to one line (repeatable); default all configured")
    tier = ap.add_mutually_exclusive_group()
    tier.add_argument("--smoke", action="store_true", help="pytest only")
    tier.add_argument("--quick", action="store_true",
                      help="build + recalc + verify phases A-C (default)")
    tier.add_argument("--full", action="store_true",
                      help="everything, including phase D exercises")
    ap.add_argument("--jobs", type=int, default=0,
                    help="parallel workers (default 2 for --full, 4 for --quick; "
                         "phase D is memory-heavy and Excel/COM destabilises "
                         "when over-committed)")
    ap.add_argument("--skip-build", action="store_true",
                    help="verify what is already in the output directory")
    ap.add_argument("--carry-forward", action="store_true",
                    help="rebuild around each workbook's existing inputs (D82)")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    out_dir = Path(args.out or cfg.output_dir)
    lobs = args.lob or list(cfg.output_lobs)
    for name in lobs:
        if name not in {l.name for l in cfg.lobs}:
            raise SystemExit(f"unknown LOB {name!r}")
    t_start = time.perf_counter()
    failures: list[str] = []

    def stage(title: str):
        print(f"\n=== {title} " + "=" * max(0, 56 - len(title)))

    stage("pytest")
    label, rc, tail = _run(["-m", "pytest", "tests", "-q"], "pytest")
    print(tail.splitlines()[-1] if tail else "(no output)")
    if rc:
        print(tail)
        return 1
    if args.smoke:
        print(f"\nsmoke OK in {time.perf_counter() - t_start:.0f}s")
        return 0

    if not args.skip_build:
        locked = _locked(out_dir)
        if locked:
            raise SystemExit(
                f"REFUSING to build: {', '.join(locked)} indicate the workbook(s) are "
                f"open in Excel. Close them and retry.")
        swept = _sweep_stale_scratch(out_dir)
        if swept:
            print(f"  (swept {swept} abandoned scratch director"
                  f"{'y' if swept == 1 else 'ies'} from a killed run)")
        stage("build")
        build_args = ["-m", "src.build_workbook", "--config", args.config,
                      "--out", str(out_dir)]
        if args.carry_forward:
            build_args.append("--carry-forward")
        for name in lobs:
            label, rc, tail = _run(build_args + ["--lob", name], name)
            print(f"  {'ok  ' if not rc else 'FAIL'} {name}")
            if rc:
                failures.append(f"build {name}: {tail}")
        if failures:
            print("\n".join(failures))
            return 1

        stage("recalculate")
        n = args.jobs or min(len(lobs), MAX_JOBS_FULL if args.full else MAX_JOBS_QUICK)
        with cf.ThreadPoolExecutor(max_workers=n) as pool:
            futs = {pool.submit(_run, ["tools/recalc.py", str(output_path(cfg, out_dir, name))],
                                name): name for name in lobs}
            for f in cf.as_completed(futs):
                label, rc, tail = f.result()
                print(f"  {'ok  ' if not rc else 'FAIL'} {label}")
                if rc:
                    failures.append(f"recalc {label}: {tail}")
        if failures:
            print("\n".join(failures))
            return 1

        # The book must follow every line, or it silently reports last month.
        if set(lobs) == set(cfg.output_lobs):
            stage("book")
            for step in (["tools/build_book.py", "--config", args.config,
                          "--out", str(out_dir)],
                         ["tools/recalc.py", str(book_path(cfg, out_dir))]):
                label, rc, tail = _run(step, "book")
                if rc:
                    failures.append(f"book: {tail}")
            print("  ok   book" if not failures else "  FAIL book")
            if failures:
                print("\n".join(failures))
                return 1
        else:
            print("\n(skipping the book: it needs every line, and this run is "
                  f"limited to {', '.join(lobs)})")

    stage("verify" + (" (full, phase D)" if args.full else " (phases A-C)"))
    verify = ["tools/verify_workbook.py", "--config", args.config, "--skip-recalc"]
    if not args.full:
        verify.append("--quick")
    n = args.jobs or min(len(lobs), MAX_JOBS_FULL if args.full else MAX_JOBS_QUICK)
    with cf.ThreadPoolExecutor(max_workers=n) as pool:
        futs = {pool.submit(_run, verify + ["--lob", name], name): name for name in lobs}
        for f in cf.as_completed(futs):
            label, rc, tail = f.result()
            print(f"  {'ok  ' if not rc else 'FAIL'} {label:<20} {tail.splitlines()[-1] if tail else ''}")
            if rc:
                failures.append(f"verify {label}: {tail}")
    if set(lobs) == set(cfg.output_lobs):
        bargs = ["tools/verify_book.py", "--config", args.config, "--skip-recalc"]
        if not args.full:
            bargs.append("--quick")
        label, rc, tail = _run(bargs, "book")
        print(f"  {'ok  ' if not rc else 'FAIL'} {'book':<20} "
              f"{tail.splitlines()[-1] if tail else ''}")
        if rc:
            failures.append(f"verify book: {tail}")

    dt = time.perf_counter() - t_start
    print(f"\n{'=' * 60}")
    if failures:
        print(f"RELEASE FAILED in {dt:.0f}s")
        for f in failures:
            print(f"  {f}")
        return 1
    print(f"RELEASE OK in {dt:.0f}s  ({len(lobs)} lines"
          f"{' + book' if set(lobs) == set(cfg.output_lobs) else ''})")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
