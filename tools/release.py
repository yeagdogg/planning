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

Width is FITTED to the machine (D94), not hard-coded: a phase-D worker is one
Python process plus one Excel instance and the pair peaks around 670 MB, so the
bound is `min(lines, cores - 2, two-thirds of free RAM / 900 MB, cap)`, cap 4 for
--full and 8 for --quick. It used to be 2, which on a 24-core box turned six
lines into three waves for no reason; six turned out too wide in the other
direction, because Excel/COM starts killing workers and each kill costs a full
serial re-run.

And phase D does not run when it cannot say anything. It mutates inputs,
recalculates through Excel and ties the results to the oracle — it tests the
ARITHMETIC — so if a line rebuilt to byte-identical formulas and names since the
last green full run, re-deriving the same answers is ceremony. Those lines fall
back to phases A-C and the run says which and why. `--force-full` overrides.

    python tools/release.py --quick
    python tools/release.py --full --lob Property
    python tools/release.py --smoke
"""

from __future__ import annotations

import argparse
import concurrent.futures as cf
import hashlib
import json
import os
import shutil
import subprocess
import sys
import time
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parents[1]))  # noqa: E402

from src.build_workbook import load_config, output_path  # noqa: E402
from tools.build_book import book_path  # noqa: E402

PY = sys.executable

ROOT = Path(__file__).resolve().parents[1]
STATE_FILE = "release-state.json"

# Measured, not guessed (D94): a phase-D worker is one Python process plus one
# Excel instance, and the pair peaks around 670 MB — Excel ~350, openpyxl ~320.
# Phases A-C never start Excel at all. The old hard-coded cap of 2 came from a
# run where a third line died, which was later identified as an ACCESS VIOLATION
# rather than memory pressure; a killed line is now retried serially, so the cap
# no longer has to absorb that failure by staying small.
MB_PER_JOB_FULL = 900
MB_PER_JOB_QUICK = 400
# The binding constraint on --full is COM STABILITY, not memory. Measured over
# three runs at six concurrent Excel instances (39 GB free, 24 cores, nowhere
# near a memory wall): 0, 1 and 2 lines killed by access violation. Each kill
# costs a full serial re-run of that line — about 7.5 minutes — which more than
# eats the wave that the sixth worker saved. Two kills took a 26-minute run to
# 47. Four is the honest setting: two waves instead of one, but a fan-out that
# usually finishes on the first pass. Phases A-C never start Excel, so they keep
# the wider cap.
HARD_JOB_CAP_FULL = 4
HARD_JOB_CAP_QUICK = 8


def _free_mb() -> int:
    """Physical memory currently available, or -1 where we cannot tell."""
    try:
        import ctypes

        class _MS(ctypes.Structure):
            _fields_ = [("dwLength", ctypes.c_ulong),
                        ("dwMemoryLoad", ctypes.c_ulong),
                        ("ullTotalPhys", ctypes.c_ulonglong),
                        ("ullAvailPhys", ctypes.c_ulonglong),
                        ("ullTotalPageFile", ctypes.c_ulonglong),
                        ("ullAvailPageFile", ctypes.c_ulonglong),
                        ("ullTotalVirtual", ctypes.c_ulonglong),
                        ("ullAvailVirtual", ctypes.c_ulonglong),
                        ("ullAvailExtendedVirtual", ctypes.c_ulonglong)]
        st = _MS()
        st.dwLength = ctypes.sizeof(_MS)
        if ctypes.windll.kernel32.GlobalMemoryStatusEx(ctypes.byref(st)):
            return int(st.ullAvailPhys // (1024 * 1024))
    except Exception:  # noqa: BLE001 - not Windows, or the call is unavailable
        pass
    return -1


def _max_jobs(full: bool, n_lobs: int) -> int:
    """Fit the fan-out to the machine rather than to the machine I had.

    Half the wall clock of a --full run used to be waves that existed only
    because the cap was two: six lines, two at a time, on a box with 24 cores
    and 60 GB. The bound that matters is memory per worker, and it is small.
    """
    per = MB_PER_JOB_FULL if full else MB_PER_JOB_QUICK
    cap = HARD_JOB_CAP_FULL if full else HARD_JOB_CAP_QUICK
    by_cpu = max(1, (os.cpu_count() or 4) - 2)
    free = _free_mb()
    # leave a third of free memory alone: Excel grows while it recalculates,
    # and a swapping box is slower than a serial one
    by_mem = max(1, int(free * 0.67) // per) if free > 0 else cap
    return max(1, min(n_lobs, by_cpu, by_mem, cap))


def _formula_fingerprint(path: Path) -> str:
    """A hash of every FORMULA in a workbook, plus its defined names.

    This is what decides whether phase D has anything to say (D94). Phase D
    mutates inputs, recalculates, and ties the results to the oracle — so it
    tests the ARITHMETIC. If a rebuild produces byte-identical formulas and
    byte-identical name resolution, the arithmetic cannot have moved, and
    re-running forty minutes of Excel to re-derive the same answers is not
    verification, it is ceremony.

    Deliberately NOT a hash of the file: as-of stamps and cached values change
    on every build, so the file always differs while the model usually does not.
    """
    import openpyxl

    h = hashlib.sha256()
    wb = openpyxl.load_workbook(path, data_only=False)
    try:
        for ws in wb.worksheets:
            h.update(f"\x00SHEET:{ws.title}".encode())
            for row in ws.iter_rows():
                for c in row:
                    v = c.value
                    if isinstance(v, str) and v.startswith("="):
                        h.update(f"\x01{c.coordinate}={v}".encode())
        for name in sorted(wb.defined_names):
            h.update(f"\x02{name}={wb.defined_names[name].value}".encode())
    finally:
        wb.close()
    return h.hexdigest()


def _load_state(out_dir: Path) -> dict:
    try:
        return json.loads((out_dir / STATE_FILE).read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def _save_state(out_dir: Path, state: dict) -> None:
    try:
        (out_dir / STATE_FILE).write_text(json.dumps(state, indent=1),
                                          encoding="utf-8")
    except OSError:
        pass


def _run(args: list[str], label: str, retries: int = 0) -> tuple[str, int, str]:
    """Run a subprocess, returning (label, returncode, tail of output).

    A failure with NO output is a real outcome — a killed process, an
    out-of-memory, a COM teardown that took the interpreter with it — so say so
    rather than printing an empty string and leaving the reader to guess.

    ``retries`` re-runs in a FRESH process on failure (D92). In-process retries
    already exist in recalc.py and are not always enough: the step that follows
    a fan-out has hit three different COM failures — 'The interface is unknown',
    an RPC error, and a proxy answering a property read with the wrong type —
    and every one of them cleared on a new process while none of them cleared
    on a new Excel instance inside the same one. Whatever state goes bad lives
    in the interpreter's COM apartment, not in Excel.
    """
    for attempt in range(retries + 1):
        r = subprocess.run([PY, *args], cwd=ROOT, capture_output=True, text=True)
        if r.returncode == 0 or attempt == retries:
            break
        _await_excel_drain()
        print(f"  ..   {label} failed; retrying in a fresh process "
              f"({attempt + 1}/{retries})")
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


def _excel_pids() -> set[int]:
    """PIDs of every running Excel, or an empty set where we cannot tell."""
    try:
        out = subprocess.run(
            ["tasklist", "/FI", "IMAGENAME eq EXCEL.EXE", "/FO", "CSV", "/NH"],
            capture_output=True, text=True, timeout=20).stdout
    except (OSError, subprocess.SubprocessError):
        return set()
    pids = set()
    for line in out.splitlines():
        parts = [p.strip('" ') for p in line.split('","')]
        if len(parts) > 1 and parts[0].upper().startswith("EXCEL"):
            try:
                pids.add(int(parts[1]))
            except ValueError:
                pass
    return pids


_EXCEL_BASELINE: set[int] = set()


def _await_excel_drain(timeout_s: float = 45.0) -> int:
    """Wait for the instances THIS RUN started to die, then kill the stragglers.

    Excel's teardown is asynchronous: the subprocess returns as soon as it has
    called Quit, while the process lingers releasing its RPC endpoints. Starting
    the next instance inside that window fails a different way each time — 'The
    interface is unknown', an RPC error, or a proxy answering a property read
    with the wrong type ("'bool' object is not callable") — which is why the
    symptom set looked like several unrelated bugs (D92).

    Waiting alone was not enough: one instance sometimes never exits at all, and
    the driver would print "still shutting down; continuing anyway" and then
    fail all three book attempts against the wedged process. So after the wait,
    stragglers are terminated.

    ONLY ones that appeared after this run started. `_EXCEL_BASELINE` is
    snapshotted before any subprocess launches, so an Excel the user already had
    open — with unsaved work in it — is never a candidate. (The release also
    refuses to start at all when a `~$` lock file says one of OUR workbooks is
    open, which is the other half of that guarantee.)
    """
    deadline = time.monotonic() + timeout_s
    ours: set[int] = set()
    while time.monotonic() < deadline:
        ours = _excel_pids() - _EXCEL_BASELINE
        if not ours:
            return 0
        time.sleep(1.0)
    for pid in ours:
        try:
            subprocess.run(["taskkill", "/PID", str(pid), "/F"],
                           capture_output=True, timeout=20)
        except (OSError, subprocess.SubprocessError):
            pass
    if ours:
        print(f"  (terminated {len(ours)} wedged Excel process(es) this run started)")
        time.sleep(2.0)
    return len(ours)


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
                    help="parallel workers (default: fitted to free memory and CPU "
                         "count — about 900MB per --full worker)")
    ap.add_argument("--skip-build", action="store_true",
                    help="verify what is already in the output directory")
    ap.add_argument("--carry-forward", action="store_true",
                    help="rebuild around each workbook's existing inputs (D82)")
    ap.add_argument("--force-full", action="store_true",
                    help="run phase D even where every formula is unchanged")
    args = ap.parse_args(argv)

    global _EXCEL_BASELINE
    _EXCEL_BASELINE = _excel_pids()      # never kill an Excel the user already had open

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
        n = args.jobs or _max_jobs(args.full, len(lobs))
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
            _await_excel_drain()
            stage("book")
            for step in (["tools/build_book.py", "--config", args.config,
                          "--out", str(out_dir)],
                         ["tools/recalc.py", str(book_path(cfg, out_dir))]):
                label, rc, tail = _run(step, "book", retries=2)
                if rc:
                    failures.append(f"book: {tail}")
            print("  ok   book" if not failures else "  FAIL book")
            if failures:
                print("\n".join(failures))
                return 1
        else:
            print("\n(skipping the book: it needs every line, and this run is "
                  f"limited to {', '.join(lobs)})")

    # ---- does phase D have anything to say? (D94) --------------------------
    # Phase D mutates inputs, recalculates through Excel, and ties the results
    # to the oracle: it tests the ARITHMETIC. If a line rebuilt to byte-identical
    # formulas and byte-identical names, the arithmetic did not move, and forty
    # minutes of Excel would re-derive answers already derived. Skip it, and SAY
    # SO — a silent skip is how a verification tier quietly stops existing.
    state = _load_state(out_dir)
    prints = state.get("phase_d_ok", {})
    fps: dict[str, str] = {}
    skip_d: set[str] = set()
    if args.full and not args.force_full:
        for name in lobs:
            try:
                fps[name] = _formula_fingerprint(output_path(cfg, out_dir, name))
            except Exception:  # noqa: BLE001 - never let bookkeeping fail a run
                continue
            if prints.get(name) == fps[name]:
                skip_d.add(name)
        if skip_d:
            print(f"\n(phase D skipped for {', '.join(sorted(skip_d))}: every formula "
                  f"and defined name is unchanged since the last green full run, so "
                  f"the oracle ties cannot move. --force-full overrides.)")

    stage("verify" + (" (full, phase D)" if args.full else " (phases A-C)"))
    base = ["tools/verify_workbook.py", "--config", args.config, "--skip-recalc"]

    def _argv(name):
        deep = args.full and name not in skip_d
        return base + ([] if deep else ["--quick"]) + ["--lob", name]

    # THE BOOK GOES FIRST (D94). It depends on the lines being built and
    # recalculated — which happened above — not on their verification, so there
    # is no reason for it to run in the wake of a six-way fan-out. Every COM
    # failure this pipeline has produced has been at whichever step followed
    # one, and moving the book out of that window is a better fix than retrying
    # inside it: the last run burned 270 seconds failing three times and still
    # went red on a book that passes on its own.
    if set(lobs) == set(cfg.output_lobs):
        bargs = ["tools/verify_book.py", "--config", args.config, "--skip-recalc"]
        if not args.full:
            bargs.append("--quick")
        label, rc, tail = _run(bargs, "book", retries=2)
        print(f"  {'ok  ' if not rc else 'FAIL'} {'book':<20} "
              f"{tail.splitlines()[-1] if tail else ''}")
        if rc:
            failures.append(f"verify book: {tail}")
        _await_excel_drain()

    n = args.jobs or _max_jobs(args.full and len(skip_d) < len(lobs), len(lobs))
    print(f"  ({n} parallel worker{'s' if n != 1 else ''})")
    with cf.ThreadPoolExecutor(max_workers=n) as pool:
        futs = {pool.submit(_run, _argv(name), name): name for name in lobs}
        for f in cf.as_completed(futs):
            label, rc, tail = f.result()
            print(f"  {'ok  ' if not rc else 'FAIL'} {label:<20} {tail.splitlines()[-1] if tail else ''}")
            if rc:
                failures.append(f"verify {label}: {tail}")
    # A line that was KILLED (access violation, out of memory) rather than
    # failing an assertion is an environment outcome, not a verification
    # outcome, and it always passes on its own. Retry those serially rather
    # than reporting a red release for a resource problem — but only those:
    # a line that actually FAILED an assertion has produced output, and
    # re-running it would just fail again more slowly.
    killed = [n_ for n_, (rc_, tail_) in
              ((k, (r_[1], r_[2])) for k, r_ in
               ((futs[f], f.result()) for f in futs))
              if rc_ and "no output" in tail_]
    for name in killed:
        failures = [x for x in failures if not x.startswith(f"verify {name}:")]
        _await_excel_drain()
        print(f"  ..   {name} was killed with no output; retrying on its own")
        label, rc, tail = _run(_argv(name), name)
        print(f"  {'ok  ' if not rc else 'FAIL'} {label:<20} "
              f"{tail.splitlines()[-1] if tail else ''}")
        if rc:
            failures.append(f"verify {label}: {tail}")

    # Record what phase D actually proved, so the next run can skip what it
    # would only re-prove. Only lines that RAN it and passed: a skipped line
    # keeps the print it already had, and any failure clears every print so the
    # next run cannot inherit a green stamp it did not earn.
    if args.full:
        if failures:
            for name in lobs:
                prints.pop(name, None)
        else:
            for name in lobs:
                if name in fps and name not in skip_d:
                    prints[name] = fps[name]
        state["phase_d_ok"] = prints
        _save_state(out_dir, state)

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
