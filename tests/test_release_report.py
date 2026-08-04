"""What the release driver tells you when a step fails (D105).

This is not cosmetics. Three separate times a real failure has been reported as
something else, and each time the misdirection cost more than the bug:

  * D97/D100 — four wrong theories and three lost runs, because the driver kept
    every line containing "Error" and the LAST one was always
    `LibreOffice (soffice) not found`: the FALLBACK engine complaining about an
    engine that was never going to run on this machine. The real fault, a
    `pywintypes.com_error` meaning "Excel is busy", sat two frames up.
  * D105 — a green `346 passed` run reported as
    `<cannot get C stack on this system>`, because stdout and stderr were
    concatenated and the caller prints the last line.

So the extractor is pinned against the actual output of the things it reads.
"""

from __future__ import annotations

from tools.release import _verdict

PYTEST_PASS = "........ [ 41%]\n346 passed in 105.35s (0:01:45)\n"

# faulthandler's dump on a first-chance RPC_E_CALL_REJECTED, which busy_retry
# handles. It goes to stderr, after the summary, and is not a failure.
FAULTHANDLER_NOISE = (
    "Windows fatal exception: code 0x80010001\n"
    "\n"
    "Current thread 0x0000d3a8 (most recent call first):\n"
    '  File "dynamic.py", line 497 in _LazyAddAttr_\n'
    "Current thread's C stack trace (most recent call first):\n"
    "  <cannot get C stack on this system>\n")


def test_a_passing_run_reports_that_it_passed():
    """The regression: stderr noise arriving after stdout's verdict must not
    become the verdict, because the caller prints only the last line."""
    got = _verdict(PYTEST_PASS, FAULTHANDLER_NOISE)
    assert got.splitlines()[-1] == "346 passed in 105.35s (0:01:45)"
    assert "C stack" not in got


def test_the_harness_result_line_wins():
    assert _verdict("Phase D: toggle exercises\nRESULT: 317 passed, 0 failed\n",
                    "") == "RESULT: 317 passed, 0 failed"


def test_failing_assertions_are_not_dropped():
    """The old filter caught these with a plain `"FAIL" in l`; the harness
    indents them, so a `^FAIL` anchor would silently lose every one."""
    out = ("RESULT: 315 passed, 2 failed\n"
           "  FAIL: [mod master OFF] A_mod = 1\n"
           "  FAIL: [basis proposed] CY LR ties oracle\n")
    got = _verdict(out, "")
    assert "RESULT: 315 passed, 2 failed" in got
    assert got.count("FAIL:") == 2


def test_the_real_com_error_survives_alongside_the_red_herring():
    """D100's exact output. The LibreOffice line is kept — on a machine with no
    Excel it IS the failure — but it must no longer be the only thing shown."""
    err = ("Traceback (most recent call last):\n"
           '  File "recalc.py", line 9, in recalc_with_excel\n'
           "pywintypes.com_error: (-2147418111, 'Call was rejected by callee.')\n"
           "Error: LibreOffice (soffice) not found\n")
    got = _verdict("", err)
    assert "com_error" in got, got
    assert "-2147418111" in got


def test_the_build_refusal_is_surfaced():
    """Refusing to build over a workbook someone has open is a verdict, not an
    exception to bury."""
    err = ("REFUSING to build: ~$Plan_LR_Workbook_2027_Commercial_Auto.xlsx "
           "indicate the workbook(s) are open in Excel. Close them and retry.\n")
    assert "REFUSING to build" in _verdict("", err)


def test_stdout_is_preferred_over_stderr():
    """A tool's verdict goes to stdout; stderr carries warnings. When both have
    something that matches, the verdict is the one that means something."""
    got = _verdict("RESULT: 317 passed, 0 failed\n",
                   "DeprecationWarning: something Error: unrelated\n")
    assert got == "RESULT: 317 passed, 0 failed"


def test_unrecognisable_output_still_says_something_useful():
    """No pattern matched: show real content, not whatever was last."""
    got = _verdict("", "the process died\n" + FAULTHANDLER_NOISE)
    assert "the process died" in got
    assert "C stack" not in got


def test_no_output_at_all_returns_empty_so_the_caller_can_say_so():
    assert _verdict("", "") == ""
    assert _verdict("", FAULTHANDLER_NOISE) == ""
