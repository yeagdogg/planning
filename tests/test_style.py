"""text_height calibration: prose rows must never clip (the pre-2.1 bug sized
rows for ~110 chars/line regardless of column width, hiding 60-75% of the
Methodology / Read Me text).

Run:  python -m pytest tests -q
"""

from __future__ import annotations

from src.xlstyle import text_height

# the worst shipped paragraphs (measured from the 2.0 workbooks)
METHODOLOGY_S1 = 626   # Methodology!B5, column width 30, 10pt
README_PURPOSE = 480   # Read Me!B6,    column width 34, 10pt


def _height_for(n_chars: int, width: float, size: int = 10) -> float:
    return text_height("x" * n_chars, width, size=size)


def test_worst_known_cells_get_enough_height():
    # empirical need: ~281pt for the 626-char cell, ~204pt for the 480-char one
    assert _height_for(METHODOLOGY_S1, 30) >= 281
    assert _height_for(README_PURPOSE, 34) >= 204


def test_height_is_monotone_in_text_and_inverse_in_width():
    assert _height_for(200, 30) <= _height_for(400, 30)
    assert _height_for(400, 90) <= _height_for(400, 30)


def test_bounds():
    assert text_height("short", 30) >= 15
    # Excel's hard ceiling is 409.5pt; we stay under it
    assert _height_for(20000, 5) <= 405.5
    assert text_height("", 30) >= 15


def test_newlines_count_as_line_breaks():
    two_paras = "a" * 30 + "\n" + "a" * 30
    assert text_height(two_paras, 90) > text_height("a" * 60, 90) - 1e-9


def test_smaller_font_fits_more_per_line():
    assert _height_for(400, 30, size=9) <= _height_for(400, 30, size=11)
