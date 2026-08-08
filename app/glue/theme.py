"""Shared visual language — lifted from wfmodelingworkbench (mofo/theme.py),
rebranded to the workbook's own palette so the app and the Excel artifacts
read as one product: navy headers, steel accents (src/xlstyle.py).
"""
from __future__ import annotations

import os

# The workbook palette (src/xlstyle.py): NAVY headers, STEEL accents.
NAVY = "#1F3864"
STEEL = "#8497B0"
STEEL_LIGHT = "#D6DCE4"
PASS_GREEN = "#548235"
FAIL_RED = "#C00000"
WARN_AMBER = "#BF8F00"
ACCENT = NAVY

# The grey ramp (W3g): eleven ad-hoc greys collapsed to four steps. TINTS
# (banding #F5F7FA, lever #EAF2FC, the LR scale pastels, STEEL_LIGHT) are
# deliberately not greys and live where they are used.
INK = "#333333"          # chart ink, table text
GREY_TEXT = "#595959"    # subtitles, echoes, notes, secondary labels
GREY_LINE = "#D9D9D9"    # borders, rules, gridlines
GREY_FILL = "#F2F2F2"    # banner and panel fills

# Every Tabulator attaches this (W3a). The Fast design silently forces the
# 'fast' Tabulator theme, whose stylesheet repaints hover/selected ROWS with
# the accent background and WHITE text — over our per-CELL inline tints
# (lever blue, banding, the LR scales) that makes cells unreadable, and the
# hover rule fires on every table whether or not selection is enabled.
# Panel appends the theme stylesheet AFTER user stylesheets, so these rules
# carry !important and target the CELL (the theme sets color on the row).
# Hover becomes a faint navy overlay that rides ON TOP of the cell tints;
# selection (kept on the input grids for Ctrl+C copy-out) becomes
# STEEL_LIGHT with dark text.
TABLE_CSS = f"""
.tabulator-row.tabulator-selectable:hover {{
  background: transparent !important;
  color: inherit !important;
}}
.tabulator-row.tabulator-selectable:hover .tabulator-cell {{
  color: var(--neutral-foreground-rest, #2b2b2b) !important;
  box-shadow: inset 0 0 0 100vmax rgba(31, 56, 100, 0.05);
}}
.tabulator-row.tabulator-selected {{
  background: {STEEL_LIGHT} !important;
  color: inherit !important;
}}
.tabulator-row.tabulator-selected .tabulator-cell {{
  color: var(--neutral-foreground-rest, #2b2b2b) !important;
}}
.tabulator .tabulator-cell {{ font-variant-numeric: tabular-nums; }}
"""

# ============================ BRANDING ====================================
# The ONLY three things to edit to rebrand. Defaults reproduce the workbook
# look. Chart/table colors and the accent are intentionally NOT here — this
# is header + sidebar + logo only.

LOGO = ""            # path to a PNG/SVG shown in the header; "" = none
HEADER_BG = ACCENT   # header bar background
HEADER_TEXT = ""     # "" keeps the template default (white)
SIDEBAR_BG = ""      # "" keeps the template default

_ROOT = os.path.dirname(os.path.dirname(os.path.dirname(os.path.abspath(__file__))))


def _logo_path() -> str:
    if not LOGO:
        return ""
    return LOGO if os.path.isabs(LOGO) else os.path.join(_ROOT, LOGO)


def brand_kwargs() -> dict:
    """FastListTemplate kwargs for the current brand — pass with ``**``."""
    kw = {"header_background": HEADER_BG}
    logo = _logo_path()
    if logo:
        kw["logo"] = logo
    if HEADER_TEXT:
        kw["header_color"] = HEADER_TEXT
    return kw


def brand_raw_css() -> list:
    """Raw CSS for brand bits the template has no parameter for. The
    sidebar is ``<div id="sidebar">`` in the Fast template's document."""
    if not SIDEBAR_BG:
        return []
    return [f"#sidebar {{ background: {SIDEBAR_BG} !important; }}"]
# ========================== END BRANDING ==================================

# --------------------------------------------------- theme-aware charting
def is_dark() -> bool:
    """The FastListTemplate theme toggle navigates with ?theme=dark — the
    session args are the server's only honest theme signal."""
    try:
        import panel as pn
        vals = (pn.state.session_args or {}).get("theme")
        return bool(vals) and b"dark" in vals[0]
    except Exception:
        return False


def chart_colors() -> dict:
    """Per-session glyph colors that read on BOTH themes. Bars are
    translucent context — they must never compete with lines."""
    dark = is_dark()
    return {
        "dark": dark,
        "actual": "#f0f0f0" if dark else INK,         # theme-neutral ink
        "bar": STEEL, "bar_alpha": 0.28,              # translucent context
        "up": STEEL, "down": NAVY, "total": PASS_GREEN,  # waterfall trio
        "ref": GREY_TEXT,                             # dashed references
    }
