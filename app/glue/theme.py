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

PALETTE = [NAVY, "#c0392b", PASS_GREEN, "#8e44ad", "#d35400",
           "#16a085", STEEL, "#7f8c8d", "#c39bd3", "#e67e22"]
ACCENT = NAVY

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

# Uniform chips for multi-select pickers (widget-scoped stylesheet).
# Choices.js sizes each chip to its label, so one long name renders a
# gigantic chip next to tiny ones. Fixed width + ellipsis; the remove X is
# absolutely positioned inside the reserved right padding.
CHIP_CSS = """
.choices__list--multiple .choices__item {
  position: relative; box-sizing: border-box;
  width: 178px; max-width: 178px;
  overflow: hidden; text-overflow: ellipsis; white-space: nowrap;
  padding: 4px 26px 4px 10px; vertical-align: middle;
}
.choices[data-type*="select-multiple"] .choices__button {
  position: absolute; right: 6px; top: 50%; transform: translateY(-50%);
  margin: 0; padding-left: 0; border-left: none;
}
"""


def series_color(label: str, results: dict) -> str:
    """Stable color keyed by the label's position in the INSERTION-ordered
    results dict — not the selection order, so a series keeps its color as
    the selection changes."""
    keys = list(results)
    idx = keys.index(label) if label in keys else 0
    return PALETTE[idx % len(PALETTE)]


# ------------------------------------------------------- FloatPanel skin
# FloatPanels mount at BODY level, where the content root picks up
# DARK-flavored widget styling REGARDLESS of the app theme (measured live in
# the workbench). Float content is SELF-CONTAINED and always LIGHT.
FLOAT_BG = "#ffffff"
FLOAT_INK = "#1f2933"
FLOAT_CSS = """
:host { color: #1f2933; }
:host * { color: #1f2933 !important; }
label { color: #1f2933 !important; }
.bk-input, input.bk-input, select.bk-input, textarea.bk-input {
  background-color: #ffffff !important; color: #1f2933 !important;
  border-color: #c9ced4 !important;
}
code, pre { background-color: #eef1f4 !important; }
a { color: #1d5fa8 !important; }
"""


def float_styles(extra: dict | None = None) -> dict:
    s = {"background": FLOAT_BG, "color": FLOAT_INK}
    if extra:
        s.update(extra)
    return s


def force_float_skin(root):
    """Attach FLOAT_CSS to every stylesheet-capable component under ``root``
    (Buttons excepted). Panel appends user stylesheets after the design's
    and the rules carry !important, so they win."""
    import panel as pn

    def _walk(o):
        yield o
        for c in getattr(o, "objects", []) or []:
            yield from _walk(c)

    for o in _walk(root):
        if isinstance(o, pn.widgets.Button):
            continue
        if hasattr(o, "stylesheets"):
            try:
                if FLOAT_CSS not in (o.stylesheets or []):
                    o.stylesheets = list(o.stylesheets or []) + [FLOAT_CSS]
            except Exception:
                pass
    return root


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
        "actual": "#f0f0f0" if dark else "#333333",   # theme-neutral ink
        "bar": "#8a97a6", "bar_alpha": 0.28,          # translucent slate
        "up": STEEL, "down": NAVY, "total": PASS_GREEN,  # waterfall trio
        "ref": "#9a9a9a",                             # dashed references
    }


def line_color(color: str, dark=None) -> str:
    """A palette color adjusted for the active theme: lifted toward white on
    dark backgrounds (darker colors get more lift), nudged off near-black on
    light. Hue preserved; nothing goes neon."""
    if dark is None:
        dark = is_dark()
    r, g, b = (int(color[i:i + 2], 16) for i in (1, 3, 5))
    lum = 0.299 * r + 0.587 * g + 0.114 * b
    if dark and lum < 110:
        f = 0.25 + (110 - lum) / 110 * 0.55
        r, g, b = (int(c + (255 - c) * f) for c in (r, g, b))
    elif not dark and lum < 90:
        r, g, b = (int(c + (255 - c) * 0.45) for c in (r, g, b))
    return f"#{r:02x}{g:02x}{b:02x}"
