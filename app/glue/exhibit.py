"""Exhibit chrome — the workbook's visual language for app pages (W2b).

One kit so every exhibit reads like the Excel ones: navy 16pt title over a
grey-italic subtitle, a "Showing: …" filter echo, KPI chips (the
workbench's _chip HTML pattern, results.py:67), grey-italic footnote
lists, and the FILL_PANEL banner ("THE BRIDGE IN ONE LINE"). All
data-derived strings pass through md_safe — these panes render raw HTML.
"""
from __future__ import annotations

import panel as pn

from .format import md_safe
from .theme import FAIL_RED, NAVY

_CSS = f"""
<style>
.plw-ex-title {{ color: {NAVY}; font-size: 1.45em; font-weight: 700;
  font-family: var(--body-font, Segoe UI), sans-serif; margin: 0; }}
.plw-ex-sub {{ color: #808080; font-style: italic; font-size: 0.9em;
  margin: 2px 0 0 0; }}
.plw-ex-echo {{ color: #808080; font-style: italic; font-size: 0.85em;
  margin: 6px 0 0 0; }}
.plw-chips {{ display: flex; flex-wrap: wrap; gap: 8px; margin: 4px 0; }}
.plw-chip {{ border: 1px solid var(--panel-border-color, #e0e0e0);
  border-radius: 6px; padding: 4px 12px; min-width: 110px;
  font-family: var(--body-font, Segoe UI), sans-serif; }}
.plw-chip .l {{ font-size: 0.72em; opacity: 0.65; }}
.plw-chip .v {{ font-size: 1.22em; font-weight: 650; color: {NAVY};
  font-variant-numeric: tabular-nums; }}
.plw-chip .v.bad {{ color: {FAIL_RED}; }}
.plw-chip .s {{ font-size: 0.7em; opacity: 0.55; }}
.plw-notes {{ color: #595959; font-style: italic; font-size: 0.82em;
  margin: 6px 0; }}
.plw-notes li {{ margin: 2px 0 2px 1em; }}
.plw-banner {{ background: #F2F2F2; color: {NAVY}; font-weight: 700;
  padding: 7px 14px; border-radius: 4px; font-size: 0.95em;
  font-family: var(--body-font, Segoe UI), sans-serif; margin: 6px 0; }}
</style>
"""


def chip(label: str, value: str, sub: str = "", bad: bool = False) -> str:
    cls = "v bad" if bad else "v"
    s = f"<div class='s'>{md_safe(sub)}</div>" if sub else ""
    return (f"<div class='plw-chip'><div class='l'>{md_safe(label)}</div>"
            f"<div class='{cls}'>{md_safe(value)}</div>{s}</div>")


def chips_html(*chips: str) -> str:
    """The full chip-row html — for panes that update in place."""
    return _CSS + "<div class='plw-chips'>" + "".join(chips) + "</div>"


def chip_row(*chips: str) -> pn.pane.HTML:
    return pn.pane.HTML(chips_html(*chips), sizing_mode="stretch_width")


def exhibit_header(title: str, subtitle: str) -> pn.pane.HTML:
    """Static title + subtitle. For a live filter echo, pair with echo()."""
    return pn.pane.HTML(
        _CSS + f"<div class='plw-ex-title'>{md_safe(title)}</div>"
        f"<div class='plw-ex-sub'>{md_safe(subtitle)}</div>",
        sizing_mode="stretch_width")


def echo_pane() -> pn.pane.HTML:
    """An updatable 'Showing: …' line; call .object = echo_html(text)."""
    return pn.pane.HTML("", sizing_mode="stretch_width")


def echo_html(text: str) -> str:
    return _CSS + f"<div class='plw-ex-echo'>{md_safe(text)}</div>"


def note_list(notes: list) -> pn.pane.HTML:
    lis = "".join(f"<li>{md_safe(n)}</li>" for n in notes)
    return pn.pane.HTML(_CSS + f"<ul class='plw-notes'>{lis}</ul>",
                        sizing_mode="stretch_width")


def banner(text: str) -> pn.pane.HTML:
    return pn.pane.HTML(_CSS + f"<div class='plw-banner'>{md_safe(text)}"
                        "</div>", sizing_mode="stretch_width")
