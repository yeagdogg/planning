"""Exhibit chrome — the workbook's visual language for app pages (W2b,
consolidated W3g).

One kit so every exhibit reads like the Excel ones: navy title over a
grey-italic subtitle, a "Showing: …" filter echo, KPI chips (the
workbench's _chip HTML pattern), grey-italic footnote lists, and the
FILL_PANEL banner. All data-derived strings pass through md_safe — these
panes render raw HTML.

W3g: ONE stylesheet (``EXHIBIT_CSS``), attached per-pane via
``stylesheets=`` at construction — never re-sent inside every ``.object``
update — built from the theme constants (the four-step grey ramp, the
four-step type scale) and opening with an explicit font so custom HTML
matches the Fast widgets instead of falling back to Segoe UI through an
undefined variable. Pages that update chip/echo HTML in place should
create their panes with ``html_pane()`` so the stylesheet rides along.
"""
from __future__ import annotations

import panel as pn

from .format import md_safe
from .theme import FAIL_RED, GREY_FILL, GREY_LINE, GREY_TEXT, NAVY, \
    PASS_GREEN, STEEL

# The four-step type scale: 1.45 title / 1.3 value / 0.85 small / 0.72
# micro (body = 1.0 implicit). Eight ad-hoc em sizes collapsed here.
EXHIBIT_CSS = f"""
:host {{ font-family: "Open Sans", "Segoe UI", sans-serif; }}
.plw-ex-title {{ color: {NAVY}; font-size: 1.45em; font-weight: 700;
  margin: 0; }}
.plw-ex-sub {{ color: {GREY_TEXT}; font-style: italic; font-size: 0.85em;
  margin: 2px 0 0 0; }}
.plw-ex-echo {{ color: {GREY_TEXT}; font-style: italic; font-size: 0.85em;
  margin: 6px 0 0 0; }}
.plw-chips {{ display: flex; flex-wrap: wrap; gap: 8px; margin: 4px 0; }}
.plw-chip {{ border: 1px solid {GREY_LINE};
  border-radius: 6px; padding: 4px 12px; min-width: 110px; }}
.plw-chip .l {{ font-size: 0.72em; opacity: 0.65; }}
.plw-chip .v {{ font-size: 1.3em; font-weight: 650; color: {NAVY};
  font-variant-numeric: tabular-nums; }}
.plw-chip .v.bad {{ color: {FAIL_RED}; }}
.plw-chip .s {{ font-size: 0.72em; opacity: 0.55; }}
.plw-notes {{ color: {GREY_TEXT}; font-style: italic; font-size: 0.85em;
  margin: 6px 0; }}
.plw-notes li {{ margin: 2px 0 2px 1em; }}
.plw-banner {{ background: {GREY_FILL}; color: {NAVY}; font-weight: 700;
  padding: 7px 14px; border-radius: 4px; font-size: 0.95em;
  margin: 6px 0; }}
.plw-cards {{ display: flex; flex-wrap: wrap; gap: 10px; }}
.plw-cards .c {{ border: 1px solid {GREY_LINE}; border-radius: 8px;
  padding: 8px 16px; min-width: 128px; }}
.plw-cards .v {{ font-size: 1.3em; font-weight: 700; color: {NAVY};
  font-variant-numeric: tabular-nums; }}
.plw-cards .l {{ color: {GREY_TEXT}; font-size: 0.72em; }}
.plw-cards .bad {{ color: {FAIL_RED}; }}
.plw-bridge h3 {{ color: {NAVY}; margin: 0 0 2px 0; }}
.plw-bridge .sub {{ color: {GREY_TEXT}; font-size: 0.85em;
  margin-bottom: 10px; }}
.plw-bridge table {{ border-collapse: collapse; min-width: 460px; }}
.plw-bridge td, .plw-bridge th {{ padding: 5px 14px; text-align: right;
  border-bottom: 1px solid {GREY_LINE};
  font-variant-numeric: tabular-nums; }}
.plw-bridge td:first-child {{ text-align: left; }}
.plw-bridge tr.total td {{ border-top: 2px solid {NAVY}; font-weight: 700;
  color: {NAVY}; }}
.plw-bridge .pts {{ color: {GREY_TEXT}; }}
.plw-badge {{ display: inline-block; padding: 1px 8px; border-radius: 9px;
  font-size: 0.72em; color: #fff; background: {STEEL};
  vertical-align: middle; margin-left: 8px; }}
.plw-badge.net {{ background: {PASS_GREEN}; }}
.plw-err {{ color: {FAIL_RED}; font-weight: 600; }}
"""


def html_pane(**kwargs) -> pn.pane.HTML:
    """An updatable HTML pane that carries the exhibit stylesheet — the
    construction every chips/cards/echo `.object =` site should use."""
    kwargs.setdefault("sizing_mode", "stretch_width")
    return pn.pane.HTML("", stylesheets=[EXHIBIT_CSS], **kwargs)


def chip(label: str, value: str, sub: str = "", bad: bool = False) -> str:
    cls = "v bad" if bad else "v"
    s = f"<div class='s'>{md_safe(sub)}</div>" if sub else ""
    return (f"<div class='plw-chip'><div class='l'>{md_safe(label)}</div>"
            f"<div class='{cls}'>{md_safe(value)}</div>{s}</div>")


def chips_html(*chips: str) -> str:
    """The chip-row markup — pure HTML; the pane's stylesheet styles it."""
    return "<div class='plw-chips'>" + "".join(chips) + "</div>"


def chip_row(*chips: str) -> pn.pane.HTML:
    p = html_pane()
    p.object = chips_html(*chips)
    return p


def exhibit_header(title: str, subtitle: str) -> pn.pane.HTML:
    """Static title + subtitle. For a live filter echo, pair with echo()."""
    p = html_pane()
    p.object = (f"<div class='plw-ex-title'>{md_safe(title)}</div>"
                f"<div class='plw-ex-sub'>{md_safe(subtitle)}</div>")
    return p


def echo_pane() -> pn.pane.HTML:
    """An updatable 'Showing: …' line; call .object = echo_html(text)."""
    return html_pane()


def echo_html(text: str) -> str:
    return f"<div class='plw-ex-echo'>{md_safe(text)}</div>"


def note_list(notes: list) -> pn.pane.HTML:
    p = html_pane()
    p.object = ("<ul class='plw-notes'>"
                + "".join(f"<li>{md_safe(n)}</li>" for n in notes)
                + "</ul>")
    return p


def banner(text: str) -> pn.pane.HTML:
    p = html_pane()
    p.object = f"<div class='plw-banner'>{md_safe(text)}</div>"
    return p
