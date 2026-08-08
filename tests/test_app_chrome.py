"""W3a: the readability floor — pinned, not hoped for.

Four things users hit every session: the fast theme's white-on-hover
repaint (fixed by theme.TABLE_CSS on every Tabulator), tables that
scrolled 21-state exhibits inside fixed boxes (grow-to-content + cap),
the TOTAL pin that silently never worked (frozen_rows=[-1] resolved once
against an empty frame), and the Book table's sticky 20-row
auto-pagination at the real 378-combo size. The style helper's lever
border is the one piece that runs on both interpreters; the page checks
need the app venv.
"""
from __future__ import annotations

import sys
from pathlib import Path

import pandas as pd
import pytest

ROOT = Path(__file__).resolve().parents[1]
sys.path.insert(0, str(ROOT))

WB_PROP = ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"
WB_GL = ROOT / "output" / "Plan_LR_Workbook_2027_General_Liability.xlsx"

_needs_panel = pytest.mark.skipif(
    __import__("importlib").util.find_spec("panel") is None,
    reason="panel not installed (system interpreter — app venv runs this)")


# --------------------------------------------------- both interpreters

def test_lever_border_is_a_non_color_signal():
    """Lever blue and the LR scale's green are both pale-cool; the STEEL
    left border is what keeps 'you may type here' legible in greyscale."""
    from app import summary

    df = pd.DataFrame({
        "state": ["AZ", "CO", "TOTAL"], "ep": [100.0, 50.0, 150.0],
        "planlr": [0.6, 0.7, 0.65], "planlr1": [0.6, 0.7, 0.65],
        "netsel": [None, 0.1, None],
        "chg1_tok": ["T", "P", ""],
    })
    css = summary.ss_styles(df, levers=("netsel",))
    assert "border-left: 3px solid" in css.loc[0, "netsel"]
    assert "border-left" not in css.loc[2, "netsel"]        # TOTAL untinted
    assert "background-color: #EAF2FC" in css.loc[1, "netsel"]


def test_table_css_targets_the_cell_not_just_the_row():
    """The fast theme sets color on the ROW; a fix at equal specificity
    loses because Panel appends the theme stylesheet after ours. The rules
    must carry !important and land on .tabulator-cell."""
    from app.glue.theme import STEEL_LIGHT, TABLE_CSS

    assert ".tabulator-selectable:hover .tabulator-cell" in TABLE_CSS
    assert ".tabulator-selected .tabulator-cell" in TABLE_CSS
    assert TABLE_CSS.count("!important") >= 4
    assert STEEL_LIGHT in TABLE_CSS                # selection restyled, kept
    assert "tabular-nums" in TABLE_CSS


# ------------------------------------------------------- app venv only

@_needs_panel
def test_total_pin_follows_the_roster():
    """The pin is positional and Panel resolves negatives ONCE — as shipped
    in W2b it resolved against the empty build-time frame and TOTAL was
    never pinned; a filter change then repointed it at a state row."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page

    session = PlanSession()
    page = ss_page.build(session)
    assert page["table"].frozen_rows == []          # empty book: no pin
    session.replace_config(importers.from_workbook(WB_GL))
    session.replace_config(importers.from_workbook(WB_PROP))

    table = page["table"]
    n_all = len(table.value)
    assert n_all > 1
    assert table.frozen_rows == [n_all - 1]
    assert table.value.iloc[-1]["state"] == "TOTAL"

    page["filters"]["line"].value = ["Property"]    # roster shrinks
    n_prop = len(table.value)
    assert table.frozen_rows == [n_prop - 1]
    assert table.value.iloc[-1]["state"] == "TOTAL"


@_needs_panel
def test_summary_table_chrome():
    from app.glue.session import PlanSession
    from app.glue.theme import TABLE_CSS
    from app.pages import state_summary as ss_page

    table = ss_page.build(PlanSession())["table"]
    assert TABLE_CSS in table.stylesheets
    assert table.selectable is False               # selection is never read
    assert table.height is None                    # grow to the roster …
    assert table.max_height == 840                 # … capped
    # constructor pops `configuration` into the private slot (the W2a scar)
    assert table._configuration["columnDefaults"]["headerWordWrap"] is True
    assert "columns" not in table._configuration   # cannot coexist w/ groups
    assert table.header_tooltips                   # the dense-column hovers


@_needs_panel
def test_book_tables_chrome_and_explicit_pagination():
    from app.glue.session import PlanSession
    from app.glue.theme import TABLE_CSS
    from app.pages import book as book_page

    page = book_page.build(PlanSession())
    table, line_tbl = page["table"], page["line_tbl"]
    for t in (table, line_tbl):
        assert TABLE_CSS in t.stylesheets
        assert t.selectable is False               # cellClick survives this
        assert t.height is None
    assert line_tbl.max_height == 300
    assert table.max_height == 640
    # 378 real combos tripped Panel's silent local/20-per-page fallback and
    # it never reset — pagination is now a decision, and a fleet-sized
    # frame must not override it
    assert (table.pagination, table.page_size) == ("local", 100)
    big = pd.DataFrame({c: [None] * 250 for c in table.value.columns})
    big["lob"] = "L"
    big["bu"] = "B"
    big["state"] = [f"S{i}" for i in range(250)]
    table.value = big
    assert (table.pagination, table.page_size) == ("local", 100)


@_needs_panel
def test_input_grid_chrome_keeps_selection():
    """Grids keep selection (Ctrl+C copy-out rides it) — TABLE_CSS restyles
    it legible instead of removing it."""
    from app.glue.theme import TABLE_CSS
    from app.grids import LR_SCHEMA, make_grid

    grid = make_grid(LR_SCHEMA, [])
    assert TABLE_CSS in grid.stylesheets
    assert grid.height is None
    assert grid.max_height == 560
    assert grid.selectable                          # default kept ON


# ------------------------------------------------------------- W3g shell

@_needs_panel
def test_nav_labels_dropped_emoji_and_slugs_survived():
    """Renames are display-only: slugs are the bookmark contract."""
    from app.main import _PAGE_SLUGS, _PAGES

    assert set(_PAGE_SLUGS.values()) == {"inputs", "book", "summary",
                                         "flow", "combo", "scenarios",
                                         "guide"}
    assert list(_PAGE_SLUGS.values())[0] == "inputs"    # the start page
    for label in _PAGES:
        assert label.isascii(), label                   # no emoji carriers
    assert list(_PAGES) == ["Inputs", "Book", "State Summary", "Flow",
                            "Deep dive", "Scenarios", "Guide"]


@_needs_panel
def test_theme_consolidation_is_real():
    """The four-step grey ramp exists; the dead palette is gone; the one
    exhibit stylesheet opens with an explicit font and rides every helper
    pane."""
    from app.glue import theme
    from app.glue.exhibit import EXHIBIT_CSS

    for name in ("INK", "GREY_TEXT", "GREY_LINE", "GREY_FILL"):
        assert hasattr(theme, name), name
    for dead in ("series_color", "PALETTE", "CHIP_CSS", "FLOAT_CSS",
                 "float_styles", "force_float_skin", "line_color"):
        assert not hasattr(theme, dead), dead
    assert "font-family" in EXHIBIT_CSS and ":host" in EXHIBIT_CSS
    assert theme.GREY_TEXT in EXHIBIT_CSS
    cc = theme.chart_colors()
    assert set(cc) >= {"actual", "bar", "bar_alpha", "up", "down",
                       "total", "ref"}


@_needs_panel
def test_exhibit_helpers_carry_the_stylesheet():
    from app.glue.exhibit import EXHIBIT_CSS, banner, exhibit_header, \
        html_pane, note_list

    for pane in (exhibit_header("t", "s"), banner("b"), note_list(["n"]),
                 html_pane()):
        assert EXHIBIT_CSS in pane.stylesheets
    # markup builders are PURE — the style block never re-sends per update
    from app.glue.exhibit import chip, chips_html, echo_html
    assert "<style" not in chips_html(chip("l", "v"))
    assert "<style" not in echo_html("x")


@_needs_panel
def test_inputs_is_the_cold_start_and_context_bar_names_the_line():
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import inputs as inputs_page

    session = PlanSession()
    page = inputs_page.build(session)
    assert "No line active" in page["context"].object
    assert "masters above" in page["rail"].object.lower() or \
        "masters" in page["rail"].object.lower()

    session.replace_config(importers.from_workbook(
        ROOT / "output" / "Plan_LR_Workbook_2027_Property.xlsx"))
    assert "Editing: Property" in page["context"].object


# ------------------------------------------------------ W3h live punch list

@_needs_panel
def test_gate_hidden_skips_and_catches_up_on_show():
    """One keystroke was re-rendering every mounted page (John: 'the app
    seems to get hung up a lot'). The gate: hidden pages skip; the
    router's on_show call — which every page answers with a full render
    — catches them up. Unattached gates are permissive so build-time
    first renders pass."""
    import panel as pn

    from app.glue.bindings import gate_hidden

    ran = []
    gated = gate_hidden(lambda *e: ran.append(1))
    gated()
    assert len(ran) == 1                    # permissive before attach
    col = pn.Column()
    gated.attach(col)
    gated()
    assert len(ran) == 2                    # visible: runs
    col.visible = False
    gated()
    assert len(ran) == 2                    # hidden: skipped
    col.visible = True
    gated()
    assert len(ran) == 3


@_needs_panel
def test_hidden_summary_skips_the_render_until_shown():
    """Page-level: hide the Summary main, bump the bus, and the table
    must NOT recompute; on_show catches it up — the same contract the
    live router relies on."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import state_summary as ss_page

    session = PlanSession()
    page = ss_page.build(session)
    session.replace_config(importers.from_workbook(WB_PROP))
    st = page["table"].value.iloc[0]["state"]
    meta = page["holder"]["meta"]
    combos = meta.view[st]

    page["main"].visible = False
    for _lob, _s, _k, _bu, row, _res in combos:
        row["netp"] = 0.19
    session.bus.bump()                      # headless debounce = synchronous
    df = page["table"].value
    assert df.loc[df["state"] == st, "netsel"].iloc[0] != pytest.approx(0.19)

    page["main"].visible = True
    page["on_show"]()
    df = page["table"].value
    assert df.loc[df["state"] == st, "netsel"].iloc[0] == pytest.approx(0.19)


@_needs_panel
def test_every_heavy_page_returns_on_show():
    """The gate is only safe because the router's on_show call fully
    re-renders — every gated page must expose the hook (the Deep dive
    page shipped W3 without one)."""
    from app.glue.session import PlanSession
    from app.pages import book, combo, flow, inputs, scenarios, \
        state_summary

    session = PlanSession()
    for mod in (inputs, book, state_summary, flow, combo, scenarios):
        page = mod.build(session)
        assert callable(page.get("on_show")), mod.__name__


@_needs_panel
def test_active_line_selectors_on_inputs_and_deep_dive():
    """W3h: "i dont know how to change the line i am looking at on deep
    dive" — both editing surfaces carry an active-line Select that
    follows activation and switches it, without stealing the active
    line on programmatic refreshes."""
    from app import importers
    from app.glue.session import PlanSession
    from app.pages import combo as combo_page
    from app.pages import inputs as inputs_page

    session = PlanSession()
    ipage = inputs_page.build(session)
    cpage = combo_page.build(session)
    session.replace_config(importers.from_workbook(WB_GL))
    session.replace_config(importers.from_workbook(WB_PROP))

    for page in (ipage, cpage):
        sel = page["line_sel"]
        assert set(sel.options) == {"Property", "General Liability"}
        assert sel.value == "Property"      # follows the ACTIVE line

    # switching on Deep dive activates and the Inputs selector follows
    cpage["line_sel"].value = "General Liability"
    assert session.page.config.lob == "General Liability"
    assert ipage["line_sel"].value == "General Liability"
    assert "General Liability" in ipage["context"].object

    # activation elsewhere (Book click-through path) moves both selectors
    session.activate("Property")
    assert cpage["line_sel"].value == "Property"
    assert ipage["line_sel"].value == "Property"
