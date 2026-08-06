"""The light workbook profile (D109).

A light build drops six tabs and the three `_calc` sections that exist only to
serve them. The failure mode is not subtle but it IS quiet: a skipped builder
registers no defined name, so anything still pointing at one says #NAME? — a
Checks panel full of errors on the sheet people are told to trust.

Excel finds that on a recalculation, and phase B does check it. These tests are
the cheap half: they run in seconds, with no Excel, and they fail on the two
static properties a recalculation would only confirm — nothing that survives
may reference a tab that did not, and nothing may point at a sheet that is not
there. The expensive half stays in the harness.
"""

from __future__ import annotations

import re
from dataclasses import replace

import pytest

from src.build_workbook import LIGHT_DROPS, build, load_config

KEEPERS = ("Read Me", "Control", "Inputs", "Rate Log", "Mod Log", "State Summary",
           "Program Flow", "Net Delivery", "Bridge", "LR Flow", "Walkthrough",
           "Flow Dashboard", "Rate Engine", "Mod Engine", "Checks", "Methodology",
           "_lists", "_calc", "_netcalc", "_oracle")


@pytest.fixture(scope="module")
def cfg():
    return load_config("config/config.yaml")


def _built(cfg, profile):
    name = cfg.output_lobs[0]
    lobs = tuple(replace(l, profile=profile) if l.name == name else l for l in cfg.lobs)
    return build(replace(cfg, lobs=lobs), name)


@pytest.fixture(scope="module")
def light(cfg):
    return _built(cfg, "light")


@pytest.fixture(scope="module")
def full(cfg):
    return _built(cfg, "full")


def test_full_is_the_default(cfg):
    """Nobody gets a trimmed workbook by accident."""
    assert all(l.profile == "full" and not l.light for l in cfg.lobs), \
        [(l.name, l.profile) for l in cfg.lobs]


def test_an_unknown_profile_is_refused(tmp_path):
    """A typo in the config must not silently build something else."""
    src = open("config/config.yaml", encoding="utf-8").read()
    bad = tmp_path / "bad.yaml"
    bad.write_text(src.replace("  - name: Property",
                               "  - name: Property\n    profile: lite", 1),
                   encoding="utf-8")
    with pytest.raises(ValueError, match="profile for Property"):
        load_config(bad)


def test_light_drops_exactly_the_named_tabs(light, full):
    assert set(full.sheetnames) - set(light.sheetnames) == set(LIGHT_DROPS)
    for keeper in KEEPERS:
        assert keeper in light.sheetnames, keeper


def test_the_load_bearing_tabs_are_never_candidates():
    """Bridge hosts 21 nr_* names that ten surviving sheets read; the two logs
    host the rl_*/ml_* families the engines are built on. Dropping any of them
    is not a lighter workbook, it is a broken one."""
    for essential in ("Bridge", "Rate Log", "Mod Log", "_calc", "_netcalc"):
        assert essential not in LIGHT_DROPS


def test_no_surviving_formula_mentions_a_dropped_tab(light):
    """The whole point. A reference to a tab that was not built is #REF! at
    best and a silently wrong lookup at worst."""
    offenders = []
    for ws in light.worksheets:
        for row in ws.iter_rows():
            for c in row:
                f = c.value
                if not isinstance(f, str) or not f.startswith("="):
                    continue
                for dropped in LIGHT_DROPS:
                    # 'One-Pager'!A1 or Portfolio!A1 — quoted when it has a space
                    if re.search(rf"'{re.escape(dropped)}'!|(?<![\w']){re.escape(dropped)}!", f):
                        offenders.append(f"{ws.title}!{c.coordinate} -> {dropped}")
    assert not offenders, offenders[:10]


def test_every_defined_name_points_at_a_sheet_that_exists(light):
    """A name registered against a skipped sheet would be #REF! on open."""
    bad = []
    for name, dn in light.defined_names.items():
        m = re.match(r"^'?([^'!]+)'?!", str(dn.attr_text))
        if m and m.group(1) not in light.sheetnames:
            bad.append(f"{name} -> {m.group(1)}")
    assert not bad, bad


def test_the_calc_sections_go_with_their_tabs(light, full):
    """Sections 3-5 of _calc read sc_*/att_*/slv_* — names the dropped tabs
    define. They have to go too, and they are the bulk of the saving: the six
    tabs themselves are barely 2,600 formulas."""
    def formulas(wb, sheet):
        return sum(1 for row in wb[sheet].iter_rows() for c in row
                   if isinstance(c.value, str) and c.value.startswith("="))

    saved = formulas(full, "_calc") - formulas(light, "_calc")
    assert saved > 20000, f"only {saved} _calc formulas dropped — sections still built?"


def test_the_read_me_does_not_advertise_a_tab_that_is_not_there(light):
    """Clicking a link to a missing sheet does nothing at all, which is a worse
    way to find out than simply not being offered it."""
    listed = {c.value for row in light["Read Me"].iter_rows() for c in row
              if isinstance(c.value, str)}
    for dropped in LIGHT_DROPS:
        assert dropped not in listed, dropped
