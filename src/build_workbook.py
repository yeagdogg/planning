"""Workbook generator CLI: builds the Calendar-Year Plan Loss Ratio workbooks.

Usage:
    python -m src.build_workbook --config config/config.yaml --out output/
    python -m src.build_workbook --lob "Property"          # single LOB only

Dimensioning (DECISIONS.md D37): ONE WORKBOOK PER LINE OF BUSINESS. Within a
workbook, every input row is one BUSINESS UNIT x STATE combination; the
concatenated selector key is "BU|State". The policy term is a workbook-level
parameter (all combos share the workbook's LOB); tbl_Seasonality is keyed by
state. Table capacities scale from the configured dimensions.

Architecture (DECISIONS.md D24): this module owns configuration, sample data,
oracle baking, the master cell layout (:class:`Layout`), the named-range
registry (:class:`Ctx`), and assembly order. The per-sheet builders live in
``sheets_inputs / sheets_engine / sheets_calc / sheets_main / sheets_report``.

Every formula the builders emit is classic-mode (Excel-2007-era functions
only); the oracle in ``src/engine.py`` is the source of truth the generated
workbooks must tie to (verified by ``tools/verify_workbook.py``).
"""

from __future__ import annotations

import argparse
import datetime as dt
import shutil
import sys
from dataclasses import dataclass, field, replace
from pathlib import Path

import yaml
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName

from . import engine
from .engine import ComboInputs, ModInputs, RateChange
from .xlstyle import quote_sheet

GENERATOR_VERSION = "3.4.2"

# Leads every seeded Mod Log comment so the Checks tripwire can spot a sample
# action left in a real book (D80). Kept short and unmistakable — a real
# comment will not start with it by accident.
SAMPLE_MOD_TAG = "SAMPLE:"


def mod_adj_on(row_expr: str) -> str:
    """Excel test: is the mod adjustment in force for this tbl_LR row? (D81)

    ONE definition, because there used to be two. The visible Bridge read a
    blank or misspelled token as ON (only a literal "OFF" turned it off) while
    _calc, the Solver and _netcalc read it as OFF (only a literal "ON" turned
    it on) — so a token cleared by a paste left Portfolio and State Summary
    quietly dropping mod drift for combos the Bridge still adjusted, and only
    the SELECTED combo tripped an existing check.

    ON-default wins: it matches engine.ComboInputs.mod_adjustment_enabled
    (True) and the visible chain, and the Checks token-validity row now makes a
    malformed token loud rather than letting either default paper over it.
    """
    return f'AND(nr_ModAdjMaster="ON",INDEX(lr_modadj,{row_expr})<>"OFF")'


class SHEETS:
    """Canonical sheet names. New code references these constants so a rename
    is a one-line change; the harness imports them for mutation targets."""
    README = "Read Me"
    CONTROL = "Control"
    INPUTS = "Inputs"
    RATE_LOG = "Rate Log"          # created by the C6 rate-log split
    MOD_LOG = "Mod Log"            # D70 stepped schedule-mod actions
    NET_DELIVERY = "Net Delivery"  # D57 net-target decomposition
    NETCALC = "_netcalc"           # its hidden engine blocks
    PROGRAM_FLOW = "Program Flow"  # D59 descriptive program-flow grids
    RATE_ENGINE = "Rate Engine"
    MOD_ENGINE = "Mod Engine"
    BRIDGE = "Bridge"
    WALKTHROUGH = "Walkthrough"
    ONE_PAGER = "One-Pager"
    COMPARE = "Compare"
    FLOW = "Flow Dashboard"
    PORTFOLIO = "Portfolio"
    STATE_SUMMARY = "State Summary"
    SCENARIOS = "Scenarios"
    SOLVER = "Solver"
    ATTRIBUTION = "Attribution"
    CHECKS = "Checks"
    METHODOLOGY = "Methodology"
    LISTS = "_lists"
    CALC = "_calc"
    ORACLE = "_oracle"

# ---------------------------------------------------------------------------
# Configuration
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class LobCfg:
    name: str
    term_months: int = 12


@dataclass(frozen=True)
class Config:
    plan_year: int
    business_units: tuple[str, ...]
    states: tuple[str, ...]
    lobs: tuple[LobCfg, ...]
    seasonality_profiles: dict
    output_lobs: tuple[str, ...]  # LOB names to generate
    output_dir: str
    filename: str
    formula_mode: str
    solver_max_rate: float
    solver_min_post_share: float
    font: str
    version: str
    spare_lr_rows: int
    rate_log_rows: int
    mod_log_rows: int
    spare_seasonality_rows: int

    @property
    def combos(self) -> list[tuple[str, str]]:
        """(BU, state) pairs, BU-major — the row order of tbl_LR."""
        return [(bu, st) for bu in self.business_units for st in self.states]

    def lob(self, name: str) -> LobCfg:
        for l in self.lobs:
            if l.name == name:
                return l
        raise KeyError(name)


def _validate_dim_name(kind: str, name: str):
    if any(ch in name for ch in "*?|"):
        raise ValueError(f"{kind} name {name!r} may not contain * ? or | characters")


def load_config(path: str | Path) -> Config:
    raw = yaml.safe_load(Path(path).read_text(encoding="utf-8"))
    lobs = tuple(
        LobCfg(name=l["name"], term_months=int(l.get("term_months", 12)))
        for l in raw["lobs"]
    )
    for l in lobs:
        if not 1 <= l.term_months <= 12:
            raise ValueError(f"term_months for {l.name} must be 1..12, got {l.term_months}")
        _validate_dim_name("LOB", l.name)
    profiles = raw.get("seasonality_profiles") or {}
    for name, weights in profiles.items():
        if len(weights) != 12 or min(weights) < 0 or sum(weights) <= 0:
            raise ValueError(f"seasonality profile {name!r} must be 12 non-negative weights")
    fmode = raw.get("formula_mode", "classic")
    if fmode != "classic":
        raise ValueError(
            f"formula_mode {fmode!r} is not implemented; only 'classic' ships (DECISIONS.md D18)."
        )
    bus = tuple(str(b) for b in raw["business_units"])
    states = tuple(str(s) for s in raw["states"])
    if not states:
        raise ValueError("at least one state is required")
    for b in bus:
        _validate_dim_name("business unit", b)
    for s in states:
        _validate_dim_name("state", s)
    # a repeated dimension value silently breaks the workbook rather than the
    # build: keys stop being unique, every MATCH resolves to the first row, and
    # the live roster (which extracts UNIQUES) runs one row short of the
    # configured count. Fail loudly at load instead (D62).
    for kind, values in (("business unit", bus), ("state", states)):
        dupes = sorted({v for v in values if list(values).count(v) > 1})
        if dupes:
            raise ValueError(
                f"duplicate {kind}(s) in the config: {', '.join(dupes)} - each "
                f"{kind} must appear once (BU x state is the row key)."
            )
    out = raw.get("output", {})
    out_lobs = out.get("lobs", "all")
    if out_lobs == "all":
        out_lobs = tuple(l.name for l in lobs)
    else:
        out_lobs = tuple(str(x) for x in out_lobs)
        for name in out_lobs:
            if name not in {l.name for l in lobs}:
                raise ValueError(f"output.lobs entry {name!r} is not a configured LOB")
    cap = raw.get("table_capacity", {})
    solver = raw.get("solver", {})
    return Config(
        plan_year=int(raw["plan_year"]),
        business_units=bus,
        states=states,
        lobs=lobs,
        seasonality_profiles=profiles,
        output_lobs=out_lobs,
        output_dir=out.get("directory", "output"),
        filename=out.get("filename", "Plan_LR_Workbook_{plan_year}_{lob}"),
        formula_mode=fmode,
        solver_max_rate=float(solver.get("max_reasonable_rate", 0.25)),
        solver_min_post_share=float(solver.get("min_post_share", 0.05)),
        font=(raw.get("theme") or {}).get("font", "Calibri"),
        version=str(raw.get("version", GENERATOR_VERSION)),
        spare_lr_rows=int(cap.get("spare_lr_rows", 6)),
        rate_log_rows=int(cap.get("rate_log_rows", 240)),
        mod_log_rows=int(cap.get("mod_log_rows", 120)),
        spare_seasonality_rows=int(cap.get("spare_seasonality_rows", 3)),
    )


# ---------------------------------------------------------------------------
# Sample data (clearly marked in-sheet; §4 / DECISIONS.md D22, D37)
#
# One sample book per LOB workbook: one row per BU x state, deterministic
# variation (no randomness — regeneration is reproducible). Dates are relative
# to plan year P. For every 12-month-term LOB, BU-A | <first state> is EXACTLY
# the §9 worked example (asserted in build()).
# ---------------------------------------------------------------------------

LOB_BASE = {
    "Property": dict(lr=0.65, mind=0.850, ep=45),
    "General Liability": dict(lr=0.62, mind=0.950, ep=35),
    "Commercial Auto": dict(lr=0.68, mind=1.000, ep=40),
    "Workers Comp": dict(lr=0.60, mind=0.900, ep=30),
    "Umbrella": dict(lr=0.55, mind=1.000, ep=12),
    "Inland Marine": dict(lr=0.58, mind=0.920, ep=8),
}
_DEFAULT_BASE = dict(lr=0.63, mind=0.950, ep=25)


def _bu(cfg: Config, i: int) -> str:
    """The i-th configured business unit, clamped to the roster.

    Sample seeding addresses BUs POSITIONALLY, never by literal name: the
    config is meant to be renamed to a real book (D62), and the seeded
    worked example / showcase combos must follow the rename.
    """
    bus = cfg.business_units
    return bus[min(i, len(bus) - 1)]


def _clamp_mod(x: float) -> float:
    return round(min(1.5, max(0.5, x)), 3)


def sample_lr_rows(cfg: Config, lob: LobCfg) -> list[dict]:
    p = cfg.plan_year
    base = LOB_BASE.get(lob.name, _DEFAULT_BASE)
    asof = dt.date(p - 1, 9, 30)
    st = cfg.states
    rows = []
    for i, (bu, state) in enumerate(cfg.combos):
        lr = round(base["lr"] + ((i * 7) % 11 - 5) * 0.004, 4)
        m_ind = base["mind"]
        m0 = _clamp_mod(m_ind + ((i * 5) % 5 - 2) * 0.005)
        m1 = _clamp_mod(m0 + ((i * 3) % 7 - 2) * 0.005)
        row = dict(
            bu=bu, state=state, lr_proj=lr, basis="current", s=0.05,
            m_ind=m_ind, m0=m0, m0_asof=asof, m1=m1, m_prior=None, m2=None,
            m_end_prior=_clamp_mod((m0 + m1) / 2.0),   # D70 educated guess
            ep=base["ep"] * (8 + (i * 13) % 17) * 100,  # ADJUSTED plan EP in 000s
            trend=None,  # blank inherits the Control default trend (D38)
            a_other=1.0, a_other_label="", modadj="ON",
            netp=None, netp1=None,  # blank = explicit rate program (D39)
        )
        # showcase combos (state indices are positions in the configured list)
        if bu == _bu(cfg, 0) and state == st[0]:
            row.update(lr_proj=0.65, s=0.05, m_ind=0.850, m0=0.860, m1=0.890)  # §9
        elif bu == _bu(cfg, 1) and len(st) > 1 and state == st[1]:
            row.update(basis="proposed", s=0.08, trend=0.01)  # explicit trend override demo
        elif bu == _bu(cfg, 1) and len(st) > 2 and state == st[2]:
            row.update(a_other=1.02, a_other_label="Commission timing")
        elif bu == _bu(cfg, 0) and len(st) > 3 and state == st[3]:
            row.update(modadj="OFF", m0=m_ind, m1=m_ind)
        elif bu == _bu(cfg, 0) and len(st) > 5 and state == st[5]:
            row.update(m_prior=_clamp_mod(m0 - 0.010), m2=_clamp_mod(m1 + 0.005))
        elif bu == _bu(cfg, 2) and len(st) > 6 and state == st[6]:
            row.update(s=0.065)  # matches the seed-pattern rate-log row (D36)
        elif bu == _bu(cfg, 2) and len(st) > 8 and state == st[8]:
            # net-rate-selection showcase (D39): +10% taken 10/1/(P-1) in the
            # log, product targets +10% NET from 1/1/P — specifics TBD
            row.update(netp=0.10)
        rows.append(row)
    return rows


def degenerate_combo(cfg: Config) -> tuple[str, str]:
    """The showcase combo seeded with NO rate-log rows (factors = 1.000)."""
    return (_bu(cfg, 2), cfg.states[min(4, len(cfg.states) - 1)])


def sample_rate_rows(cfg: Config, lob: LobCfg) -> list[dict]:
    p = cfg.plan_year
    st = cfg.states

    def row(bu, state, y_off, m, d, filed, status, cons, ach=None, comment=""):
        return dict(bu=bu, state=state, eff=dt.date(p + y_off, m, d), filed=filed,
                    status=status, considered=cons, achievement=ach, comment=comment)

    rows = [
        # §9 worked example (BU-A | first state)
        row(_bu(cfg, 0), st[0], -2, 7, 1, 0.04, "taken", "Y", comment="Worked example (§9)"),
        row(_bu(cfg, 0), st[0], -1, 7, 1, 0.06, "taken", "Y", comment="Worked example (§9)"),
        row(_bu(cfg, 0), st[0], 0, 4, 1, 0.05, "planned", "N", 1.00, "Planned spring increase"),
    ]
    if len(st) > 6:
        rows.append(row(_bu(cfg, 2), st[6], -1, 7, 1, 0.065, "planned", "N", 1.00,
                        "Pattern: indication's selected change entered as a planned row"))
    if len(st) > 2:
        rows.append(row(_bu(cfg, 1), st[2], -1, 4, 1, 0.050, "taken", "Y"))
        rows.append(row(_bu(cfg, 1), st[2], 0, 5, 15, 0.030, "planned", "N", 1.00,
                        "Mid-month effective date example"))
    if len(st) > 3:
        rows.append(row(_bu(cfg, 1), st[3], -1, 10, 1, 0.080, "taken", "Y"))
        rows.append(row(_bu(cfg, 1), st[3], 0, 3, 1, 0.060, "planned", "N", 0.80,
                        "80% achievement assumed"))
    if len(st) > 5:
        rows.append(row(_bu(cfg, 0), st[5], -1, 6, 1, 0.045, "taken", "Y"))
    if len(st) > 8:
        rows.append(row(_bu(cfg, 2), st[8], -1, 10, 1, 0.10, "taken", "N",
                        comment="Net-selection showcase: fall filing (after the indication) "
                                "feeds the +10% net target"))
        # the net selection supersedes this row in the ENGINE (D39); Net
        # Delivery adopts it as the default filing whose delivery it
        # decomposes, so the tab demonstrates the pickup on first open (D63)
        rows.append(row(_bu(cfg, 2), st[8], 0, 4, 1, 0.04, "planned", "N", 1.00,
                        "Planned filing the net-delivery default picks up"))

    # deterministic coverage across the remaining book: most combos get a taken
    # row in P-1; roughly a quarter also carry a planned row in P.
    special = {(_bu(cfg, 0), st[0]), degenerate_combo(cfg)}
    if len(st) > 2:
        special.add((_bu(cfg, 1), st[2]))
    if len(st) > 3:
        special.add((_bu(cfg, 1), st[3]))
    if len(st) > 6:
        special.add((_bu(cfg, 2), st[6]))
    if len(st) > 8:
        special.add((_bu(cfg, 2), st[8]))
    for i, (bu, state) in enumerate(cfg.combos):
        if (bu, state) in special:
            continue
        if i % 3 != 2:
            rows.append(row(bu, state, -1, (i % 12) + 1, 1,
                            round(0.02 + (i % 5) * 0.01, 3), "taken", "Y"))
        if i % 4 == 1:
            rows.append(row(bu, state, 0, (i % 10) + 1, 1,
                            round(0.03 + (i % 4) * 0.01, 3), "planned", "N", 1.00))
    if len(rows) > cfg.rate_log_rows:
        raise ValueError(
            f"sample rate log ({len(rows)} rows) exceeds table_capacity.rate_log_rows "
            f"({cfg.rate_log_rows})")
    return rows


def sample_seasonality_rows(cfg: Config) -> list[dict]:
    """Seed a couple of state rows with a named profile (marked sample)."""
    profile = next(iter(cfg.seasonality_profiles.values()), None)
    if profile is None:
        return []
    picks = [s for j, s in enumerate(cfg.states) if j in (6, 7)]
    return [dict(state=s, weights=list(profile)) for s in picks]


# ---------------------------------------------------------------------------
# Sample rows -> oracle inputs
# ---------------------------------------------------------------------------


def sample_mod_rows(cfg: Config, lob: LobCfg) -> list[dict]:
    """Sample stepped mod actions (D70).

    Deliberately NOT on the worked-example combo: build() asserts that combo
    still reproduces engine.worked_example() exactly, and a mod action would
    move it. The seeds demonstrate the three things the mechanic adds — a
    dated action, partial achievement, and two actions compounding.

    Every comment leads with SAMPLE_MOD_TAG so the Checks tripwire can see a
    seeded action that outlived its welcome. These rows attach POSITIONALLY
    (2nd BU x 3rd state, 3rd BU x 7th state), so on a real roster they land on
    real combos and move real plan LRs — the failure the tag exists to catch.
    """
    p = cfg.plan_year
    st = cfg.states

    def row(bu, state, m, d, chg, status, ach=None, comment=""):
        return dict(bu=bu, state=state, eff=dt.date(p, m, d), chg=chg,
                    status=status, achievement=ach,
                    comment=f"{SAMPLE_MOD_TAG} {comment}")

    rows: list[dict] = []
    if len(st) > 2:
        rows.append(row(_bu(cfg, 1), st[2], 3, 1, 0.020, "taken",
                        comment="Directed mod tightening, already actioned"))
        rows.append(row(_bu(cfg, 1), st[2], 9, 1, 0.015, "planned", 0.70,
                        comment="Second push; 70% achievement assumed"))
    if len(st) > 6:
        rows.append(row(_bu(cfg, 2), st[6], 4, 15, 0.025, "planned", 0.80,
                        comment="Mid-month action — day-blended like a filing"))
    return rows


def sample_to_combo(
    cfg: Config,
    lob: LobCfg,
    lr_row: dict,
    rate_rows: list[dict],
    mod_rows: list[dict] | None = None,
    seasonality_on: bool = False,
    season_rows: list[dict] | None = None,
    trend_default: float = 0.0,
) -> ComboInputs:
    """Assemble engine inputs for one BU x state combo, mirroring workbook logic."""
    season = None
    if seasonality_on and season_rows:
        for srow in season_rows:
            if srow["state"] == lr_row["state"] and sum(srow["weights"]) > 0:
                season = tuple(srow["weights"])
    changes = tuple(
        RateChange(
            effective=r["eff"],
            filed_pct=r["filed"],
            status=r["status"],
            considered=(r["considered"] == "Y"),
            achievement=1.0 if r["achievement"] is None else r["achievement"],
            comment=r["comment"],
        )
        for r in rate_rows
        if r["bu"] == lr_row["bu"] and r["state"] == lr_row["state"]
    )
    mod_changes = tuple(
        RateChange(
            effective=r["eff"],
            filed_pct=r["chg"],
            status=r["status"],
            considered=False,
            achievement=1.0 if r["achievement"] is None else r["achievement"],
            comment=r["comment"],
        )
        for r in (mod_rows or [])
        if r["bu"] == lr_row["bu"] and r["state"] == lr_row["state"]
    )
    return ComboInputs(
        mod_changes=mod_changes,
        lr_proj=lr_row["lr_proj"],
        lr_basis=lr_row["basis"],
        sel_change=lr_row["s"],
        mods=ModInputs(
            m_ind=lr_row["m_ind"], m0=lr_row["m0"], m0_asof=lr_row["m0_asof"],
            m1=lr_row["m1"], m_prior=lr_row["m_prior"], m2=lr_row["m2"],
            m_end_prior=lr_row.get("m_end_prior"),
        ),
        rate_changes=changes,
        net_trend=trend_default if lr_row["trend"] is None else lr_row["trend"],
        a_other=lr_row["a_other"],
        a_other_label=lr_row["a_other_label"],
        # D81: ON-default, matching the workbook's single mod_adj_on() test and
        # ComboInputs' own default — a blank or misspelled token no longer means
        # OFF here and ON on the Bridge.
        mod_adjustment_enabled=(lr_row["modadj"] != "OFF"),
        term_months=lob.term_months,
        seasonality=season,
        plan_ep=lr_row["ep"],
        net_sel_p=lr_row.get("netp"),
        net_sel_p1=lr_row.get("netp1"),
    )


# ---------------------------------------------------------------------------
# Layout: every anchor the sheet builders share. 1-based rows/columns.
#
# Sizes depend on the configured dimensions, so call Layout.configure(cfg)
# before building or verifying (build() does this). Class-level attributes
# keep all L.XXX call sites unchanged; generation is single-threaded.
# ---------------------------------------------------------------------------


class Layout:
    # ---- Inputs sheet (dynamic; set by configure) ----
    # Workbook parameters live ABOVE the tables so the term input is never a
    # 350-row scroll away; the tables cascade below them.
    WB_HDR = 7            # workbook-parameters block (LOB + term)
    TERM_ROW = 8
    LR_HDR = 11
    LR_FIRST = 12
    LR_ROWS = 69
    LR_LAST = 80
    RL_HDR = 6            # Rate Log sheet (own sheet; fixed header)
    RL_FIRST = 7
    RL_ROWS = 240
    RL_LAST = 246
    ML_HDR = 6            # Mod Log sheet (D70; own sheet, same shape)
    ML_FIRST = 7
    ML_ROWS = 120
    ML_LAST = 126
    SE_HDR = 84
    SE_FIRST = 85
    SE_ROWS = 24
    SE_LAST = 108

    # ---- Rate Engine sheet (fixed) ----
    RE_COH_HDR = 16
    RE_COH_FIRST = 17
    N_COH = 48
    RE_COH_LAST = RE_COH_FIRST + N_COH - 1
    RE_MATRIX_COL = 22   # column V — S/T/U hold the net-selection columns (D39)
    N_MONTHS = 36
    RE_MATRIX_LAST_COL = RE_MATRIX_COL + N_MONTHS - 1
    RE_ROW_NUM = 66
    RE_ROW_DEN = 67
    RE_ROW_EM = 68
    RE_ROW_MINW = 69
    RE_ROW_MAXW = 70
    RE_ROW_VIOL = 71
    RE_ROW_MODM = 72
    RE_ROW_MODNUM = 73   # monthly mod numerator SUM w-e-M (quarterly price chart)
    RE_ROW_MIDX = 15

    # ---- Mod Engine ----
    ME_COH_FIRST = 17
    ME_COH_LAST = ME_COH_FIRST + N_COH - 1

    # ---- _calc sheet (dynamic) ----
    CALC_RES_HDR = 3
    CALC_RES_FIRST = 4
    CALC_RES_LAST = 72
    CALC_BLOCK_FIRST = 76
    CALC_BLOCK_STRIDE = 54

    # ---- Portfolio (dynamic) ----
    PF_HDR = 5
    PF_FIRST = 6
    PF_LAST = 74
    PF_W_TOTAL = 76
    PF_S_TOTAL = 77

    # ---- Bridge (fixed; answer-first layout — resolver demoted below) ----
    BR_WF_HDR = 8
    BR_WF_FIRST = 9
    BR_IN_FIRST = 38     # resolver block rows (outline-grouped audit detail)
    BR_CHART_DATA = 60

    # ---- Checks (fixed) ----
    CK_HDR = 5
    CK_FIRST = 6

    @classmethod
    def configure(cls, cfg: Config):
        cls.WB_HDR = 7
        cls.TERM_ROW = cls.WB_HDR + 1
        cls.LR_HDR = cls.TERM_ROW + 3
        cls.LR_FIRST = cls.LR_HDR + 1
        cls.LR_ROWS = len(cfg.business_units) * len(cfg.states) + cfg.spare_lr_rows
        cls.LR_LAST = cls.LR_FIRST + cls.LR_ROWS - 1
        # the rate log lives on its OWN sheet (fixed header rows there);
        # seasonality follows tbl_LR directly on Inputs
        cls.RL_HDR = 6
        cls.RL_FIRST = cls.RL_HDR + 1
        cls.RL_ROWS = cfg.rate_log_rows
        cls.RL_LAST = cls.RL_FIRST + cls.RL_ROWS - 1
        # the mod log (D70) mirrors it on its own sheet
        cls.ML_HDR = 6
        cls.ML_FIRST = cls.ML_HDR + 1
        cls.ML_ROWS = cfg.mod_log_rows
        cls.ML_LAST = cls.ML_FIRST + cls.ML_ROWS - 1
        cls.SE_HDR = cls.LR_LAST + 4
        cls.SE_FIRST = cls.SE_HDR + 1
        cls.SE_ROWS = len(cfg.states) + cfg.spare_seasonality_rows
        cls.SE_LAST = cls.SE_FIRST + cls.SE_ROWS - 1
        cls.CALC_RES_LAST = cls.CALC_RES_FIRST + cls.LR_ROWS - 1
        cls.CALC_BLOCK_FIRST = cls.CALC_RES_LAST + 4
        cls.PF_LAST = cls.PF_FIRST + cls.LR_ROWS - 1
        cls.PF_W_TOTAL = cls.PF_LAST + 2
        cls.PF_S_TOTAL = cls.PF_LAST + 3


# ---------------------------------------------------------------------------
# Build context: workbook + config + oracle + named-range registry
# ---------------------------------------------------------------------------


@dataclass
class Ctx:
    cfg: Config
    lob: LobCfg
    wb: Workbook
    lr_rows: list
    rate_rows: list
    mod_rows: list
    season_rows: list
    oracle_m: engine.EngineResult      # worked-example combo, monthly
    oracle_c: engine.EngineResult      # worked-example combo, continuous
    oracle_solver_r: float
    we_key: str                        # worked-example combo key "BU-A|<state>"
    we_row: dict
    carried: object | None = None      # carry.CarriedInputs when rebuilt from a file (D82)
    names: dict = field(default_factory=dict)   # name -> (ref, description)
    lay_dyn: dict = field(default_factory=dict)  # dynamic anchors shared between builders

    def define(self, name: str, sheet: str, ref: str, description: str):
        """Register a workbook-level defined name (documented on Methodology).

        Excel resolves defined names CASE-INSENSITIVELY, so two names differing
        only by case silently collide — guard against it at build time (D32).
        """
        for existing in self.names:
            if existing.lower() == name.lower() and existing != name:
                raise ValueError(
                    f"Defined-name case collision: {name!r} vs {existing!r} "
                    "(Excel names are case-insensitive)")
        if name in self.names:
            raise ValueError(f"Defined name {name!r} registered twice")
        full = f"{quote_sheet(sheet)}!{ref}"
        self.names[name] = (full, description)

    def flush_names(self):
        for name, (ref, _desc) in sorted(self.names.items()):
            self.wb.defined_names.add(DefinedName(name, attr_text=ref))

    @property
    def p(self) -> int:
        return self.cfg.plan_year

    @property
    def net_seed_bu(self) -> str:
        """The BU Net Delivery opens on: one that carries a net selection, so
        the tab demonstrates itself on a seeded book, falling back to the worked
        example for a carried-forward book with no net rows.

        Program Flow deliberately does NOT share this. The two tabs are the same
        book seen twice — prescriptive and descriptive — and it is tempting to
        open them on the same BU so clicking between them compares like with
        like. But each default is tied by verify_workbook phase C: Net Delivery's
        to the net showcase combo, Program Flow's to the worked example. Making
        them agree would mean either coupling an oracle tie to a cosmetic seed,
        or giving up the demo on one of them. Both tabs name the BU in view in
        bold beside the picker, which is where the answer belongs anyway.
        """
        net = next((r for r in self.lr_rows if r.get("netp") is not None), None)
        return (net or self.we_row)["bu"]


# Final tab order (approved redesign): orientation -> inputs -> the all-combo
# answers -> the selected-combo deep-dive chain -> what-ifs -> audit -> docs.
SHEET_ORDER = [
    SHEETS.README, SHEETS.CONTROL, SHEETS.INPUTS, SHEETS.RATE_LOG,
    SHEETS.MOD_LOG,
    SHEETS.PORTFOLIO, SHEETS.STATE_SUMMARY, SHEETS.PROGRAM_FLOW,
    SHEETS.NET_DELIVERY, SHEETS.BRIDGE,
    SHEETS.WALKTHROUGH, SHEETS.ONE_PAGER, SHEETS.FLOW, SHEETS.SCENARIOS,
    SHEETS.COMPARE, SHEETS.SOLVER, SHEETS.ATTRIBUTION, SHEETS.RATE_ENGINE,
    SHEETS.MOD_ENGINE, SHEETS.CHECKS, SHEETS.METHODOLOGY,
    SHEETS.LISTS, SHEETS.CALC, SHEETS.NETCALC, SHEETS.ORACLE,
]
HIDDEN_SHEETS = {SHEETS.LISTS, SHEETS.CALC, SHEETS.NETCALC, SHEETS.ORACLE}


# ---------------------------------------------------------------------------
# Assembly
# ---------------------------------------------------------------------------


def build(cfg: Config, lob_name: str, carried=None) -> Workbook:
    """Assemble one LOB workbook.

    ``carried`` (a carry.CarriedInputs) replaces the sample seeds with inputs
    read out of an existing workbook, so a rebuild keeps the user's book (D82).
    Everything downstream is unchanged: the seeds were only ever one source of
    the same row dicts, and the oracle constants are baked from whichever data
    is actually present — which is what finally makes the Checks oracle ties
    apply to a real book instead of reporting N/A forever.
    """
    from . import (sheets_briefs, sheets_calc, sheets_engine, sheets_inputs,
                   sheets_main, sheets_netdelivery, sheets_programflow,
                   sheets_modlog, sheets_ratelog, sheets_report,
                   sheets_walkthrough)

    Layout.configure(cfg)
    lob = cfg.lob(lob_name)
    if carried is not None:
        lr_rows = [dict(r) for r in carried.lr_rows]
        rate_rows = [dict(r) for r in carried.rate_rows]
        mod_rows = [dict(r) for r in carried.mod_rows]
        season_rows = [dict(r) for r in carried.season_rows]
        if not lr_rows:
            raise ValueError(
                f"carry-forward source has no populated tbl_LR rows: {carried.source}")
        for cap, rows, what in ((Layout.LR_ROWS, lr_rows, "tbl_LR"),
                                (Layout.RL_ROWS, rate_rows, "rate-log"),
                                (Layout.ML_ROWS, mod_rows, "mod-log"),
                                (Layout.SE_ROWS, season_rows, "seasonality")):
            if len(rows) > cap:
                raise ValueError(
                    f"carry-forward has {len(rows)} {what} rows but the configured "
                    f"capacity is {cap}; raise it in config/config.yaml and retry")
        if carried.term_months:
            lob = replace(lob, term_months=carried.term_months)
    else:
        lr_rows = sample_lr_rows(cfg, lob)
        rate_rows = sample_rate_rows(cfg, lob)
        mod_rows = sample_mod_rows(cfg, lob)
        season_rows = sample_seasonality_rows(cfg)

    # ---- oracle bake for the worked-example combo --------------------------
    # Seeded builds pin the documented §9 combo; a carried build takes the
    # first real row, so the baked constants describe the user's own book.
    we_row = None
    if carried is None:
        we_row = next((r for r in lr_rows
                       if r["bu"] == _bu(cfg, 0) and r["state"] == cfg.states[0]), None)
    if we_row is None:
        we_row = lr_rows[0]
    we_key = f"{we_row['bu']}|{we_row['state']}"
    combo = sample_to_combo(cfg, lob, we_row, rate_rows, mod_rows)
    if carried is None and cfg.plan_year == 2027 and lob.term_months == 12:
        # The seeded sample must reproduce §9 exactly for annual-term books
        # (comments/EP/s may differ, so compare results, not input dataclasses).
        p_ref, combo_ref = engine.worked_example()
        ref = engine.run_bridge(p_ref, combo_ref, "monthly")
        got = engine.run_bridge(cfg.plan_year, combo, "monthly")
        assert abs(ref.cy_lr_p - got.cy_lr_p) < 1e-12 and abs(ref.crl_ind - got.crl_ind) < 1e-12, (
            f"sample seed for {we_key} must reproduce engine.worked_example() results"
        )
    oracle_m = engine.run_bridge(cfg.plan_year, combo, "monthly")
    oracle_c = engine.run_bridge(cfg.plan_year, combo, "continuous")
    solver_res = engine.solve_rate_for_target(
        cfg.plan_year, combo, oracle_m.cy_lr_p,
        dt.date(cfg.plan_year, 4, 1),
        max_reasonable_rate=cfg.solver_max_rate,
        min_post_share=cfg.solver_min_post_share,
    )

    wb = Workbook()
    wb.remove(wb.active)
    for name in SHEET_ORDER:
        ws = wb.create_sheet(name)
        if name in HIDDEN_SHEETS:
            ws.sheet_state = "hidden"

    ctx = Ctx(
        cfg=cfg, lob=lob, wb=wb, lr_rows=lr_rows, rate_rows=rate_rows,
        mod_rows=mod_rows, season_rows=season_rows, oracle_m=oracle_m, oracle_c=oracle_c,
        oracle_solver_r=solver_res.required_change, we_key=we_key, we_row=we_row,
        carried=carried,
    )

    # Build order: reference data first, then engines, then presentation.
    sheets_inputs.build_lists(ctx)
    sheets_inputs.build_inputs(ctx)
    sheets_ratelog.build_rate_log(ctx)
    sheets_modlog.build_mod_log(ctx)
    sheets_inputs.build_control(ctx)
    sheets_engine.build_rate_engine(ctx)
    sheets_engine.build_mod_engine(ctx)
    sheets_main.build_bridge(ctx)
    sheets_calc.build_calc(ctx)
    sheets_main.build_portfolio(ctx)
    sheets_main.build_state_summary(ctx)
    sheets_netdelivery.build_netcalc(ctx)
    sheets_netdelivery.build_net_delivery(ctx)
    sheets_programflow.build_program_flow(ctx)
    sheets_main.build_scenarios(ctx)
    sheets_main.build_solver(ctx)
    sheets_main.build_attribution(ctx)
    sheets_report.build_flow_dashboard(ctx)
    sheets_walkthrough.build_walkthrough(ctx)
    sheets_briefs.build_one_pager(ctx)
    sheets_briefs.build_compare(ctx)
    sheets_report.build_oracle_sheet(ctx)
    sheets_report.build_checks(ctx)
    sheets_report.build_methodology(ctx)
    sheets_report.build_readme(ctx)

    ctx.flush_names()
    # D71: every chart is built by now, so sweep their titles out of the plot
    # area in one place — a per-chart fix is a thing the next chart forgets.
    from .xlstyle import assert_formulas_balanced, unoverlay_titles
    unoverlay_titles(wb)
    # D89: a formula that does not close its parentheses makes the whole file
    # unopenable, and Excel says so only from the recalculation step, minutes
    # later, without naming the cell. Catch it here, where the cell has a name.
    assert_formulas_balanced(wb)
    wb.calculation.calcMode = "auto"
    wb.calculation.fullCalcOnLoad = True
    return wb


def output_path(cfg: Config, out_dir: Path, lob_name: str) -> Path:
    safe = lob_name.replace(" ", "_")
    return out_dir / f"{cfg.filename.format(plan_year=cfg.plan_year, lob=safe)}.xlsx"


def backup_path(path: Path) -> Path:
    """Timestamped sibling of ``path`` that does not already exist."""
    stamp = dt.datetime.now().strftime("%Y%m%d-%H%M%S")
    cand = path.with_suffix(f".{stamp}.bak.xlsx")
    n = 1
    while cand.exists():
        cand = path.with_suffix(f".{stamp}-{n}.bak.xlsx")
        n += 1
    return cand


def main(argv=None) -> int:
    ap = argparse.ArgumentParser(
        description="Generate the CY Plan LR workbooks (one per LOB).",
        epilog="Regeneration REPLACES a workbook with a freshly built copy. It refuses to "
               "overwrite a file holding real (non-sample) inputs unless you pass --force "
               "or --carry-forward; either way a timestamped .bak.xlsx is written first.")
    ap.add_argument("--config", default="config/config.yaml")
    ap.add_argument("--out", default=None, help="output directory (default: from config)")
    ap.add_argument("--lob", default=None,
                    help="generate a single LOB workbook (default: all configured)")
    ap.add_argument("--carry-forward", nargs="?", const="auto", default=None,
                    metavar="SOURCE",
                    help="rebuild around the inputs already in SOURCE instead of the "
                         "sample seeds; with no value, carries each LOB forward from its "
                         "own existing file in the output directory")
    ap.add_argument("--force", action="store_true",
                    help="overwrite a workbook holding real inputs (a .bak is still kept)")
    ap.add_argument("--no-backup", action="store_true",
                    help="skip the .bak copy (not recommended)")
    args = ap.parse_args(argv)

    cfg = load_config(args.config)
    out_dir = Path(args.out or cfg.output_dir)
    out_dir.mkdir(parents=True, exist_ok=True)
    lob_names = [args.lob] if args.lob else list(cfg.output_lobs)
    for name in lob_names:
        if name not in {l.name for l in cfg.lobs}:
            raise SystemExit(f"unknown LOB {name!r}; configured: {[l.name for l in cfg.lobs]}")
    if args.carry_forward and args.carry_forward != "auto" and len(lob_names) > 1:
        raise SystemExit("--carry-forward SOURCE names one file; pass --lob too, or use "
                         "bare --carry-forward to carry each LOB from its own workbook.")

    from .carry import CarryError, holds_sample_data, read_inputs

    for name in lob_names:
        path = output_path(cfg, out_dir, name)
        carried = None
        if args.carry_forward:
            src = (path if args.carry_forward == "auto" else Path(args.carry_forward))
            if args.carry_forward == "auto" and not src.exists():
                print(f"  {name}: no existing workbook to carry forward — seeding samples")
            else:
                try:
                    carried = read_inputs(src)
                except CarryError as e:
                    raise SystemExit(f"carry-forward failed for {name}: {e}")
                if carried.missing:
                    print(f"  {name}: NOTE source lacks {', '.join(carried.missing)} "
                          "(older generator?) — those columns come up blank")
                print(f"  {name}: carrying forward {carried.summary()} from {src.name}")

        # The guard, and the reason this whole feature exists: a rebuild used to
        # silently replace a book somebody had been pasting into for months.
        # A file still holding the seeded sample has nothing to lose, so it is
        # overwritten quietly and without a .bak — backups mark real data.
        if path.exists() and not holds_sample_data(path, cfg, name):
            if carried is None and not args.force:
                raise SystemExit(
                    f"REFUSING to overwrite {path.name}: it holds real inputs, not the "
                    f"seeded sample.\n"
                    f"  Keep them:  python -m src.build_workbook --lob \"{name}\" "
                    f"--carry-forward\n"
                    f"  Discard:    add --force (a .bak copy is written either way)")
            if not args.no_backup:
                bak = backup_path(path)
                shutil.copy2(path, bak)
                print(f"  {name}: backed up existing workbook -> {bak.name}")

        wb = build(cfg, name, carried=carried)
        wb.save(path)
        print(f"wrote {path}")
    print(
        "NOTE: openpyxl stores formulas without cached results. Open each file once in "
        "Excel (or run tools/recalc.py) so values populate, then check the Checks sheet."
    )
    return 0


if __name__ == "__main__":
    # ``python -m src.build_workbook`` loads this file TWICE: as __main__ and
    # (via the builders' relative imports) as src.build_workbook — two separate
    # Layout classes. Delegate to the canonical module so configure() runs on
    # the SAME Layout object the sheet builders read; otherwise they silently
    # fall back to the class-body defaults (D46).
    from src.build_workbook import main as _canonical_main
    sys.exit(_canonical_main())
