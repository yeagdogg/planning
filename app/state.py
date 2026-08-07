"""The Scenario — the app's config object, and deliberately the same shape a
workbook rebuild carries (src/carry.py::CarriedInputs). That identity is the
whole on-ramp: an imported workbook, a saved scenario file, and the Excel
export are one vocabulary, and the row dicts feed
src/build_workbook.py::sample_to_combo unchanged.
"""
from __future__ import annotations

from dataclasses import dataclass, field
from typing import Optional


@dataclass
class Scenario:
    lob: str
    plan_year: int
    term_months: int
    lr_rows: list = field(default_factory=list)      # carry.py key vocabulary
    rate_rows: list = field(default_factory=list)
    mod_rows: list = field(default_factory=list)
    season_rows: list = field(default_factory=list)
    season_on: bool = False        # the workbook's global seasonality toggle
    mod_master: bool = True        # nr_ModAdjMaster — OFF pins every A_mod
    trend_default: float = 0.0     # nr_Trend fallback for blank row trends
    name: str = "untitled"
    source: str = ""               # provenance: workbook path / scenario file

    # ---- roster ------------------------------------------------------------
    def combo_keys(self) -> list:
        return [f"{r['bu']}|{r['state']}" for r in self.lr_rows
                if r.get("bu") and r.get("state")]

    def row(self, key: str) -> Optional[dict]:
        for r in self.lr_rows:
            if f"{r.get('bu')}|{r.get('state')}" == key:
                return r
        return None

    def ep(self, key: str) -> float:
        r = self.row(key)
        v = (r or {}).get("ep")
        return float(v) if isinstance(v, (int, float)) else 0.0
