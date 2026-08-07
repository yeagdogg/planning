"""Scenario validators — the workbook's Checks input rules at ENTRY time.

Same rules, same severities, different moment: the workbook discloses after
the fact on the Checks tab; the app flags while you type. FAIL means the
engine's answer is untrustworthy or the row is unusable; WARN is advisory
(the D107 target identity, considered-on-planned, window strays). UI-free
and tested on both interpreters.
"""
from __future__ import annotations

import datetime as dt


def _num(v):
    return v if isinstance(v, (int, float)) and not isinstance(v, bool) else None


def validate(scn) -> list:
    """[(severity, message)] — severity in {"FAIL", "WARN"}."""
    out = []
    p = scn.plan_year

    # ---- tbl_LR ------------------------------------------------------------
    seen = {}
    for r in scn.lr_rows:
        key = f"{r.get('bu')}|{r.get('state')}"
        seen[key] = seen.get(key, 0) + 1
    dupes = sorted(k for k, n in seen.items() if n > 1)
    if dupes:
        out.append(("FAIL", f"duplicate BU|State rows in tbl_LR: "
                            f"{', '.join(dupes)} — keys must be unique"))
    for r in scn.lr_rows:
        key = f"{r.get('bu')}|{r.get('state')}"
        ep = _num(r.get("ep"))
        if ep is not None and ep < 0:
            out.append(("FAIL", f"{key}: Plan EP is negative"))
        asof = r.get("m0_asof")
        if isinstance(asof, dt.date) and asof >= dt.date(p, 12, 31):
            out.append(("FAIL", f"{key}: M_0 as-of {asof} does not precede "
                                f"12/31/{p}"))
        for fld in ("netp", "netp1"):
            v = _num(r.get(fld))
            if v is not None and not (-0.5 <= v <= 1.0):
                out.append(("FAIL", f"{key}: {fld} {v:+.1%} outside "
                                    f"[-50%, +100%]"))
        mtok = r.get("modadj")
        if mtok not in (None, "ON", "OFF"):
            out.append(("WARN", f"{key}: mod-adj token {mtok!r} is not "
                                f"ON/OFF/blank (blank = ON)"))
        # D107 advisory: target = (combined - expense) / (ALAE x ULAE)
        t, c, e = (_num(r.get(k)) for k in ("target", "combined", "expense"))
        al, ul = _num(r.get("alae")), _num(r.get("ulae"))
        if None not in (t, c, e, al, ul) and al * ul != 0:
            if abs(t * al * ul - (c - e)) > 0.005 * al * ul:
                out.append(("WARN", f"{key}: Target LR {t:.1%} disagrees "
                                    f"with (Combined − Expense)/(ALAE×ULAE) "
                                    f"= {(c - e) / (al * ul):.1%} by more "
                                    f"than half a point"))

    # ---- rate / mod logs ---------------------------------------------------
    combos = set(scn.combo_keys())
    win_lo, win_hi = dt.date(p - 2, 1, 1), dt.date(p + 1, 12, 31)
    month_seen = {}
    for label, rows, pct_field in (("Rate Log", scn.rate_rows, "filed"),
                                   ("Mod Log", scn.mod_rows, "chg")):
        for r in rows:
            key = f"{r.get('bu')}|{r.get('state')}"
            if r.get("status") not in ("taken", "planned"):
                out.append(("FAIL", f"{label} {key}: status "
                                    f"{r.get('status')!r} is not "
                                    f"taken/planned"))
            if key not in combos:
                out.append(("WARN", f"{label} {key}: no matching tbl_LR "
                                    f"row — this change reaches nothing"))
            eff = r.get("eff")
            if isinstance(eff, dt.date) and not (win_lo <= eff <= win_hi):
                out.append(("WARN", f"{label} {key}: {eff} outside the "
                                    f"Jan {p - 2}–Dec {p + 1} engine window"))
            if _num(r.get(pct_field)) is None:
                out.append(("FAIL", f"{label} {key}: {pct_field} missing"))
            if (label == "Rate Log" and r.get("considered") == "Y"
                    and r.get("status") == "planned"):
                out.append(("WARN", f"{label} {key}: considered flag on a "
                                    f"planned row (§3.2.7)"))
            if label == "Rate Log" and isinstance(eff, dt.date):
                mk = (key, eff.year, eff.month)
                month_seen[mk] = month_seen.get(mk, 0) + 1
    for (key, y, m), n in sorted(month_seen.items()):
        if n > 1:
            out.append(("WARN", f"Rate Log {key}: {n} changes share cohort "
                                f"month {y}-{m:02d} (exact-blend support, "
                                f"D4)"))

    # ---- seasonality -------------------------------------------------------
    for r in scn.season_rows:
        w = r.get("weights") or []
        if r.get("state") and sum(w) <= 0:
            out.append(("WARN", f"Seasonality {r.get('state')}: weight sum "
                                f"is not positive — row is ignored"))

    return out
