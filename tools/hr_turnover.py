#!/usr/bin/env python3
"""
Data Challenge - HR / Employee Turnover domain pack
===================================================

Detects that an uploaded table is HR employee/turnover data (by column
signatures, not by filename) and runs turnover-specific checks that produce
data-driven findings with the key strategic questions an HR reviewer should
answer. Plugs into the same Result/Finding model as the generic engine.

Detection: needs an attrition/turnover status column plus at least two other
HR signature columns (employee id, department, job role, income/salary,
tenure, satisfaction...). Column names are matched against synonym lists, so
files from different HR systems still map.
"""
from __future__ import annotations

import pandas as pd

from data_challenge import Result

# --------------------------------------------------------------------------- #
# Column signatures (lowercase, matched on normalised header names)
# --------------------------------------------------------------------------- #
_SYNONYMS = {
    "attrition":   ["attrition", "turnover", "churn", "left", "quit", "resigned",
                    "terminated", "leaver", "exited", "status"],
    "employee_id": ["employeenumber", "employeeid", "employee_id", "emp_id",
                    "empid", "staffid", "worker_id"],
    "department":  ["department", "dept", "division", "area", "business_unit"],
    "job_role":    ["jobrole", "job_role", "role", "position", "job_title", "title"],
    "income":      ["monthlyincome", "salary", "monthly_income", "annual_salary",
                    "base_pay", "compensation", "wage"],
    "tenure":      ["yearsatcompany", "tenure", "years_at_company", "seniority",
                    "service_years", "yearsofservice"],
    "overtime":    ["overtime", "over_time", "extra_hours"],
    "rating":      ["performancerating", "performance_rating", "rating",
                    "last_evaluation", "performance_score"],
    "satisfaction": ["jobsatisfaction", "job_satisfaction", "satisfaction",
                     "satisfaction_level", "engagement"],
    "promotion":   ["yearssincelastpromotion", "years_since_last_promotion",
                    "last_promotion", "promotion_gap"],
    "manager":     ["yearswithcurrmanager", "years_with_manager", "manager_tenure"],
    "hire_date":   ["hire_date", "hiredate", "start_date", "joining_date"],
    "term_date":   ["termination_date", "terminationdate", "end_date", "exit_date",
                    "leaving_date", "fecha_baja"],
}

_LEFT_VALUES = {"yes", "1", "true", "left", "terminated", "resigned", "quit", "y"}


def _norm(name: str) -> str:
    return str(name).lower().replace(" ", "").replace("-", "").replace("_", "")


def _find(df: pd.DataFrame, key: str) -> str | None:
    wanted = [_norm(s) for s in _SYNONYMS[key]]
    for c in df.columns:
        if _norm(c) in wanted:
            return c
    return None


def detect_hr(df: pd.DataFrame) -> dict | None:
    """Return a role->column mapping if this looks like HR turnover data."""
    m = {k: _find(df, k) for k in _SYNONYMS}
    if not m["attrition"]:
        return None
    others = sum(1 for k, v in m.items() if v and k != "attrition")
    if others < 2:
        return None
    # the attrition column must actually contain leaver/stayer style values
    vals = df[m["attrition"]].dropna().astype(str).str.strip().str.lower().unique()
    if not any(v in _LEFT_VALUES for v in vals):
        return None
    return m


def _left_mask(df: pd.DataFrame, col: str) -> pd.Series:
    return df[col].astype(str).str.strip().str.lower().isin(_LEFT_VALUES)


def _rate(mask: pd.Series) -> float:
    return float(mask.mean()) if len(mask) else 0.0


# --------------------------------------------------------------------------- #
# Checks — each one computes real numbers and asks the key question
# --------------------------------------------------------------------------- #
def run_hr_checks(df: pd.DataFrame, m: dict) -> Result:
    r = Result()
    left = _left_mask(df, m["attrition"])
    n, n_left = len(df), int(left.sum())
    rate = _rate(left)

    r.profile = {"employees": n, "leavers": n_left,
                 "turnover_rate": f"{rate:.1%}"}
    if m["department"]:
        r.profile["departments"] = int(df[m["department"]].nunique())

    # 1) Overall turnover level
    r.add("Turnover level",
          f"{n_left} of {n} employees ({rate:.1%}) have left. A common healthy "
          f"benchmark is ~10%; above 15% the cost of replacement (recruiting, "
          f"onboarding, lost knowledge) starts compounding.",
          f"{rate:.1%}",
          "Is this level of turnover planned (restructuring, seasonal staff) or "
          "unplanned? What does it cost per replacement, and which single "
          "department would you fix first if you could only fix one?",
          "High" if rate > 0.15 else "Medium")

    # 2) Overtime as a driver
    if m["overtime"]:
        ot = df[m["overtime"]].astype(str).str.strip().str.lower().isin(
            {"yes", "1", "true", "y"})
        if ot.any() and (~ot).any():
            r_ot, r_no = _rate(left[ot]), _rate(left[~ot])
            if r_no > 0 and r_ot / max(r_no, 1e-9) >= 1.5:
                r.add("Overtime and burnout",
                      f"Employees doing overtime leave at {r_ot:.1%} vs {r_no:.1%} "
                      f"for those who don't - {r_ot / r_no:.1f}x higher.",
                      f"{r_ot:.0%} vs {r_no:.0%}",
                      "Overtime looks like a burnout signal, not a productivity one. "
                      "Is overtime concentrated in specific teams or managers, and is "
                      "it structural (understaffing) or seasonal? What would one extra "
                      "hire per affected team cost vs the attrition it prevents?",
                      "High")

    # 3) Early-tenure attrition
    if m["tenure"]:
        ten = pd.to_numeric(df[m["tenure"]], errors="coerce")
        early = ten <= 2
        if n_left > 0 and early.notna().any():
            share_early = float((left & early).sum()) / n_left
            if share_early >= 0.35:
                r.add("Early leavers",
                      f"{share_early:.0%} of all leavers had 2 years or less at the "
                      f"company.",
                      f"{share_early:.0%} <= 2 yrs",
                      "People leaving this early usually signals hiring or onboarding "
                      "problems, not career fatigue: are expectations set in recruiting "
                      "matching the real job? Is there a structured first-year "
                      "programme, and does anyone own new-hire retention as a KPI?",
                      "High")

    # 4) Are we losing the good ones? (regrettable attrition)
    if m["rating"]:
        rat = pd.to_numeric(df[m["rating"]], errors="coerce")
        if rat.notna().any():
            top = rat >= rat.max()
            r_top, r_rest = _rate(left[top]), _rate(left[~top])
            if r_top >= r_rest and top.sum() >= 20:
                r.add("Regrettable attrition",
                      f"Top-rated employees leave at {r_top:.1%}, the rest at "
                      f"{r_rest:.1%}. Losing average performers is turnover; losing "
                      f"the best is talent drain.",
                      f"{r_top:.0%} top vs {r_rest:.0%}",
                      "Which of the leavers would you have fought to keep? Is there a "
                      "retention plan (pay, scope, promotion) for the top quartile, "
                      "and does their manager know they are a flight risk before the "
                      "resignation letter arrives?",
                      "High")

    # 5) Pay gap between leavers and stayers
    if m["income"]:
        inc = pd.to_numeric(df[m["income"]], errors="coerce")
        med_l = float(inc[left].median()) if n_left else 0.0
        med_s = float(inc[~left].median()) if (~left).any() else 0.0
        if med_s > 0 and med_l < med_s * 0.85:
            gap = 1 - med_l / med_s
            r.add("Compensation gap",
                  f"Median pay of leavers is {med_l:,.0f} vs {med_s:,.0f} for those "
                  f"who stay - {gap:.0%} lower.",
                  f"-{gap:.0%} median pay",
                  "Are people leaving because they are underpaid, or are the "
                  "lower-paid roles simply the revolving-door ones? Benchmark those "
                  "roles against market: is the raise needed to retain cheaper than "
                  "the cost to replace?",
                  "Medium")

    # 6) Hotspots: department / role with the worst rate
    for key, label in (("department", "department"), ("job_role", "role")):
        col = m[key]
        if not col:
            continue
        grp = df.groupby(col).agg(n=(col, "size"))
        grp["rate"] = df.groupby(col)[m["attrition"]].apply(
            lambda s: _rate(s.astype(str).str.strip().str.lower().isin(_LEFT_VALUES)))
        grp = grp[grp["n"] >= 20]
        if len(grp) >= 2:
            worst = grp["rate"].idxmax()
            w_rate, w_n = grp.loc[worst, "rate"], int(grp.loc[worst, "n"])
            if rate > 0 and w_rate / rate >= 1.5:
                r.add(f"Hotspot {label}: {worst}",
                      f"'{worst}' loses {w_rate:.1%} of its {w_n} people vs "
                      f"{rate:.1%} company-wide - {w_rate / rate:.1f}x the average.",
                      f"{w_rate:.0%} vs {rate:.0%}",
                      f"What is different in '{worst}': the manager, the workload, "
                      f"the pay band, or the nature of the job itself? Have exit "
                      f"interviews been read side by side for this group, and what do "
                      f"they repeat?",
                      "High")

    # 7) Promotion stagnation
    if m["promotion"]:
        promo = pd.to_numeric(df[m["promotion"]], errors="coerce")
        if n_left and promo.notna().any():
            stuck = promo >= 4
            if stuck.any():
                r_stuck, r_move = _rate(left[stuck]), _rate(left[~stuck])
                if r_stuck > r_move * 1.3:
                    r.add("Promotion stagnation",
                          f"Employees 4+ years without a promotion leave at "
                          f"{r_stuck:.1%} vs {r_move:.1%} for the rest.",
                          f"{r_stuck:.0%} vs {r_move:.0%}",
                          "Is there a real career path for senior-in-role people, or "
                          "is promotion the only progression currency? Would lateral "
                          "moves, scope growth or pay-in-band reviews retain them?",
                          "Medium")

    # 8) Missing the WHEN: no termination date
    if not m["term_date"]:
        r.add("No exit dates",
              "The file says WHO left but not WHEN: there is no termination/exit "
              "date column, so seasonality, cohort analysis and trend-over-time "
              "are impossible.",
              "no date column",
              "Can HR export exit dates? Without them you cannot tell whether "
              "turnover is improving or worsening, or whether leavers cluster "
              "after bonus payouts, reorganisations or manager changes.",
              "Medium")

    return r
