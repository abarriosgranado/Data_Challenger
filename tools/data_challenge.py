#!/usr/bin/env python3
"""
Data Challenge Machine
======================

A reusable FP&A "challenge the data" tool. Point it at a CSV or Excel file of
transactional / expense data and it will:

  1. Auto-detect the Date, Amount, Category and dimension columns (overridable).
  2. Run a battery of data-quality + FP&A assumption checks.
  3. Emit a styled Excel report with a findings table (severity, evidence,
     question to ask) plus a data-profile summary.
  4. Optionally compare the file against a *baseline* (a prior version) and
     challenge what CHANGED — new categories, mix shifts, month-over-month
     swings, totals moving materially.

It is deliberately generic: it works on any tabular spend/expense file, not just
the travel dataset it was first built for.

Usage
-----
    python tools/data_challenge.py INPUT.csv
    python tools/data_challenge.py INPUT.xlsx --sheet "Raw Data"
    python tools/data_challenge.py NEW.csv --baseline OLD.csv
    python tools/data_challenge.py INPUT.csv --out my_report.xlsx \
        --date-col Date --amount-col "Amount (EUR)" --category-col "Expense Type"

Dependencies: pandas, openpyxl (both preinstalled in the Numerus env).
"""
from __future__ import annotations

import argparse
import os
import sys
from dataclasses import dataclass, field
from datetime import datetime

import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


# --------------------------------------------------------------------------- #
# Configuration thresholds (tune here)
# --------------------------------------------------------------------------- #
@dataclass
class Config:
    round_share_warn: float = 0.60        # >60% whole-number amounts => estimates?
    outlier_z: float = 3.0                # per-category z-score for outliers
    outlier_min_n: int = 5                # min category size to test outliers
    concentration_top_n: int = 3          # top-N categories concentration
    concentration_warn: float = 0.70      # >70% in top-N => concentration risk
    partial_period_ratio: float = 0.60    # last period < 60% of median => partial
    material_change: float = 0.15         # >15% total swing vs baseline is material
    new_category_flag: bool = True
    high_value_pctile: float = 0.99       # flag single txns above this percentile
    font: str = "Arial"


# --------------------------------------------------------------------------- #
# Finding model
# --------------------------------------------------------------------------- #
@dataclass
class Finding:
    area: str
    observation: str
    evidence: str
    question: str
    severity: str  # High / Medium / Low


@dataclass
class Result:
    findings: list = field(default_factory=list)
    profile: dict = field(default_factory=dict)
    changes: list = field(default_factory=list)  # for baseline mode

    def add(self, area, observation, evidence, question, severity):
        self.findings.append(Finding(area, observation, evidence, question, severity))


# --------------------------------------------------------------------------- #
# Loading & column detection
# --------------------------------------------------------------------------- #
def load_table(path: str, sheet: str | None = None) -> pd.DataFrame:
    ext = os.path.splitext(path)[1].lower()
    if ext in (".xlsx", ".xlsm", ".xls"):
        df = pd.read_excel(path, sheet_name=sheet or 0)
    else:
        # robust CSV: sniff separator, tolerate latin-1
        for sep in (None, ";", ",", "\t"):
            try:
                df = pd.read_csv(path, sep=sep, engine="python", encoding="latin-1")
                if df.shape[1] > 1:
                    break
            except Exception:
                continue
        else:
            raise ValueError(f"Could not parse {path}")
    # drop fully empty rows/cols
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.columns = [str(c).strip() for c in df.columns]
    return df


def parse_dates(series: pd.Series) -> pd.Series:
    """Robustly parse a column to dates.

    Pure-numeric columns (Year, Amount) are rejected so they are not mistaken
    for dates. Otherwise try day-first and month-first parsing and keep whichever
    resolves more values, so both DD/MM/YYYY and ISO/US files work.
    """
    numeric_share = pd.to_numeric(series, errors="coerce").notna().mean()
    non_numeric = 1 - numeric_share
    # a column that is almost entirely bare numbers is not a date column
    if non_numeric < 0.5:
        return pd.Series(pd.NaT, index=series.index)
    a = pd.to_datetime(series, errors="coerce", dayfirst=False)
    b = pd.to_datetime(series, errors="coerce", dayfirst=True)
    return a if a.notna().sum() >= b.notna().sum() else b


def _score_date(series: pd.Series) -> float:
    return parse_dates(series).notna().mean()


def _date_distinctness(series: pd.Series) -> float:
    """Fraction of distinct parsed dates. A real transaction date has many;
    a derived 'Month'/'Year' column has very few, which lets us break ties."""
    valid = parse_dates(series).dropna()
    if len(valid) == 0:
        return 0.0
    return valid.nunique() / len(valid)


def _score_numeric(series: pd.Series) -> float:
    coerced = pd.to_numeric(
        series.astype(str).str.replace(",", ".", regex=False).str.replace(r"[^\d.\-]", "", regex=True),
        errors="coerce",
    )
    return coerced.notna().mean()


def detect_columns(df: pd.DataFrame, overrides: dict) -> dict:
    cols = list(df.columns)
    lc = {c: c.lower() for c in cols}
    picked = {"date": overrides.get("date"), "amount": overrides.get("amount"),
              "category": overrides.get("category")}

    # date
    if not picked["date"]:
        best, best_s = None, 0.5
        for c in cols:
            s = _score_date(df[c])
            # exact/strong name match beats derived period columns
            if lc[c] in ("date", "transaction date", "posting date"):
                hint = 0.4
            elif any(k in lc[c] for k in ("date", "day")):
                hint = 0.2
            elif any(k in lc[c] for k in ("period", "month", "year")):
                hint = 0.05
            else:
                hint = 0.0
            # reward genuine high-cardinality dates over derived buckets
            distinct = _date_distinctness(df[c]) if s > 0.5 else 0.0
            score = s + hint + 0.25 * distinct
            if score > best_s:
                best, best_s = c, score
        picked["date"] = best
    # amount
    if not picked["amount"]:
        best, best_s = None, 0.5
        for c in cols:
            if c == picked["date"]:
                continue
            s = _score_numeric(df[c])
            # measure-like names (money or durations/quantities)
            if any(k in lc[c] for k in
                   ("amount", "cost", "value", "price", "total", "spend",
                    "eur", "usd", "gbp", "expense", "time", "duration",
                    "minutes", "hours", "qty", "quantity", "distance")):
                hint = 0.25
            else:
                hint = 0.0
            # identifier / coordinate / attribute columns are NOT measures
            if any(k in lc[c] for k in
                   ("id", "code", "lat", "lon", "zip", "postal", "phone",
                    "year", "age", "rating")):
                hint -= 0.5
            # an all-unique coerced column is an identifier, not a measure
            if s > 0.9 and df[c].nunique(dropna=True) / max(len(df), 1) > 0.95 \
               and not str(df[c].dtype).startswith("float"):
                hint -= 0.5
            # prefer columns that are natively numeric in the file
            if str(df[c].dtype) in ("int64", "float64", "Int64", "Float64"):
                hint += 0.1
            if s + hint > best_s:
                best, best_s = c, s + hint
        picked["amount"] = best
    # category
    if not picked["category"]:
        cand = [c for c in cols if c not in (picked["date"], picked["amount"])]
        # prefer a name hint, else the low-cardinality text column
        named = [c for c in cand if any(k in lc[c] for k in
                 ("categor", "type", "class", "account", "group"))]
        if named:
            picked["category"] = named[0]
        elif cand:
            def card(c):
                n = df[c].nunique(dropna=True)
                return n if 1 < n <= max(30, len(df) // 5) else 10 ** 9
            picked["category"] = min(cand, key=card)
    # dimensions = everything else non-numeric-ish
    dims = [c for c in cols if c not in picked.values() and c is not None]
    picked["dimensions"] = dims
    return picked


def normalize(df: pd.DataFrame, cols: dict) -> pd.DataFrame:
    out = pd.DataFrame()
    out.attrs["amount_name"] = str(cols.get("amount") or "value").replace("_", " ")
    if cols["date"]:
        out["date"] = parse_dates(df[cols["date"]])
    if cols["amount"]:
        out["amount"] = pd.to_numeric(
            df[cols["amount"]].astype(str).str.replace(",", ".", regex=False)
            .str.replace(r"[^\d.\-]", "", regex=True), errors="coerce")
    if cols["category"]:
        out["category"] = df[cols["category"]].astype(str).str.strip()
    for d in cols["dimensions"]:
        out[f"dim::{d}"] = df[d].astype(str).str.strip()
    dist = _derive_distance(df)
    if dist is not None and dist.notna().mean() > 0.5:
        out["distance_km"] = dist
    out = _drop_total_rows(out)
    # remember whether the measure is money or a duration/quantity, so checks
    # that only make sense for money (e.g. round-number = estimates) can adapt
    amt_name = (cols.get("amount") or "").lower()
    if any(k in amt_name for k in ("time", "duration", "minute", "hour", "day")):
        out.attrs["measure_kind"] = "duration"
    elif any(k in amt_name for k in ("qty", "quantity", "count", "units")):
        out.attrs["measure_kind"] = "quantity"
    else:
        out.attrs["measure_kind"] = "money"
    return out


def _haversine_km(lat1, lon1, lat2, lon2):
    """Great-circle distance in km between two coordinate series."""
    import numpy as np
    rad = np.pi / 180.0
    dlat = (lat2 - lat1) * rad
    dlon = (lon2 - lon1) * rad
    a = (np.sin(dlat / 2) ** 2
         + np.cos(lat1 * rad) * np.cos(lat2 * rad) * np.sin(dlon / 2) ** 2)
    return 2 * 6371.0 * np.arcsin(np.sqrt(a))


_ORIGIN_HINTS = ("store", "origin", "pickup", "restaurant", "start", "from", "source")


def _derive_distance(df: pd.DataFrame) -> pd.Series | None:
    """If the raw data carries two lat/lon pairs (origin + destination),
    compute the trip distance so findings can answer their own questions."""
    lat_cols = [c for c in df.columns if "lat" in str(c).lower()]
    lon_cols = [c for c in df.columns if "lon" in str(c).lower() or "lng" in str(c).lower()]
    if len(lat_cols) != 2 or len(lon_cols) != 2:
        return None

    def is_origin(name):
        return any(h in str(name).lower() for h in _ORIGIN_HINTS)

    lat_o = next((c for c in lat_cols if is_origin(c)), lat_cols[0])
    lat_d = next(c for c in lat_cols if c != lat_o)
    lon_o = next((c for c in lon_cols if is_origin(c)), lon_cols[0])
    lon_d = next(c for c in lon_cols if c != lon_o)

    vals = {}
    for name, c in (("lat_o", lat_o), ("lon_o", lon_o), ("lat_d", lat_d), ("lon_d", lon_d)):
        vals[name] = pd.to_numeric(df[c], errors="coerce")
    ok = (vals["lat_o"].abs().le(90) & vals["lat_d"].abs().le(90)
          & vals["lon_o"].abs().le(180) & vals["lon_d"].abs().le(180)
          & vals["lat_o"].abs().gt(0.1) & vals["lat_d"].abs().gt(0.1))
    dist = pd.Series(_haversine_km(vals["lat_o"], vals["lon_o"], vals["lat_d"], vals["lon_d"]),
                     index=df.index, dtype="float64")
    dist[~ok] = float("nan")
    dist[dist > 500] = float("nan")   # a same-day trip beyond 500 km is coord noise
    return dist


def _row_label(t: pd.DataFrame, idx) -> str:
    """Describe one row the way a person would point at it:
    'Order ID ialx566343618 on 2022-03-19 (Electronics)'."""
    row = t.loc[idx]
    parts = []
    for c in t.columns:                      # an identifier, if the file has one
        if c.startswith("dim::") and _is_id_like(c[5:]):
            parts.append(f"{c[5:].replace('_', ' ')} {row[c]}")
            break
    if "date" in t.columns and pd.notna(row.get("date")):
        parts.append(f"on {row['date'].date()}")
    best, desc = 10, None                    # the most descriptive text column
    for c in t.columns:
        if (c.startswith("dim::") and not _is_id_like(c[5:])
                and "time" not in c.lower()):     # a bare clock time reads poorly
            vals = t[c].dropna().astype(str)
            numeric_share = pd.to_numeric(vals, errors="coerce").notna().mean() if len(vals) else 1
            if numeric_share > 0.5:               # numbers make poor descriptions
                continue
            nu = t[c].nunique()
            if nu > best:
                best, desc = nu, c
    if desc is not None and str(row[desc]) not in ("nan", "None", ""):
        parts.append(f"('{row[desc]}')")
    return " ".join(str(p) for p in parts) if parts else "one record"


_TOTAL_WORDS = {"total", "totals", "subtotal", "sub-total", "grand total",
                "sum", "sum:", "grand-total"}


def _drop_total_rows(out: pd.DataFrame) -> pd.DataFrame:
    """Remove trailing 'TOTAL'/'Subtotal' summary rows that files often carry:
    either a row whose text cells say 'total', or a row that has an amount but
    no category and no descriptive dimensions."""
    if len(out) == 0:
        return out
    text_cols = [c for c in out.columns if c == "category" or c.startswith("dim::")]
    mask_keep = pd.Series(True, index=out.index)
    for c in text_cols:
        vals = out[c].astype(str).str.strip().str.lower()
        mask_keep &= ~vals.isin(_TOTAL_WORDS)
    if "amount" in out.columns and text_cols:
        blank = pd.Series(True, index=out.index)
        for c in text_cols:
            v = out[c].astype(str).str.strip().str.lower()
            blank &= v.isin(["", "nan", "none"])
        mask_keep &= ~(blank & out["amount"].notna())
    return out[mask_keep].reset_index(drop=True)


# --------------------------------------------------------------------------- #
# Checks
# --------------------------------------------------------------------------- #
def run_checks(t: pd.DataFrame, cfg: Config) -> Result:
    r = Result()
    n = len(t)
    has_amt = "amount" in t.columns
    has_date = "date" in t.columns
    has_cat = "category" in t.columns
    kind = t.attrs.get("measure_kind", "money")
    amt_name = str(t.attrs.get("amount_name", "value")).replace("_", " ")

    total = float(t["amount"].sum(skipna=True)) if has_amt else None
    r.profile.update({
        "rows": n,
        "total_amount": round(total, 2) if total is not None else "n/a",
        "date_min": str(t["date"].min().date()) if has_date and t["date"].notna().any() else "n/a",
        "date_max": str(t["date"].max().date()) if has_date and t["date"].notna().any() else "n/a",
        "categories": int(t["category"].nunique()) if has_cat else "n/a",
    })

    # --- Completeness: blanks / unparsed ---
    if has_amt:
        missing_amt = int(t["amount"].isna().sum())
        if missing_amt:
            r.add("Completeness",
                  f"{missing_amt} of {n} rows have no {amt_name} recorded — those records silently disappear from every total and average.",
                  f"{missing_amt} of {n} rows",
                  f"Which system or export step is losing the {amt_name} values, and who is responsible for fixing it at the source?",
                  "High")
    if has_date:
        missing_date = int(t["date"].isna().sum())
        if missing_date:
            r.add("Completeness",
                  f"{missing_date} rows have no valid date and are excluded from any trend or period view.",
                  f"{missing_date} of {n} rows",
                  "Should records without a date be parked in a 'to clarify' bucket until the source assigns one, instead of quietly distorting the monthly picture?",
                  "High")

    # --- Reliability: round-number share (estimate detection; money only —
    #     whole numbers are normal for durations/quantities) ---
    if has_amt and t.attrs.get("measure_kind", "money") == "money":
        amt = t["amount"].dropna()
        if len(amt):
            round_share = float((amt == amt.round(0)).mean())
            if round_share >= cfg.round_share_warn:
                r.add("Data reliability",
                      f"{round_share:.0%} of the amounts are perfectly round numbers — that is what hand-typed estimates look like, not real receipts.",
                      f"{round_share:.0%} round",
                      "Are people typing estimates instead of submitting receipts? If so, should receipts be required above a certain amount?",
                      "High")

    # --- Zero / negative amounts ---
    if has_amt:
        zero = int((t["amount"] == 0).sum())
        neg = int((t["amount"] < 0).sum())
        if zero:
            zrows = t[t["amount"] == 0]
            zdesc = ""
            dcols = [c for c in zrows.columns if c.startswith("dim::")]
            if len(zrows) and dcols:
                # pick the most descriptive dimension (highest cardinality = free text)
                desc_col = max(dcols, key=lambda c: t[c].nunique(dropna=True))
                zdesc = f" (e.g. '{zrows.iloc[0][desc_col]}')"
            if kind == "duration":
                r.add("Zero times",
                      f"{zero} record(s) show a {amt_name} of 0{zdesc} — a time of zero is physically impossible and points to a capture failure.",
                      f"{zero} at 0",
                      "What breaks when these zeros get created — the device, the app or the process? Should the system refuse to save a zero time?",
                      "Medium")
            else:
                r.add("Zero-value lines",
                      f"{zero} line(s) recorded at exactly 0{zdesc}.",
                      f"{zero} at 0",
                      "Are these genuinely free, or lost amounts? Should a zero only be allowed when someone marks it as 'free' on purpose?",
                      "Low")
        if neg:
            if kind == "duration":
                r.add("Negative times",
                      f"{neg} record(s) show a negative {amt_name}, which is impossible.",
                      f"{neg} negative",
                      "How can a negative time be recorded at all? This looks like a bug in the capture system worth fixing at the source.",
                      "High")
            else:
                r.add("Credits / negatives",
                      f"{neg} negative line(s) detected — treated as credits/refunds in the totals.",
                      f"{neg} negative",
                      "Are these real refunds that should offset the original cost, or typing errors? Someone should confirm each one.",
                      "Medium")

    # --- Duplicates ---
    if has_amt:
        key_cols = [c for c in ("date", "amount", "category") if c in t.columns]
        desc_dims = [c for c in t.columns if c.startswith("dim::")]
        dup_key = key_cols + desc_dims[:1]
        dup_mask = t.duplicated(subset=dup_key, keep=False)
        dups = int(dup_mask.sum())
        if dups:
            # cite a concrete example so the reviewer sees WHAT is duplicated
            ex = t[dup_mask].sort_values(dup_key[0] if dup_key else t.columns[0]).head(2)
            ex_txt = ""
            if len(ex):
                row = ex.iloc[0]
                parts = []
                if "date" in ex.columns and pd.notna(row.get("date")):
                    parts.append(str(row["date"].date()))
                if "category" in ex.columns:
                    parts.append(str(row["category"]))
                if "amount" in ex.columns and pd.notna(row.get("amount")):
                    parts.append(f"{row['amount']:g}")
                ex_txt = f" Example: {' / '.join(parts)} appears more than once."
            if kind == "money":
                dq = ("Could any of these have been paid twice? What stops the same expense being "
                      "submitted twice, and how would we get double payments back?")
            else:
                dq = ("Are these repeats created by the system (double clicks, retries, a faulty "
                      "export)? They inflate counts and distort averages until they are removed.")
            r.add("Duplicates detected",
                  f"{dups} rows are exact copies of another row.{ex_txt}",
                  f"{dups} rows",
                  dq,
                  "Medium")

    # --- Category hygiene: casing / whitespace collisions ---
    if has_cat:
        raw = t["category"].dropna()
        norm_map = {}
        for v in raw.unique():
            k = v.strip().lower()
            norm_map.setdefault(k, set()).add(v)
        collisions = {k: v for k, v in norm_map.items() if len(v) > 1}
        if collisions:
            example = "; ".join(" / ".join(sorted(v)) for v in list(collisions.values())[:3])
            r.add("Inconsistent categories",
                  f"{len(collisions)} category label(s) differ only by case/spacing (e.g. {example}) — totals are silently split across variants.",
                  f"{len(collisions)} collisions",
                  "Who owns the category taxonomy, and can the source system enforce a picklist so free-text variants cannot occur?",
                  "Medium")

    # --- Concentration (money) / speed gap between categories (durations) ---
    if has_amt and has_cat and total and kind == "money":
        by_cat = t.groupby("category")["amount"].sum().sort_values(ascending=False)
        top = by_cat.head(cfg.concentration_top_n)
        share = float(top.sum() / total)
        if share >= cfg.concentration_warn:
            names = ", ".join(top.index[:cfg.concentration_top_n])
            r.add("Where the money goes",
                  f"Just {cfg.concentration_top_n} categories account for {share:.0%} of everything spent ({names}).",
                  f"{share:.0%} in top {cfg.concentration_top_n}",
                  f"If costs need to come down, this is where the money actually is. Has anyone negotiated better rates for {names}? Saving 5% there is worth more than cutting every small category to zero.",
                  "Medium")
    elif has_amt and has_cat and kind == "duration":
        stats = t.dropna(subset=["amount"]).groupby("category")["amount"].agg(["mean", "count"])
        stats = stats[stats["count"] >= 30]
        if len(stats) >= 2:
            slow, fast = stats["mean"].idxmax(), stats["mean"].idxmin()
            ms, mf = float(stats.loc[slow, "mean"]), float(stats.loc[fast, "mean"])
            if mf > 0 and ms / mf >= 1.25:
                obs = f"On average, '{slow}' takes {ms:.0f} while '{fast}' takes {mf:.0f} — {ms / mf - 1:.0%} longer."
                q = (f"Why does '{slow}' take so much longer than '{fast}'? Is it the nature of the work, "
                     f"or are those cases getting less priority, fewer people or worse routing?")
                # before asking a human, try to answer it with the data itself:
                # do the slower category's trips simply cover more distance?
                if "distance_km" in t.columns:
                    dmed = t.dropna(subset=["amount", "distance_km"]).groupby("category")["distance_km"].median()
                    if slow in dmed.index and fast in dmed.index and dmed[fast] > 0:
                        ds, dfst = float(dmed[slow]), float(dmed[fast])
                        time_ratio, dist_ratio = ms / mf, ds / dfst
                        if dist_ratio >= 0.7 * time_ratio:
                            obs += (f" The data largely answers this itself: '{slow}' trips cover a median of "
                                    f"{ds:.1f} km versus {dfst:.1f} km for '{fast}' — the gap is mostly distance, "
                                    f"not performance.")
                            q = (f"Since distance explains most of the '{slow}' vs '{fast}' gap, should delivery "
                                 f"targets be set per distance band instead of per category, so long-haul "
                                 f"categories are not unfairly compared with short-hop ones?")
                        elif dist_ratio <= 1.3:
                            obs += (f" And it is NOT distance: both cover similar ground "
                                    f"({ds:.1f} km vs {dfst:.1f} km median per trip).")
                            q = (f"Distance is ruled out — '{slow}' and '{fast}' travel similar distances, yet "
                                 f"'{slow}' takes {ms / mf - 1:.0%} longer. So what is different: courier priority, "
                                 f"preparation time before pickup, or routing?")
                        else:
                            obs += (f" Distance explains part of it ('{slow}' trips: {ds:.1f} km median vs "
                                    f"{dfst:.1f} km), but not all.")
                            q = (f"After allowing for the longer distances, '{slow}' is still slower than "
                                 f"'{fast}'. What explains the remainder — priority, prep time, or routing?")
                r.add("Some categories are much slower", obs,
                      f"{ms:.0f} vs {mf:.0f}", q, "Medium")

    # --- Per-category outliers: point at the concrete case, like a human would ---
    if has_amt and has_cat:
        for cat, grp in t.dropna(subset=["amount"]).groupby("category"):
            if len(grp) < cfg.outlier_min_n:
                continue
            mu, sd = grp["amount"].mean(), grp["amount"].std(ddof=0)
            med = float(grp["amount"].median())
            if sd and sd > 0 and med > 0:
                hi = grp[(grp["amount"] - mu) / sd >= cfg.outlier_z]
                if len(hi):
                    worst_idx = hi["amount"].idxmax()
                    val = float(t.loc[worst_idx, "amount"])
                    ratio = val / med
                    label = _row_label(t, worst_idx)
                    more = f" {len(hi) - 1} more case(s) like this were found." if len(hi) > 1 else ""
                    dist_note = ""
                    if kind == "duration" and "distance_km" in t.columns:
                        d_case = t.loc[worst_idx, "distance_km"]
                        d_med = grp["distance_km"].median() if "distance_km" in grp.columns else None
                        if pd.notna(d_case) and d_med and pd.notna(d_med) and d_med > 0:
                            if d_case >= 1.5 * d_med:
                                dist_note = (f" That trip was also unusually long ({d_case:.1f} km vs a typical "
                                             f"{d_med:.1f} km), which may explain part of it.")
                            else:
                                dist_note = (f" Distance does not explain it: the trip was {d_case:.1f} km, "
                                             f"close to the typical {d_med:.1f} km.")
                    if kind == "duration":
                        obs = (f"A '{cat}' case normally takes about {med:g}, but {label} "
                               f"took {val:g} — {ratio:.1f}x the usual.{more}{dist_note}")
                        q = (f"What happened with that specific case? And should an alert fire the moment a "
                             f"'{cat}' {amt_name} passes double the usual, so the team can react the same day "
                             f"instead of discovering it in a report?")
                    else:
                        obs = (f"A typical '{cat}' line is about {med:g}, but {label} "
                               f"came to {val:g} — {ratio:.1f}x the norm.{more}")
                        q = (f"Was that specific item approved before the money was spent? Should anything "
                             f"above twice the usual '{cat}' cost require sign-off first?")
                    r.add(f"Out of the ordinary: {cat}", obs,
                          f"{val:g} vs usual {med:g}", q, "Medium")

    # --- High-value single transactions ---
    if has_amt:
        amt = t["amount"].dropna()
        if len(amt) > 20:
            thresh = amt.quantile(cfg.high_value_pctile)
            big = t[t["amount"] >= thresh].sort_values("amount", ascending=False)
            if len(big):
                top1 = big.iloc[0]
                label = _row_label(t, big.index[0])
                med_all = float(amt.median())
                xdist = ""
                if kind == "duration" and "distance_km" in t.columns:
                    d_case = t.loc[big.index[0], "distance_km"]
                    d_med = t["distance_km"].median()
                    if pd.notna(d_case) and pd.notna(d_med) and d_med > 0:
                        xdist = (f" The trip itself was {d_case:.1f} km (typical: {d_med:.1f} km)"
                                 + (", so distance is not the reason." if d_case < 1.5 * d_med else
                                    " — unusually far, which may be part of the story."))
                if kind == "duration":
                    r.add("The most extreme case",
                          f"The slowest case in the whole file is {label}: {top1['amount']:g} against a usual {med_all:g}.{xdist}",
                          f"{top1['amount']:g} vs usual {med_all:g}",
                          "Was the customer affected, and did anyone follow up? A case this far out of line usually has a story — a breakdown, a wrong address, a lost order — that is worth knowing.",
                          "Low")
                else:
                    r.add("The most extreme case",
                          f"The single biggest line is {label}: {top1['amount']:g} — {top1['amount'] / total:.0%} of everything spent.",
                          f"{top1['amount']:g}",
                          "Did anyone sign this off before the money was spent — and could it have been cheaper if it had been arranged earlier?",
                          "Low")

    # --- Partial-period detection ---
    if has_date and has_amt and t["date"].notna().any():
        per = t.dropna(subset=["date"]).copy()
        per["ym"] = per["date"].dt.to_period("M")
        counts = per.groupby("ym").size()
        if len(counts) >= 2:
            med = counts.median()
            last_ym, last_n = counts.index[-1], counts.iloc[-1]
            if last_n < cfg.partial_period_ratio * med:
                r.add("Unfinished period",
                      f"The last period ({last_ym}) only has {int(last_n)} records, when a normal month has about {int(med)} — it is clearly not finished yet.",
                      f"{int(last_n)} vs {int(med)}",
                      "Anyone comparing months will think activity collapsed at the end. Should unfinished periods be left out of charts and forecasts automatically?",
                      "Medium")

    # --- Currency mono-culture heuristic (money only) ---
    if has_amt and kind == "money":
        for c in t.columns:
            if c.startswith("dim::"):
                vals = set(v.lower() for v in t[c].dropna().unique())
                noneuro = {"poland", "uk", "united kingdom", "usa", "united states",
                           "switzerland", "japan", "china", "sweden", "norway", "denmark",
                           "czech", "hungary", "turkey", "croatia"}
                if vals & noneuro and any(k in c.lower() for k in ("country", "region", "location")):
                    r.add("Currency / FX",
                          "Data spans non-euro geographies but is reported in a single currency with no FX rate or conversion date attached.",
                          "mixed geographies",
                          "Which corporate FX convention applies (transaction-date spot vs monthly average), and should the source capture original currency so conversion is auditable?",
                          "High")
                    break

    if not r.findings:
        r.add("Clean", "No structural red flags detected by the standard check battery.",
              "0 issues", "Data passed automated checks — still confirm source-of-truth and period close manually.", "Low")
    return r


# --------------------------------------------------------------------------- #
# Generic data-quality profiling — works on ANY tabular dataset
# --------------------------------------------------------------------------- #
_ID_LIKE = ("id", "code", "lat", "lon", "zip", "postal", "phone", "guid", "uuid")


def _is_id_like(name: str) -> bool:
    n = str(name).lower()
    return any(k in n for k in _ID_LIKE)


def _raw_row_label(df: pd.DataFrame, idx) -> str:
    """Point at one raw-dataframe row like a person would: 'Order ID xxx on 2022-03-19'."""
    row = df.loc[idx]
    parts = []
    for c in df.columns:
        if _is_id_like(c) and df[c].nunique(dropna=True) / max(len(df), 1) > 0.5:
            parts.append(f"{str(c).replace('_', ' ')} {row[c]}")
            break
    for c in df.columns:
        if "date" in str(c).lower() and pd.notna(row.get(c)):
            val = row[c]
            if hasattr(val, "date"):
                val = val.date()
            parts.append(f"on {str(val)[:10]}")
            break
    return " ".join(str(p) for p in parts) if parts else "one record"


def clean_raw(df: pd.DataFrame) -> pd.DataFrame:
    """Drop summary/footer rows from a raw dataframe before profiling:
    rows whose text cells say 'total' etc., or rows that are mostly empty."""
    if len(df) == 0:
        return df
    keep = pd.Series(True, index=df.index)
    for c in df.columns:
        if df[c].dtype == object:
            vals = df[c].astype(str).str.strip().str.lower()
            keep &= ~vals.isin(_TOTAL_WORDS)
    keep &= df.isna().mean(axis=1) <= 0.5
    return df[keep]


def profile_columns(df: pd.DataFrame, max_cols: int = 50) -> list:
    """Per-column statistics: type, nulls, distinct, min/max/mean/median/std,
    top value. Returns JSON-safe rows for any tabular dataset."""
    df = clean_raw(df)
    rows = []
    n = len(df)
    for c in list(df.columns)[:max_cols]:
        s = df[c]
        nulls = int(s.isna().sum())
        distinct = int(s.nunique(dropna=True))
        row = {"column": str(c), "dtype": str(s.dtype), "nulls": nulls,
               "null_pct": round(100 * nulls / max(n, 1), 1), "distinct": distinct,
               "min": None, "max": None, "mean": None, "median": None,
               "std": None, "top": None}
        if pd.api.types.is_numeric_dtype(s) and distinct > 1:
            v = s.dropna().astype(float)
            row.update(min=round(float(v.min()), 2), max=round(float(v.max()), 2),
                       mean=round(float(v.mean()), 2), median=round(float(v.median()), 2),
                       std=round(float(v.std()), 2))
        else:
            vs = s.dropna().astype(str)
            if "date" in str(c).lower() and len(vs):
                parsed = pd.to_datetime(vs, errors="coerce")
                if parsed.notna().mean() > 0.8:
                    row["min"] = str(parsed.min().date())
                    row["max"] = str(parsed.max().date())
            if len(vs):
                vc = vs.value_counts()
                row["top"] = f"{vc.index[0]} ({int(vc.iloc[0])}x)"
        rows.append(row)
    return rows


def generic_checks(df: pd.DataFrame, cfg: Config, r: Result, skip_cols=()):
    """Column-level quality checks on the RAW dataframe (all columns), each
    phrased as a detected fact + a strategic question. Pass the detected
    measure column in skip_cols: it is already analysed per category by
    run_checks, so re-flagging it here would just repeat the same story."""
    df = clean_raw(df)
    n = len(df)
    if n == 0:
        return

    # --- fully identical rows ---
    dup = int(df.duplicated(keep=False).sum())
    if dup:
        sev = "High" if dup / n > 0.02 else "Medium"
        r.add("Row duplication",
              f"{dup} fully identical rows detected out of {n}.",
              f"{dup} rows",
              "Is the extract producing repeats, and which system enforces row-level uniqueness as the single source of truth?",
              sev)

    # --- missing data per column (worst 5) ---
    null_cols = sorted(((c, int(df[c].isna().sum())) for c in df.columns
                        if df[c].isna().any()), key=lambda x: -x[1])
    for c, k in null_cols[:5]:
        share = k / n
        sev = "High" if share > 0.20 else ("Medium" if share > 0.05 else "Low")
        r.add(f"Missing data: {c}",
              f"'{c}' has {k} missing values ({share:.1%} of rows).",
              f"{k} nulls",
              f"When '{c}' is empty, that row simply drops out of any view that uses it. Should the field be mandatory when the record is created?",
              sev)
    if len(null_cols) > 5:
        rest = ", ".join(c for c, _ in null_cols[5:10])
        r.add("Missing data: other columns",
              f"{len(null_cols) - 5} further column(s) also contain nulls ({rest}...).",
              f"{len(null_cols)} cols with nulls",
              "Is there a data-quality SLA with the source system covering completeness of all fields?",
              "Low")

    # --- constant columns ---
    consts = [c for c in df.columns if df[c].nunique(dropna=True) <= 1][:3]
    for c in consts:
        r.add(f"Constant column: {c}",
              f"'{c}' holds a single value across all {n} rows — it carries no information.",
              "1 distinct",
              f"Is '{c}' filtered upstream (hiding real variation), or should it be dropped from the extract?",
              "Low")

    # --- numeric columns: outliers / negatives / zero inflation ---
    num_cols = [c for c in df.columns
                if pd.api.types.is_numeric_dtype(df[c])
                and not _is_id_like(c) and c not in (skip_cols or ())
                and df[c].nunique(dropna=True) > 10]
    flagged = 0
    for c in num_cols:
        if flagged >= 5:
            break
        v = df[c].dropna().astype(float)
        if len(v) < 20:
            continue
        q1, q3 = v.quantile(0.25), v.quantile(0.75)
        iqr = q3 - q1
        if iqr > 0:
            extreme = v[(v > q3 + 3 * iqr) | (v < q1 - 3 * iqr)]
            if len(extreme):
                lo, hi = float(extreme.min()), float(extreme.max())
                both = f"as low as {lo:g} and as high as {hi:g}" if lo != hi else f"reaching {hi:g}"
                worst = extreme.idxmax() if hi > float(q3) else extreme.idxmin()
                wval = float(df.loc[worst, c])
                label = _raw_row_label(df, worst)
                r.add(f"Values that don't look right: {c}",
                      f"'{c}' normally sits between {q1:g} and {q3:g}, but {len(extreme)} record(s) fall far outside — "
                      f"{both}. For example {label} has {c} = {wval:g}.",
                      f"{len(extreme)} suspicious",
                      f"Can '{c}' really be {wval:g}? If this field has a hard possible range (like a 1-5 rating), "
                      f"the system should refuse anything outside it when the data is entered.",
                      "Medium")
                flagged += 1
        neg = int((v < 0).sum())
        if neg and (v >= 0).mean() > 0.9:
            r.add(f"Negative values: {c}",
                  f"'{c}' is almost entirely positive but contains {neg} negative value(s).",
                  f"{neg} negative",
                  f"How can '{c}' be negative? If it never should be, the entry screen must reject it.",
                  "Medium")
        zshare = float((v == 0).mean())
        if zshare > 0.30:
            r.add(f"Zero inflation: {c}",
                  f"{zshare:.0%} of '{c}' values are exactly zero — averages over this column are misleading.",
                  f"{zshare:.0%} zeros",
                  f"Do the zeros in '{c}' mean 'none', 'unknown' or 'does not apply'? Each one should be handled differently before averages are trusted.",
                  "Medium")

    # --- dominant category (near-constant text columns) ---
    obj_cols = [c for c in df.columns if df[c].dtype == object
                and 1 < df[c].nunique(dropna=True) <= 50]
    shown = 0
    for c in obj_cols:
        if shown >= 3:
            break
        vc = df[c].dropna().astype(str).value_counts(normalize=True)
        if len(vc) and vc.iloc[0] > 0.95:
            r.add(f"Dominant value: {c}",
                  f"Almost every row ({vc.iloc[0]:.0%}) has the same value in '{c}': '{vc.index[0]}'.",
                  f"{vc.iloc[0]:.0%} one value",
                  f"Is reality really this uniform, or is '{c}' being auto-filled with a default that nobody changes?",
                  "Low")
            shown += 1

    # --- future dates ---
    today = pd.Timestamp.now()
    for c in df.columns:
        if "date" in str(c).lower():
            parsed = pd.to_datetime(df[c], errors="coerce")
            if parsed.notna().mean() > 0.8:
                fut = int((parsed > today).sum())
                if fut:
                    r.add(f"Future dates: {c}",
                          f"'{c}' contains {fut} date(s) that have not happened yet.",
                          f"{fut} future",
                          f"Are the future dates in '{c}' planned events, or did someone mistype the year? Should dates after today be blocked at entry?",
                          "Medium")


# --------------------------------------------------------------------------- #
# Baseline comparison — challenge UPDATED assumptions
# --------------------------------------------------------------------------- #
def compare_baseline(new: pd.DataFrame, old: pd.DataFrame, cfg: Config, r: Result):
    if "amount" not in new.columns or "amount" not in old.columns:
        return
    nt, ot = float(new["amount"].sum()), float(old["amount"].sum())
    if ot:
        chg = nt / ot - 1
        r.changes.append(("Total spend", f"{ot:,.2f}", f"{nt:,.2f}", f"{chg:+.1%}"))
        if abs(chg) >= cfg.material_change:
            r.add("Assumption change",
                  f"Total spend moved {chg:+.1%} vs the baseline file.",
                  f"{chg:+.1%}",
                  "What drove the total to move materially since the last version — scope, price, volume, or a data fix?",
                  "High")
    if "category" in new.columns and "category" in old.columns:
        nb = new.groupby("category")["amount"].sum()
        ob = old.groupby("category")["amount"].sum()
        new_cats = set(nb.index) - set(ob.index)
        gone_cats = set(ob.index) - set(nb.index)
        if new_cats and cfg.new_category_flag:
            r.add("New categories",
                  f"New category(ies) appeared: {', '.join(sorted(new_cats))}.",
                  f"{len(new_cats)} new",
                  "Are the new categories genuine new spend, or re-labelling of existing lines? Does the taxonomy still tie out?",
                  "Medium")
        if gone_cats:
            r.add("Dropped categories",
                  f"Category(ies) disappeared vs baseline: {', '.join(sorted(gone_cats))}.",
                  f"{len(gone_cats)} dropped",
                  "Why did these categories vanish — no activity, or reclassified/merged elsewhere?",
                  "Medium")
        for cat in sorted(set(nb.index) & set(ob.index)):
            o, nn = float(ob[cat]), float(nb[cat])
            if o and abs(nn / o - 1) >= max(cfg.material_change, 0.25):
                d = nn / o - 1
                r.changes.append((f"Category: {cat}", f"{o:,.0f}", f"{nn:,.0f}", f"{d:+.0%}"))
                r.add("Category swing",
                      f"'{cat}' changed {d:+.0%} vs baseline ({o:,.0f} -> {nn:,.0f}).",
                      f"{d:+.0%}",
                      f"What is the business reason for the '{cat}' move? Is the new assumption supportable?",
                      "Medium" if abs(d) < 0.5 else "High")


# --------------------------------------------------------------------------- #
# Report writer
# --------------------------------------------------------------------------- #
def write_report(r: Result, out_path: str, src_name: str, cols: dict, baseline_name: str | None, cfg: Config):
    F = cfg.font
    title = Font(name=F, size=16, bold=True, color="7F1D1D")
    sub = Font(name=F, size=10, italic=True, color="595959")
    hf = Font(name=F, size=10, bold=True, color="FFFFFF")
    hfill = PatternFill("solid", fgColor="C00000")
    bf = Font(name=F, size=10)
    bold = Font(name=F, size=10, bold=True)
    ev = Font(name=F, size=10, bold=True, color="C00000")
    sev_fill = {"High": PatternFill("solid", fgColor="F4CCCC"),
                "Medium": PatternFill("solid", fgColor="FCE5CD"),
                "Low": PatternFill("solid", fgColor="FFF2CC")}
    alt = PatternFill("solid", fgColor="FBEAEA")
    thin = Side(style="thin", color="BFBFBF")
    bd = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical="top")
    ctr = Alignment(horizontal="center", vertical="top", wrap_text=True)

    wb = Workbook()

    # --- Findings sheet ---
    ws = wb.active
    ws.title = "Data Challenge"
    ws.sheet_view.showGridLines = False
    ws["A1"] = "Data Challenge Report"
    ws["A1"].font = title
    stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
    meta = f"Source: {src_name} | Generated: {stamp}"
    if baseline_name:
        meta += f" | Baseline: {baseline_name}"
    ws["A2"] = meta
    ws["A2"].font = sub
    ws["A3"] = (f"Detected columns  ->  Date: {cols.get('date')} | Amount: {cols.get('amount')} | "
                f"Category: {cols.get('category')}")
    ws["A3"].font = sub
    for c, w in zip("ABCDEF", [5, 22, 44, 16, 48, 10]):
        ws.column_dimensions[c].width = w
    hr = 5
    for i, h in enumerate(["#", "Issue Area", "Observation / Inconsistency",
                           "Evidence", "Question to Ask", "Severity"], 1):
        cell = ws.cell(hr, i, h)
        cell.font = hf; cell.fill = hfill; cell.border = bd
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    order = {"High": 0, "Medium": 1, "Low": 2}
    fs = sorted(r.findings, key=lambda x: order.get(x.severity, 3))
    rr = hr + 1
    for i, f in enumerate(fs, 1):
        ws.cell(rr, 1, i).font = bold; ws.cell(rr, 1).alignment = ctr; ws.cell(rr, 1).border = bd
        ws.cell(rr, 2, f.area).font = bold; ws.cell(rr, 2).alignment = wrap; ws.cell(rr, 2).border = bd
        ws.cell(rr, 3, f.observation).font = bf; ws.cell(rr, 3).alignment = wrap; ws.cell(rr, 3).border = bd
        ws.cell(rr, 4, f.evidence).font = ev; ws.cell(rr, 4).alignment = ctr; ws.cell(rr, 4).border = bd
        ws.cell(rr, 5, f.question).font = bf; ws.cell(rr, 5).alignment = wrap; ws.cell(rr, 5).border = bd
        c = ws.cell(rr, 6, f.severity); c.font = bold; c.alignment = ctr; c.border = bd
        c.fill = sev_fill.get(f.severity, alt)
        ws.row_dimensions[rr].height = 46
        rr += 1
    ws.freeze_panes = "A6"

    # --- Profile sheet ---
    ps = wb.create_sheet("Data Profile")
    ps.sheet_view.showGridLines = False
    ps["A1"] = "Data Profile"; ps["A1"].font = title
    ps.column_dimensions["A"].width = 26; ps.column_dimensions["B"].width = 26
    pr = 3
    for k, v in r.profile.items():
        ps.cell(pr, 1, k.replace("_", " ").title()).font = bold
        ps.cell(pr, 2, v).font = bf
        pr += 1

    # --- Change vs baseline sheet ---
    if r.changes:
        cs = wb.create_sheet("Change vs Baseline")
        cs.sheet_view.showGridLines = False
        cs["A1"] = "Change vs Baseline"; cs["A1"].font = title
        for c, w in zip("ABCD", [30, 18, 18, 14]):
            cs.column_dimensions[c].width = w
        for i, h in enumerate(["Metric", "Baseline", "Current", "Change"], 1):
            cell = cs.cell(3, i, h); cell.font = hf; cell.fill = hfill; cell.border = bd
            cell.alignment = Alignment(horizontal="center")
        cr = 4
        for metric, old, new, delta in r.changes:
            cs.cell(cr, 1, metric).font = bold; cs.cell(cr, 1).border = bd
            cs.cell(cr, 2, old).font = bf; cs.cell(cr, 2).border = bd; cs.cell(cr, 2).alignment = Alignment(horizontal="right")
            cs.cell(cr, 3, new).font = bf; cs.cell(cr, 3).border = bd; cs.cell(cr, 3).alignment = Alignment(horizontal="right")
            cs.cell(cr, 4, delta).font = ev; cs.cell(cr, 4).border = bd; cs.cell(cr, 4).alignment = Alignment(horizontal="center")
            cr += 1

    wb.save(out_path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Challenge the data in a CSV/Excel expense file.")
    ap.add_argument("input", help="CSV or Excel file to challenge")
    ap.add_argument("--baseline", help="Prior version to compare against (challenge updated assumptions)")
    ap.add_argument("--sheet", help="Sheet name for Excel inputs")
    ap.add_argument("--baseline-sheet", help="Sheet name for the baseline Excel file")
    ap.add_argument("--out", help="Output .xlsx report path")
    ap.add_argument("--date-col")
    ap.add_argument("--amount-col")
    ap.add_argument("--category-col")
    args = ap.parse_args(argv)

    cfg = Config()
    overrides = {"date": args.date_col, "amount": args.amount_col, "category": args.category_col}

    df = load_table(args.input, args.sheet)
    cols = detect_columns(df, overrides)
    t = normalize(df, cols)
    result = run_checks(t, cfg)

    baseline_name = None
    if args.baseline:
        baseline_name = os.path.basename(args.baseline)
        bdf = load_table(args.baseline, args.baseline_sheet)
        bcols = detect_columns(bdf, overrides)
        bt = normalize(bdf, bcols)
        compare_baseline(t, bt, cfg, result)

    out = args.out or (os.path.splitext(args.input)[0] + "_challenge_report.xlsx")
    write_report(result, out, os.path.basename(args.input), cols, baseline_name, cfg)

    # console summary
    print(f"Data Challenge Report -> {out}")
    print(f"Detected: date={cols.get('date')} amount={cols.get('amount')} category={cols.get('category')}")
    print(f"Profile: {result.profile}")
    hi = sum(1 for f in result.findings if f.severity == "High")
    md = sum(1 for f in result.findings if f.severity == "Medium")
    lo = sum(1 for f in result.findings if f.severity == "Low")
    print(f"Findings: {len(result.findings)} (High {hi} / Medium {md} / Low {lo})")
    for f in result.findings:
        print(f"  [{f.severity:6}] {f.area}: {f.observation}")
    if result.changes:
        print("Changes vs baseline:")
        for m, o, n, d in result.changes:
            print(f"  {m}: {o} -> {n} ({d})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
