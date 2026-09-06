#!/usr/bin/env python3
"""
Data Challenge -> PDF
=====================

Prototype of the "drop a file in, get a review PDF out" app.

A finance user points this at any expense/transaction file (CSV or Excel, with
column names/order that may vary). It:

  1. Reuses the rule engine in ``data_challenge.py`` to auto-detect the Date,
     Amount and Category columns and run the data-quality / FP&A check battery.
  2. Renders the findings as a one- or two-page PDF that mirrors the workbook's
     "Data Challenge" tab: #, Issue Area, Observation, Evidence, Question to Ask,
     Severity, with severity colour-coding and a High/Medium/Low legend.

Only the Data Challenge review sheet is produced (no summary tabs), by design.

Usage
-----
    python tools/data_challenge_pdf.py INPUT.csv
    python tools/data_challenge_pdf.py INPUT.xlsx --sheet "Raw Data"
    python tools/data_challenge_pdf.py INPUT.csv --out review.pdf \
        --date-col Date --amount-col "Amount (EUR)" --category-col "Expense Type"

Dependencies: pandas, openpyxl (preinstalled) + fpdf2.
"""
from __future__ import annotations

import argparse
import os
import sys
from datetime import datetime

from fpdf import FPDF

# Reuse the existing rule engine unchanged.
from data_challenge import (
    Config,
    load_table,
    detect_columns,
    normalize,
    run_checks,
    generic_checks,
    profile_columns,
)


# --------------------------------------------------------------------------- #
# Styling
# --------------------------------------------------------------------------- #
BRAND = (127, 29, 29)        # dark red title
HEADER_FILL = (192, 0, 0)    # header band
WHITE = (255, 255, 255)
GREY = (89, 89, 89)
GRID = (191, 191, 191)
SEV_FILL = {
    "High": (244, 204, 204),
    "Medium": (252, 229, 205),
    "Low": (255, 242, 204),
}
SEV_ORDER = {"High": 0, "Medium": 1, "Low": 2}

# The built-in PDF fonts are latin-1 only; map common Unicode punctuation/symbols
# to safe equivalents so any uploaded file renders without crashing.
_UNICODE_MAP = {
    "\u2014": "-", "\u2013": "-", "\u2012": "-", "\u2212": "-",
    "\u2018": "'", "\u2019": "'", "\u201c": '"', "\u201d": '"',
    "\u2026": "...", "\u20ac": "EUR ", "\u00a0": " ",
    "\u2192": "->", "\u2265": ">=", "\u2264": "<=", "\u00d7": "x",
    "\u03bc": "avg ", "\u00b5": "avg ",
}


def _s(text) -> str:
    """Make arbitrary text safe for the latin-1 core fonts."""
    out = str(text)
    for bad, good in _UNICODE_MAP.items():
        out = out.replace(bad, good)
    return out.encode("latin-1", "replace").decode("latin-1")

# Column layout (mm) for an A4 landscape page. Sum must fit inside the margins.
# Two layouts: with and without an "Assigned To" column.
COLS_PLAIN = [
    ("#", 8),
    ("Issue Area", 34),
    ("Observation / Inconsistency", 82),
    ("Evidence", 30),
    ("Question to Ask", 88),
    ("Severity", 20),
]
COLS_ASSIGNED = [
    ("#", 8),
    ("Issue Area", 30),
    ("Observation / Inconsistency", 64),
    ("Evidence", 24),
    ("Question to Ask", 74),
    ("Assigned To", 39),
    ("Severity", 18),
]
# Layout when responses are included: the Response column is the widest one and
# every row's height grows with the longest wrapped cell, so answers are always
# fully visible (never truncated).
COLS_FULL = [
    ("#", 8),
    ("Issue Area", 26),
    ("Observation / Inconsistency", 46),
    ("Evidence", 20),
    ("Question to Ask", 56),
    ("Assigned To", 26),
    ("Severity", 17),
    ("Response", 75),
]

# Active layout; render_pdf swaps this based on whether assignments are given.
COLS = COLS_PLAIN


class ReviewPDF(FPDF):
    def __init__(self, src_name: str, cols: dict):
        super().__init__(orientation="L", unit="mm", format="A4")
        self.src_name = src_name
        self.detected = cols
        self.section = "findings"   # controls which table header repeats per page
        self.set_auto_page_break(auto=True, margin=15)
        self.set_margins(10, 10, 10)
        self.set_title("Data Challenger Report")

    # ---- repeated header on every page ---------------------------------- #
    def header(self):
        self.set_font("Helvetica", "B", 16)
        self.set_text_color(*BRAND)
        self.cell(0, 8, _s("Data Challenger Report"), ln=1)

        stamp = datetime.now().strftime("%Y-%m-%d %H:%M")
        self.set_font("Helvetica", "I", 8)
        self.set_text_color(*GREY)
        self.cell(0, 4, _s(f"Source: {self.src_name}   |   Generated: {stamp}"), ln=1)
        d = self.detected
        self.cell(
            0, 4,
            _s(f"Detected columns  ->  Date: {d.get('date')}   |   "
               f"Amount: {d.get('amount')}   |   Category: {d.get('category')}"),
            ln=1,
        )
        self.ln(1)
        if self.section == "findings":
            self._table_header()

    def _table_header(self):
        self.set_font("Helvetica", "B", 8.5)
        self.set_text_color(*WHITE)
        self.set_fill_color(*HEADER_FILL)
        self.set_draw_color(*GRID)
        for title, w in COLS:
            self.cell(w, 7, _s(title), border=1, align="C", fill=True)
        self.ln(7)

    def footer(self):
        self.set_y(-12)
        self.set_font("Helvetica", "I", 7)
        self.set_text_color(*GREY)
        self.cell(0, 5, f"Page {self.page_no()}/{{nb}}", align="C")


def _row_height(pdf: ReviewPDF, values, line_h=4.2) -> float:
    """Compute the tallest wrapped cell so every column in the row aligns."""
    max_lines = 1
    for (title, w), text in zip(COLS, values):
        # dry_run returns the wrapped lines without drawing. Use the same inner
        # width (w - 2) as the actual drawing code so rows never overflow.
        lines = pdf.multi_cell(w - 2, line_h, _s(text), dry_run=True, output="LINES")
        max_lines = max(max_lines, len(lines))
    return max_lines * line_h + 2


PROFILE_COLS = [("Column", 40), ("Type", 18), ("Nulls", 14), ("Null %", 14),
                ("Distinct", 16), ("Min", 24), ("Max", 24), ("Mean", 22),
                ("Median", 22), ("Std", 20), ("Top value", 60)]


def _profile_section(pdf: ReviewPDF, profile: list):
    pdf.section = "profile"
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 12)
    pdf.set_text_color(*BRAND)
    pdf.cell(0, 8, _s("Column Profile"), new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "B", 7.5)
    pdf.set_text_color(*WHITE)
    pdf.set_fill_color(*HEADER_FILL)
    for title, w in PROFILE_COLS:
        pdf.cell(w, 6, _s(title), border=1, align="C", fill=True)
    pdf.ln(6)
    pdf.set_text_color(0, 0, 0)
    for p in profile:
        if pdf.get_y() + 6 > pdf.h - pdf.b_margin:
            pdf.add_page()
            pdf.set_font("Helvetica", "B", 7.5)
            pdf.set_text_color(*WHITE)
            pdf.set_fill_color(*HEADER_FILL)
            for title, w in PROFILE_COLS:
                pdf.cell(w, 6, _s(title), border=1, align="C", fill=True)
            pdf.ln(6)
            pdf.set_text_color(0, 0, 0)
        vals = [p.get("column"), p.get("dtype"), p.get("nulls"), f"{p.get('null_pct')}%",
                p.get("distinct"), p.get("min"), p.get("max"), p.get("mean"),
                p.get("median"), p.get("std"), p.get("top")]
        for (title, w), v in zip(PROFILE_COLS, vals):
            pdf.set_font("Helvetica", "B" if title == "Column" else "", 7.5)
            txt = "" if v is None else str(v)
            if len(txt) > 34:
                txt = txt[:31] + "..."
            pdf.cell(w, 6, _s(txt), border=1,
                     align="L" if title in ("Column", "Type", "Top value") else "R")
        pdf.ln(6)


def render_pdf(result, out_path: str, src_name: str, cols: dict,
               assignments: dict | None = None, profile: list | None = None,
               responses: dict | None = None):
    global COLS
    if assignments and responses is not None:
        COLS = COLS_FULL
    elif assignments:
        COLS = COLS_ASSIGNED
    else:
        COLS = COLS_PLAIN
    pdf = ReviewPDF(src_name, cols)
    pdf.alias_nb_pages()
    pdf.add_page()

    findings = sorted(result.findings, key=lambda f: SEV_ORDER.get(f.severity, 3))
    line_h = 4.2

    for i, f in enumerate(findings, 1):
        if assignments and responses is not None:
            who = assignments.get(id(f), {}).get("name", "(unassigned)")
            resp = responses.get(id(f)) or "(pending)"
            values = [str(i), f.area, f.observation, f.evidence, f.question, who,
                      f.severity, resp]
        elif assignments:
            who = assignments.get(id(f), {}).get("name", "(unassigned)")
            values = [str(i), f.area, f.observation, f.evidence, f.question, who, f.severity]
        else:
            values = [str(i), f.area, f.observation, f.evidence, f.question, f.severity]
        h = _row_height(pdf, values, line_h)

        # page break if the row would not fit
        if pdf.get_y() + h > pdf.h - pdf.b_margin:
            pdf.add_page()

        x0, y0 = pdf.get_x(), pdf.get_y()
        pdf.set_draw_color(*GRID)

        for idx, ((title, w), text) in enumerate(zip(COLS, values)):
            x, y = pdf.get_x(), pdf.get_y()
            # severity cell gets a colour fill
            fill = False
            if title == "Severity":
                pdf.set_fill_color(*SEV_FILL.get(f.severity, (255, 255, 255)))
                fill = True
            # outer border box for the whole cell height
            pdf.rect(x, y, w, h)

            # styling per column
            if title == "#":
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(0, 0, 0)
                align = "C"
            elif title == "Issue Area":
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(0, 0, 0)
                align = "L"
            elif title == "Evidence":
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(*HEADER_FILL)
                align = "C"
            elif title == "Assigned To":
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(0, 0, 0)
                align = "L"
            elif title == "Severity":
                pdf.set_font("Helvetica", "B", 8)
                pdf.set_text_color(0, 0, 0)
                align = "C"
            else:
                pdf.set_font("Helvetica", "", 8)
                pdf.set_text_color(0, 0, 0)
                align = "L"

            if fill:
                pdf.set_xy(x, y)
                pdf.cell(w, h, "", fill=True)  # paint background

            # vertically nudge single-line cells a touch
            pdf.set_xy(x + 1, y + 1)
            pdf.multi_cell(w - 2, line_h, _s(text), align=align)
            pdf.set_xy(x + w, y)  # move to next column start

        pdf.set_xy(x0, y0 + h)

    # legend
    pdf.ln(3)
    pdf.set_font("Helvetica", "B", 8)
    pdf.set_text_color(0, 0, 0)
    pdf.cell(20, 6, "Severity:", border=0)
    for sev in ("High", "Medium", "Low"):
        pdf.set_fill_color(*SEV_FILL[sev])
        pdf.cell(20, 6, sev, border=1, align="C", fill=True)
        pdf.cell(3, 6, "", border=0)
    pdf.ln(10)

    # profile line
    prof = result.profile
    pdf.set_font("Helvetica", "I", 7.5)
    pdf.set_text_color(*GREY)
    prof_str = "   |   ".join(
        f"{k.replace('_', ' ').title()}: {v}" for k, v in prof.items()
    )
    pdf.multi_cell(0, 4, _s(f"Data profile:   {prof_str}"))

    if profile:
        _profile_section(pdf, profile)

    pdf.output(out_path)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Produce a Data Challenge review PDF from a spend file.")
    ap.add_argument("input", help="CSV or Excel file to challenge")
    ap.add_argument("--sheet", help="Sheet name for Excel inputs")
    ap.add_argument("--out", help="Output .pdf path")
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
    generic_checks(df, cfg, result, skip_cols=[cols.get("amount")])
    col_profile = profile_columns(df)

    out = args.out or (os.path.splitext(args.input)[0] + "_challenge_review.pdf")
    render_pdf(result, out, os.path.basename(args.input), cols, profile=col_profile)

    hi = sum(1 for f in result.findings if f.severity == "High")
    md = sum(1 for f in result.findings if f.severity == "Medium")
    lo = sum(1 for f in result.findings if f.severity == "Low")
    print(f"Data Challenge review PDF -> {out}")
    print(f"Detected: date={cols.get('date')} amount={cols.get('amount')} category={cols.get('category')}")
    print(f"Findings: {len(result.findings)} (High {hi} / Medium {md} / Low {lo})")
    return 0


if __name__ == "__main__":
    sys.exit(main())
