#!/usr/bin/env python3
"""
Data Challenge - Assignment & Response-Link Layer
=================================================

Workspace foundation for the "ask a named colleague to answer each question"
workflow. It sits on top of the existing rule engine and PDF renderer and adds:

  1. Users master data  -> a governed list of reviewers (tools/users_master.json).
  2. Assignment          -> each finding is routed to a user via master-data rules.
  3. Tokenised links     -> every finding gets a unique, unguessable, expiring
                            response token + link (one per question).
  4. A request store     -> a JSON "outbox" (one record per question) that a
                            future email step / web app would consume, plus a
                            human-readable CSV, plus a PDF that shows the assignee.

NOT included here (needs infrastructure, by design): actually sending the email
(needs an email integration + credentials) and the page where the recipient
types a reply (needs a hosted backend + database). The 'base_url' is therefore a
placeholder until a server exists. This module produces everything up to that
handoff so the loop is ready to wire in.

Usage
-----
    python tools/challenge_workflow.py INPUT.csv
    python tools/challenge_workflow.py INPUT.xlsx --sheet "Raw Data" \
        --users tools/users_master.json --out-prefix Travel_Review

Outputs (given --out-prefix NAME):
    NAME_review.pdf        review sheet with an "Assigned To" column
    NAME_requests.json     the request store / outbox (tokens, links, status)
    NAME_assignments.csv   flat who-answers-what list

Dependencies: pandas, openpyxl, fpdf2 (all already available).
"""
from __future__ import annotations

import argparse
import csv
import json
import os
import secrets
import sys
from datetime import datetime, timedelta, timezone

from data_challenge import (Config, load_table, detect_columns, normalize,
                            run_checks, generic_checks, profile_columns)
from data_challenge_pdf import render_pdf


# --------------------------------------------------------------------------- #
# Master data
# --------------------------------------------------------------------------- #
def load_users(path: str) -> dict:
    with open(path, "r", encoding="utf-8") as fh:
        data = json.load(fh)
    by_id = {u["id"]: u for u in data.get("users", [])}
    data["_by_id"] = by_id
    return data


def resolve_assignee(area: str, master: dict) -> dict | None:
    """Route an issue area to a user via master-data rules; fall back to default."""
    a = (area or "").lower()
    for rule in master.get("routing", []):
        if rule.get("match", "").lower() in a:
            u = master["_by_id"].get(rule.get("assignee"))
            if u and u.get("active", True):
                return u
    u = master["_by_id"].get(master.get("default_assignee"))
    return u if (u and u.get("active", True)) else None


# --------------------------------------------------------------------------- #
# Tokens & request store
# --------------------------------------------------------------------------- #
def make_token() -> str:
    # URL-safe, unguessable; a real deployment would also sign or DB-back this.
    return secrets.token_urlsafe(18)


def build_report_id(src_name: str) -> str:
    stamp = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    stem = os.path.splitext(os.path.basename(src_name))[0]
    return f"{stem}-{stamp}"


def build_requests(result, master: dict, src_name: str, cols: dict):
    """Assign every finding, mint a token/link, and return (requests, assign_map).

    assign_map is keyed by id(finding) so the PDF renderer can show the assignee
    against the exact same finding objects.
    """
    settings = master.get("settings", {})
    base_url = settings.get("base_url", "https://REPLACE-ME/respond")
    expiry_days = int(settings.get("link_expiry_days", 14))
    now = datetime.now(timezone.utc)
    report_id = build_report_id(src_name)

    order = {"High": 0, "Medium": 1, "Low": 2}
    findings = sorted(result.findings, key=lambda f: order.get(f.severity, 3))

    requests = []
    assign_map = {}
    for i, f in enumerate(findings, 1):
        user = resolve_assignee(f.area, master)
        token = make_token()
        link = f"{base_url}?token={token}"
        rec = {
            "token": token,
            "response_link": link,
            "report_id": report_id,
            "finding_no": i,
            "issue_area": f.area,
            "observation": f.observation,
            "evidence": f.evidence,
            "question": f.question,
            "severity": f.severity,
            "assignee_id": user["id"] if user else None,
            "assignee_name": user["name"] if user else "(unassigned)",
            "assignee_email": user["email"] if user else None,
            "status": "pending",
            "created_utc": now.isoformat(),
            "expires_utc": (now + timedelta(days=expiry_days)).isoformat(),
            "response": None,
            "responded_utc": None,
        }
        requests.append(rec)
        assign_map[id(f)] = {
            "name": user["name"] if user else "(unassigned)",
            "email": user["email"] if user else None,
            "link": link,
        }

    store = {
        "report_id": report_id,
        "source_file": os.path.basename(src_name),
        "generated_utc": now.isoformat(),
        "base_url": base_url,
        "link_expiry_days": expiry_days,
        "detected_columns": {k: cols.get(k) for k in ("date", "amount", "category")},
        "requests": requests,
    }
    return store, assign_map


def save_store(store: dict, path: str):
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(store, fh, indent=2, ensure_ascii=False)


def save_assignments_csv(store: dict, path: str):
    fields = ["finding_no", "severity", "issue_area", "assignee_name",
              "assignee_email", "question", "status", "response_link", "expires_utc"]
    with open(path, "w", newline="", encoding="utf-8") as fh:
        w = csv.DictWriter(fh, fieldnames=fields, extrasaction="ignore")
        w.writeheader()
        for r in store["requests"]:
            w.writerow(r)


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    ap = argparse.ArgumentParser(description="Assign Data Challenge findings to users and mint response links.")
    ap.add_argument("input", help="CSV or Excel file to challenge")
    ap.add_argument("--sheet", help="Sheet name for Excel inputs")
    ap.add_argument("--users", default=os.path.join(os.path.dirname(__file__), "users_master.json"),
                    help="Path to users master-data JSON")
    ap.add_argument("--out-prefix", help="Output filename prefix (defaults to input stem)")
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

    master = load_users(args.users)
    store, assign_map = build_requests(result, master, args.input, cols)
    store["column_profile"] = col_profile

    prefix = args.out_prefix or os.path.splitext(os.path.basename(args.input))[0]
    pdf_path = f"{prefix}_review.pdf"
    json_path = f"{prefix}_requests.json"
    csv_path = f"{prefix}_assignments.csv"

    render_pdf(result, pdf_path, os.path.basename(args.input), cols,
               assignments=assign_map, profile=col_profile)
    save_store(store, json_path)
    save_assignments_csv(store, csv_path)

    print(f"Review PDF        -> {pdf_path}")
    print(f"Request store     -> {json_path}")
    print(f"Assignments CSV   -> {csv_path}")
    print(f"Report id         : {store['report_id']}")
    if store["base_url"].startswith("https://REPLACE-ME"):
        print("NOTE: base_url is a placeholder - set settings.base_url in the users "
              "master data once a response server/form exists. Links won't resolve until then.")
    print("Assignments:")
    for r in store["requests"]:
        print(f"  #{r['finding_no']} [{r['severity']:6}] {r['issue_area']:22} -> "
              f"{r['assignee_name']} <{r['assignee_email']}>")
    return 0


if __name__ == "__main__":
    sys.exit(main())
