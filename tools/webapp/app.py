#!/usr/bin/env python3
"""
Data Challenge - Local web prototype (Phase 1)
==============================================

Web app that runs on YOUR machine (localhost). It reuses the existing rule engine
and PDF generator and adds the screens for the full loop:

  1. Upload a data file (CSV/XLSX).
  2. See what was detected and who each question is assigned to (editable).
  3. Download the review PDF.
  4. Every question has a unique link: the reviewer opens it and answers IN THE APP.
  5. Responses flow back into the report and are shown on screen.

This is a LOCAL prototype: no login, no emails sent yet (links are shown to
copy/paste), and data is stored in tools/webapp/data/. Phase 2 (email) and
Phase 3 (hosting + hardening) come later.

Run:
    python tools/webapp/app.py
Then open in your browser:  http://127.0.0.1:5000
"""
from __future__ import annotations

import io
import json
import os
import sys
from datetime import datetime, timezone

# permitir importar el motor que vive en tools/
HERE = os.path.dirname(os.path.abspath(__file__))
TOOLS = os.path.dirname(HERE)
sys.path.insert(0, TOOLS)

from flask import (Flask, request, redirect, url_for, send_file, abort,
                   render_template_string, flash)
from werkzeug.utils import secure_filename

from data_challenge import (Config, load_table, detect_columns, normalize,
                            run_checks, generic_checks, profile_columns,
                            Result, Finding)
from data_challenge_pdf import render_pdf
from challenge_workflow import load_users, build_requests
from challenge_notify import notify_report
from hr_turnover import detect_hr, run_hr_checks

DATA = os.path.join(HERE, "data")
UPLOADS = os.path.join(DATA, "uploads")
REPORTS = os.path.join(DATA, "reports")
TOKENS_IDX = os.path.join(DATA, "tokens.json")
USERS_PATH = os.path.join(TOOLS, "users_master.json")
for d in (DATA, UPLOADS, REPORTS):
    os.makedirs(d, exist_ok=True)

app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "local-prototype-not-for-production")

# MVP folder structure: department -> topic. Reports are routed into a topic
# when their source filename contains one of the 'match' substrings.
FOLDERS = [
    {"dept": "LOGISTICS",  "topic": "Delivery Times",  "slug": "delivery-times",
     "match": ["delivery"]},
    {"dept": "ACCOUNTING", "topic": "Travel Expenses", "slug": "travel-expenses",
     "match": ["travel", "expense"]},
    {"dept": "HUMAN RESOURCES", "topic": "Employee Turnover", "slug": "employee-turnover",
     "match": ["turnover", "attrition", "employee", "churn"]},
]


def folder_by_slug(slug: str) -> dict | None:
    for f in FOLDERS:
        if f["slug"] == slug:
            return f
    return None


def report_in_folder(store: dict, folder: dict) -> bool:
    name = (store.get("source_file") or "").lower()
    return any(m in name for m in folder["match"])


def all_reports() -> list:
    out = []
    for fn in sorted(os.listdir(REPORTS)):
        if fn.endswith(".json"):
            s = load_report(fn[:-5])
            if s:
                out.append(s)
    return out


# --------------------------------------------------------------------------- #
# almacenamiento sencillo en JSON
# --------------------------------------------------------------------------- #
def _load_tokens() -> dict:
    if os.path.exists(TOKENS_IDX):
        with open(TOKENS_IDX, encoding="utf-8") as fh:
            return json.load(fh)
    return {}


def _save_tokens(idx: dict):
    with open(TOKENS_IDX, "w", encoding="utf-8") as fh:
        json.dump(idx, fh, indent=2, ensure_ascii=False)


def _report_path(rid: str) -> str:
    return os.path.join(REPORTS, secure_filename(rid) + ".json")


def load_report(rid: str) -> dict | None:
    p = _report_path(rid)
    if not os.path.exists(p):
        return None
    with open(p, encoding="utf-8") as fh:
        return json.load(fh)


def save_report(store: dict):
    with open(_report_path(store["report_id"]), "w", encoding="utf-8") as fh:
        json.dump(store, fh, indent=2, ensure_ascii=False)


def day_label(day: str) -> str:
    """Human-readable label for a YYYY-MM-DD day, e.g. 'Friday, 04 Sep 2026'."""
    try:
        return datetime.strptime(day, "%Y-%m-%d").strftime("%A, %d %b %Y")
    except (TypeError, ValueError):
        return day or ""


def find_request(store: dict, token: str) -> dict | None:
    for r in store["requests"]:
        if r["token"] == token:
            return r
    return None


# --------------------------------------------------------------------------- #
# reconstruir Result para el PDF
# --------------------------------------------------------------------------- #
def result_from_store(store: dict):
    res = Result()
    res.profile = store.get("profile", {})
    assign_map = {}
    resp_map = {}
    for r in store["requests"]:
        f = Finding(r["issue_area"], r["observation"], r["evidence"],
                    r["question"], r["severity"])
        res.findings.append(f)
        assign_map[id(f)] = {"name": r.get("assignee_name", "(unassigned)"),
                             "email": r.get("assignee_email"),
                             "link": r.get("response_link")}
        if r.get("status") == "answered" and r.get("response"):
            who = r.get("responder") or r.get("assignee_name") or ""
            when = (r.get("responded_utc") or "")[:10]
            resp_map[id(f)] = f"{r['response']}\n- {who}, {when}"
    return res, assign_map, resp_map


# --------------------------------------------------------------------------- #
# plantillas (HTML embebido)
# --------------------------------------------------------------------------- #
BASE = """
<!doctype html><html lang="es"><head><meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Data Challenger</title>
<style>
 body{font-family:Arial,Helvetica,sans-serif;margin:0;color:#222;background:#f6f6f8}
 header{background:#7f1d1d;color:#fff;padding:14px 24px}
 header h1{margin:0;font-size:20px}
 .wrap{max-width:1150px;margin:24px auto;padding:0 20px}
 .card{background:#fff;border:1px solid #e2e2e6;border-radius:8px;padding:20px;margin-bottom:20px}
 table{border-collapse:collapse;width:100%;font-size:13px}
 th,td{border:1px solid #d9d9de;padding:7px 8px;vertical-align:top;text-align:left}
 th{background:#c00000;color:#fff}
 tr:nth-child(even) td{background:#fbeaea}
 .sev-High{background:#f4cccc!important;font-weight:bold;text-align:center}
 .sev-Medium{background:#fce5cd!important;font-weight:bold;text-align:center}
 .sev-Low{background:#fff2cc!important;font-weight:bold;text-align:center}
 .ev{color:#c00000;font-weight:bold;text-align:center;white-space:nowrap}
 .btn{display:inline-block;background:#c00000;color:#fff;padding:9px 16px;border:0;
      border-radius:6px;text-decoration:none;font-size:14px;cursor:pointer}
 .btn.secondary{background:#555}
 input[type=file],select,textarea{font-size:14px;padding:6px;border:1px solid #bbb;border-radius:5px}
 #users-table input[type=text],#users-table input[type=email]{width:95%;font-size:13px;padding:5px;
   border:1px solid #ccc;border-radius:4px}
 #users-table input:disabled{background:transparent;border-color:transparent;color:#222}
 /* Response column: grows with the answer, wraps long text, never clipped */
 .doc-table{table-layout:auto}
 .doc-table .resp-col{min-width:240px;max-width:none;white-space:pre-wrap;
   word-break:break-word;overflow-wrap:anywhere}
 textarea{width:100%;min-height:120px}
 .muted{color:#777;font-size:13px}
 .flash{background:#fff3cd;border:1px solid #ffe08a;padding:10px 14px;border-radius:6px;margin-bottom:14px}
 .pill{padding:2px 8px;border-radius:10px;font-size:12px}
 .pill.pending{background:#eee;color:#555}
 .pill.answered{background:#d7f5dd;color:#1b7a34}
 .pill.notified{background:#dbe9fb;color:#1d4e89}
 .pill.dryrun{background:#fdf1d6;color:#8a6d1a}
 .resp-empty{color:#999;font-style:italic}
 code{background:#f0f0f2;padding:2px 5px;border-radius:4px;font-size:12px}
 .linkcell{max-width:230px;word-break:break-all}
</style></head><body>
<header><h1>Data Challenger</h1></header>
<div class="wrap">
 {% with msgs = get_flashed_messages() %}{% for m in msgs %}<div class="flash">{{m}}</div>{% endfor %}{% endwith %}
 {{ body|safe }}
</div></body></html>
"""

HOME = """
<div class="card">
 <h2>1) Upload the file</h2>
 <p class="muted">Formats: CSV or Excel (.xlsx). Column names may vary; they are auto-detected.</p>
 <form method="post" action="{{ url_for('upload') }}" enctype="multipart/form-data">
   <p><input type="file" name="file" accept=".csv,.xlsx,.xls" required></p>
   <p>Excel sheet (optional): <input type="text" name="sheet" placeholder="e.g. Raw Data"><br>
      <span class="muted">Only for Excel files with several tabs: type the name of the sheet that holds
      the data. Leave it blank to use the first sheet. Ignored for CSV files.</span></p>
   <p><button class="btn" type="submit">Analyse &amp; create review</button></p>
 </form>
</div>
<div class="card">
 <h3>Available reviewers (master data)</h3>
 <form method="post" action="{{ url_for('users_save') }}" id="users-form">
 <table id="users-table">
  <tr><th>Name</th><th>Email</th><th>Role</th><th>Department</th><th>Active</th></tr>
  {% for u in users %}
  <tr>
   <td><input type="text" name="name_{{u.id}}" value="{{u.name}}" disabled required></td>
   <td><input type="email" name="email_{{u.id}}" value="{{u.email}}" disabled required></td>
   <td><input type="text" name="role_{{u.id}}" value="{{u.role}}" disabled></td>
   <td><input type="text" name="dept_{{u.id}}" value="{{u.department}}" disabled></td>
   <td style="text-align:center"><input type="checkbox" name="active_{{u.id}}" {{ 'checked' if u.active else '' }} disabled></td>
  </tr>
  {% endfor %}
 </table>
 <p style="margin-top:12px">
  <button class="btn" type="button" id="btn-edit" onclick="editUsers()">Edit</button>
  <button class="btn" type="button" id="btn-add" onclick="addRow()" style="display:none">Add row</button>
  <button class="btn" type="submit" id="btn-save" style="display:none">Save</button>
  <button class="btn secondary" type="button" id="btn-cancel" onclick="location.reload()" style="display:none">Cancel</button>
 </p>
 </form>
 <script>
 function editUsers(){
   document.querySelectorAll('#users-table input').forEach(i => i.disabled = false);
   document.getElementById('btn-edit').style.display = 'none';
   ['btn-add','btn-save','btn-cancel'].forEach(id => document.getElementById(id).style.display = 'inline-block');
 }
 function addRow(){
   var t = document.getElementById('users-table');
   var r = t.insertRow(-1);
   r.innerHTML = '<td><input type="text" name="new_name" placeholder="Full name" required></td>' +
                 '<td><input type="email" name="new_email" placeholder="email@company.com" required></td>' +
                 '<td><input type="text" name="new_role" placeholder="Role"></td>' +
                 '<td><input type="text" name="new_dept" placeholder="Department"></td>' +
                 '<td style="text-align:center"><input type="checkbox" name="new_active" checked disabled title="New reviewers start active"></td>';
 }
 </script>
</div>
<div class="card">
 <h3>Departments</h3>
 {% for dept, topics in folders %}
  <p style="margin:12px 0 4px"><b>&#128193; {{ dept }}</b></p>
  <ul style="margin:4px 0">
  {% for t in topics %}
   <li><a href="{{ url_for('folder_view', slug=t.slug) }}">{{ t.topic }}</a>
       <span class="muted">&mdash; {{ t.n_reports }} review(s), {{ t.n_comments }} comment(s)</span></li>
  {% endfor %}
  </ul>
 {% endfor %}
</div>
"""

REPORT = """
<div class="card">
 <h2>Review: <span class="muted">{{ store.source_file }}</span></h2>
 <p>
   <a class="btn" href="{{ url_for('report_pdf', rid=store.report_id) }}">Download PDF</a>
   <a class="btn secondary" href="{{ url_for('home') }}">New review</a>
 </p>
 <form method="post" action="{{ url_for('report_notify', rid=store.report_id) }}" style="display:inline">
   <button class="btn" type="submit">Send questions via Slack</button>
 </form>
</div>
<div class="card">
 <h3>2) Findings, owners and responses</h3>
 {% if nav %}
 <div style="text-align:center;margin:6px 0 2px">
   <span class="muted">Responses as of</span><br>
   <span style="font-size:20px;font-weight:bold">{{ nav.label }}</span><br>
   <span class="muted">day {{ nav.pos }} of {{ nav.total }}</span>
 </div>
 <p style="text-align:center;margin-top:10px">
   {% if nav.prev %}<a class="btn" href="{{ url_for('report_view', rid=store.report_id, day=nav.prev) }}">&#9664; Previous day ({{ nav.prev }})</a>
   {% else %}<span class="btn secondary" style="opacity:.4">&#9664; Previous day</span>{% endif %}
   {% if nav.next %}<a class="btn" href="{{ url_for('report_view', rid=store.report_id, day=nav.next) }}">Next day ({{ nav.next }}) &#9654;</a>
   {% else %}<span class="btn secondary" style="opacity:.4">Next day &#9654;</span>{% endif %}
 </p>
 {% endif %}
 <form method="post" action="{{ url_for('report_assign', rid=store.report_id) }}">
 <table>
  <tr><th>#</th><th>Issue Area</th><th>What we detected</th><th>Evidence</th><th>Strategic question</th>
      <th>Assigned to</th><th>Sev.</th><th>Response</th></tr>
  {% for r in store.requests %}
  <tr>
   <td>{{r.finding_no}}</td>
   <td><b>{{r.issue_area}}</b></td>
   <td>{{r.observation}}</td>
   <td class="ev">{{r.evidence}}</td>
   <td>{{r.question}}</td>
   <td><select name="assignee_{{r.token}}">
        {% for u in users %}<option value="{{u.id}}" {{ 'selected' if u.id==r.assignee_id else '' }}>{{u.name}}</option>{% endfor %}
       </select></td>
   <td class="sev-{{r.severity}}">{{r.severity}}</td>
   <td>
     {% if answered_map[r.token] %}
       <span class="pill answered">answered</span>
       <div style="margin-top:6px">{{r.response}}</div>
       <div class="muted"><i>{{ r.responder or r.assignee_name }} &middot; {{ (r.responded_utc or '')[:16].replace('T',' ') }} UTC</i></div>
     {% else %}
       <span class="pill pending">pending</span>
       {% if r.notified_via=='slack' %}
         <span class="pill notified">sent via Slack {{ (r.notified_utc or '')[:10] }}</span>
       {% elif r.notified_via=='dry-run' %}
         <span class="pill dryrun">dry-run {{ (r.notified_utc or '')[:10] }}</span>
       {% else %}
         <div class="resp-empty">not notified yet</div>
       {% endif %}
       {% if r.notify_error %}<div class="muted" style="color:#b00">{{ r.notify_error }}</div>{% endif %}
     {% endif %}
   </td>
  </tr>
  {% endfor %}
 </table>
 <p style="margin-top:14px"><button class="btn" type="submit">Save owners</button></p>
 </form>
 <p class="muted">Response links are no longer shown here: they are delivered privately to each
 owner via Slack DM. Reviewers answer through their link; answers appear in the Response column.</p>
</div>

"""

RESPOND = """
<div class="card">
 <h2>Answer a review question</h2>
 <p class="muted">File: {{ store.source_file }} &middot; Severity: <b>{{ req.severity }}</b></p>
 <p><b>Issue area:</b> {{ req.issue_area }}</p>
 <p><b>What we detected:</b> {{ req.observation }}</p>
 <p><b>Question:</b> {{ req.question }}</p>
 <p class="muted">Assigned to: {{ req.assignee_name }}{% if req.assignee_email %} &lt;{{ req.assignee_email }}&gt;{% endif %}</p>
 <form method="post" action="{{ url_for('respond_post') }}">
  <input type="hidden" name="token" value="{{ req.token }}">
  <p><textarea name="response" placeholder="Type your answer here..." required>{{ req.response or '' }}</textarea></p>
  <p><input type="text" name="responder" placeholder="Your name (optional)" value="{{ req.responder or '' }}"></p>
  <p><button class="btn" type="submit">Submit response</button></p>
 </form>
</div>
"""

FOLDER = """
<div class="card">
 <h2>&#128193; {{ folder.dept }} / {{ folder.topic }}</h2>
 {% if not days %}
   <p class="muted">No responses have been recorded in this folder yet.</p>
   <p><a class="btn secondary" href="{{ url_for('home') }}">Home</a></p>
 {% else %}
   <p>
     <div style="text-align:center;margin:6px 0 2px">
     <span class="muted">You are viewing</span><br>
     <span style="font-size:22px;font-weight:bold">{{ day_label }}</span><br>
     <span class="muted">day {{ pos }} of {{ days|length }} &middot; {{ items|length }} comment(s)</span>
   </div>
   <p style="text-align:center;margin-top:12px">
     {% if prev_day %}<a class="btn" href="{{ url_for('folder_view', slug=folder.slug, day=prev_day) }}">&#9664; Previous day ({{ prev_day }})</a>
     {% else %}<span class="btn secondary" style="opacity:.4">&#9664; Previous day</span>{% endif %}
     {% if next_day %}<a class="btn" href="{{ url_for('folder_view', slug=folder.slug, day=next_day) }}">Next day ({{ next_day }}) &#9654;</a>
     {% else %}<span class="btn secondary" style="opacity:.4">Next day &#9654;</span>{% endif %}
     <a class="btn secondary" href="{{ url_for('home') }}">Home</a>
   </p>
 {% endif %}
</div>
{% for doc in docs %}
<div class="card">
 <h3>Document: {{ doc.source_file }}
     <a class="btn secondary" style="float:right" href="{{ url_for('report_view', rid=doc.report_id) }}">Open review</a></h3>
 <p class="muted">Responses obtained up to {{ day }}: <b>{{ doc.answered }}/{{ doc.total }}</b></p>
 <table class="doc-table">
  <tr><th>#</th><th>Issue Area</th><th>Question</th><th>Assigned to</th><th>Sev.</th><th>Response as of {{ day }}</th></tr>
  {% for r in doc.rows %}
  <tr>
   <td>{{ r.finding_no }}</td>
   <td><b>{{ r.issue_area }}</b></td>
   <td>{{ r.question }}</td>
   <td>{{ r.assignee_name }}</td>
   <td class="sev-{{ r.severity }}">{{ r.severity }}</td>
   <td class="resp-col">{% if r.answered %}<span class="pill answered">answered</span>
         <div style="margin-top:6px">{{ r.response }}</div>
       {% else %}<span class="pill pending">pending</span>{% endif %}</td>
  </tr>
  {% endfor %}
 </table>
</div>
{% endfor %}
"""

THANKS = """
<div class="card">
 <h2>Thank you! Response saved.</h2>
 <p class="muted">You can close this page. Your answer now appears in the review report.</p>
</div>
"""


def render(body_tmpl, **ctx):
    body = render_template_string(body_tmpl, **ctx)
    return render_template_string(BASE, body=body)


# --------------------------------------------------------------------------- #
# rutas
# --------------------------------------------------------------------------- #
@app.route("/")
def home():
    users = load_users(USERS_PATH)["users"]
    reports = all_reports()
    folders = []
    for f in FOLDERS:
        in_f = [s for s in reports if report_in_folder(s, f)]
        n_comments = sum(1 for s in in_f for r in s["requests"] if r["status"] == "answered")
        entry = dict(f, n_reports=len(in_f), n_comments=n_comments)
        placed = False
        for dept, topics in folders:
            if dept == f["dept"]:
                topics.append(entry)
                placed = True
        if not placed:
            folders.append((f["dept"], [entry]))
    return render(HOME, users=users, folders=folders)


@app.route("/folder/<slug>")
@app.route("/folder/<slug>/<day>")
def folder_view(slug, day=None):
    folder = folder_by_slug(slug)
    if not folder:
        abort(404)
    stores = [s for s in all_reports() if report_in_folder(s, folder)]

    items = []
    for s in stores:
        for r in s["requests"]:
            if r.get("status") == "answered" and r.get("responded_utc"):
                items.append({
                    "day": r["responded_utc"][:10],
                    "time": r["responded_utc"][11:16],
                    "report_id": s["report_id"],
                    "source_file": s.get("source_file", s["report_id"]),
                    "issue_area": r["issue_area"],
                    "severity": r["severity"],
                    "question": r["question"],
                    "response": r.get("response") or "",
                    "responder": r.get("responder") or r.get("assignee_name") or "",
                })
    days = sorted({i["day"] for i in items})
    if not days:
        return render(FOLDER, folder=folder, days=[], day=None, prev_day=None,
                      next_day=None, pos=0, items=[], docs=[])
    if day not in days:
        day = days[-1]
    idx = days.index(day)
    prev_day = days[idx - 1] if idx > 0 else None
    next_day = days[idx + 1] if idx < len(days) - 1 else None
    day_items = sorted((i for i in items if i["day"] == day), key=lambda i: i["time"])

    # Documents active that day: each report shown with the responses obtained
    # up to (and including) the selected day.
    docs = []
    active_ids = {i["report_id"] for i in day_items}
    for s in stores:
        if s["report_id"] not in active_ids:
            continue
        rows = []
        answered = 0
        for r in s["requests"]:
            done = bool(r.get("status") == "answered" and r.get("responded_utc")
                        and r["responded_utc"][:10] <= day)
            if done:
                answered += 1
            rows.append({"finding_no": r["finding_no"], "issue_area": r["issue_area"],
                         "question": r["question"], "assignee_name": r.get("assignee_name", ""),
                         "severity": r["severity"], "answered": done,
                         "response": r.get("response") or ""})
        docs.append({"report_id": s["report_id"], "source_file": s.get("source_file", ""),
                     "rows": rows, "answered": answered, "total": len(rows)})
    return render(FOLDER, folder=folder, days=days, day=day, day_label=day_label(day),
                  prev_day=prev_day, next_day=next_day, pos=idx + 1,
                  items=day_items, docs=docs)


@app.route("/users/save", methods=["POST"])
def users_save():
    master = load_users(USERS_PATH)
    # update existing reviewers
    for u in master["users"]:
        uid = u["id"]
        if f"name_{uid}" in request.form:
            u["name"] = request.form[f"name_{uid}"].strip() or u["name"]
            u["email"] = request.form[f"email_{uid}"].strip() or u["email"]
            u["role"] = request.form.get(f"role_{uid}", "").strip()
            u["department"] = request.form.get(f"dept_{uid}", "").strip()
            u["active"] = f"active_{uid}" in request.form
    # append new reviewers (ids are issued sequentially and stay stable)
    names = request.form.getlist("new_name")
    emails = request.form.getlist("new_email")
    roles = request.form.getlist("new_role")
    depts = request.form.getlist("new_dept")
    next_num = max((int(u["id"][1:]) for u in master["users"]
                    if u["id"][1:].isdigit()), default=0) + 1
    added = 0
    for i, name in enumerate(names):
        if not name.strip() or i >= len(emails) or not emails[i].strip():
            continue
        master["users"].append({
            "id": f"u{next_num:03d}",
            "name": name.strip(),
            "email": emails[i].strip(),
            "role": roles[i].strip() if i < len(roles) else "",
            "department": depts[i].strip() if i < len(depts) else "",
            "active": True,
        })
        next_num += 1
        added += 1
    master.pop("_by_id", None)  # runtime index, never persisted
    with open(USERS_PATH, "w", encoding="utf-8") as fh:
        json.dump(master, fh, indent=2, ensure_ascii=False)
    flash(f"Reviewers saved. {added} added." if added else "Reviewers saved.")
    return redirect(url_for("home"))


@app.route("/upload", methods=["GET", "POST"])
def upload():
    if request.method == "GET":
        flash("Please upload a file from the home page.")
        return redirect(url_for("home"))
    f = request.files.get("file")
    if not f or not f.filename:
        flash("Please select a file first.")
        return redirect(url_for("home"))
    fname = secure_filename(f.filename)
    saved = os.path.join(UPLOADS, fname)
    f.save(saved)
    sheet = request.form.get("sheet") or None

    try:
        df = load_table(saved, sheet)
        cfg = Config()
        hr = detect_hr(df)
        if hr:
            # HR / employee-turnover data: domain-specific checks + questions
            result = run_hr_checks(df, hr)
            generic_checks(df, cfg, result, skip_cols=[hr.get("income")])
            cols = {"date": hr.get("hire_date"), "amount": hr.get("income"),
                    "category": hr.get("department")}
        else:
            cols = detect_columns(df, {"date": None, "amount": None, "category": None})
            t = normalize(df, cols)
            result = run_checks(t, cfg)
            generic_checks(df, cfg, result, skip_cols=[cols.get("amount")])
        col_profile = profile_columns(df)
    except Exception as e:
        flash(f"Could not process the file: {e}")
        return redirect(url_for("home"))

    master = load_users(USERS_PATH)
    store, _ = build_requests(result, master, fname, cols)
    store["profile"] = result.profile
    store["column_profile"] = col_profile
    for r in store["requests"]:
        r["responder"] = None
        # Point the link at THIS app so the URL delivered via Slack resolves
        # (the master-data base_url is still a placeholder).
        r["response_link"] = url_for("respond", token=r["token"], _external=True)
    save_report(store)

    idx = _load_tokens()
    for r in store["requests"]:
        idx[r["token"]] = store["report_id"]
    _save_tokens(idx)

    return redirect(url_for("report_view", rid=store["report_id"]))


@app.route("/report/<rid>")
def report_view(rid):
    store = load_report(rid)
    if not store:
        abort(404)
    users = load_users(USERS_PATH)["users"]

    # Day-by-day navigation inside the findings table: pick a day and the
    # Response column shows only the answers obtained up to that day.
    day = request.args.get("day")
    days = sorted({r["responded_utc"][:10] for r in store["requests"]
                   if r.get("status") == "answered" and r.get("responded_utc")})
    nav = None
    if days:
        if day not in days:
            day = days[-1]
        idx = days.index(day)
        nav = {"day": day, "label": day_label(day),
               "prev": days[idx - 1] if idx > 0 else None,
               "next": days[idx + 1] if idx < len(days) - 1 else None,
               "pos": idx + 1, "total": len(days)}
    answered_map = {
        r["token"]: bool(r.get("status") == "answered" and r.get("responded_utc")
                         and (not day or r["responded_utc"][:10] <= day))
        for r in store["requests"]
    }
    return render(REPORT, store=store, users=users, nav=nav, answered_map=answered_map)


@app.route("/report/<rid>/assign", methods=["POST"])
def report_assign(rid):
    store = load_report(rid)
    if not store:
        abort(404)
    by_id = load_users(USERS_PATH)["_by_id"]
    for r in store["requests"]:
        uid = request.form.get(f"assignee_{r['token']}")
        if uid and uid in by_id:
            u = by_id[uid]
            r["assignee_id"] = u["id"]
            r["assignee_name"] = u["name"]
            r["assignee_email"] = u["email"]
    save_report(store)
    flash("Owners updated.")
    return redirect(url_for("report_view", rid=rid))


@app.route("/report/<rid>/notify", methods=["POST"])
def report_notify(rid):
    store = load_report(rid)
    if not store:
        abort(404)
    summary = notify_report(store)
    save_report(store)
    flash(f"Notifications ({summary['mode']}): {summary['sent']} sent, "
          f"{summary['skipped']} skipped, {summary['failed']} failed.")
    return redirect(url_for("report_view", rid=rid))


@app.route("/report/<rid>/pdf")
def report_pdf(rid):
    store = load_report(rid)
    if not store:
        abort(404)
    res, assign_map, resp_map = result_from_store(store)
    out = os.path.join(DATA, secure_filename(rid) + "_review.pdf")
    render_pdf(res, out, store["source_file"], store["detected_columns"],
               assignments=assign_map, responses=resp_map)
    return send_file(out, as_attachment=True, download_name=f"{rid}_review.pdf")


@app.route("/respond")
def respond():
    token = request.args.get("token", "")
    idx = _load_tokens()
    rid = idx.get(token)
    store = load_report(rid) if rid else None
    if not store:
        abort(404)
    req = find_request(store, token)
    if not req:
        abort(404)
    return render(RESPOND, store=store, req=req)


@app.route("/respond", methods=["POST"])
def respond_post():
    token = request.form.get("token", "")
    idx = _load_tokens()
    rid = idx.get(token)
    store = load_report(rid) if rid else None
    if not store:
        abort(404)
    req = find_request(store, token)
    if not req:
        abort(404)
    req["response"] = (request.form.get("response") or "").strip()
    req["responder"] = (request.form.get("responder") or "").strip() or None
    req["status"] = "answered"
    req["responded_utc"] = datetime.now(timezone.utc).isoformat()
    save_report(store)
    return render(THANKS)


if __name__ == "__main__":
    port = int(os.environ.get("PORT", "5000"))
    print("Data Challenger web (local prototype)")
    print(f"Open in your browser:  http://127.0.0.1:{port}")
    app.run(host="127.0.0.1", port=port, debug=False)
