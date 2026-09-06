# Data Challenger

Drop in a data file (CSV / Excel) and get back a structured **review**: what looks
inconsistent, the evidence, and the strategic question a named colleague should
answer. Findings are assigned to reviewers, each question gets a unique response
link, answers flow back into the report, and everything can be exported as a PDF.

## Features

- **Auto-detection of columns** (date, amount, category) — column names may vary.
- **Domain packs**: HR / employee-turnover data is recognised by its column
  signatures (attrition, employee id, salary, tenure...) and challenged with
  HR-specific questions (overtime vs attrition, early leavers, regrettable
  attrition, pay gap, hotspot departments/roles...).
- **Department folders**: reviews are organised into folders
  (Logistics / Accounting / Human Resources) with a **day-by-day** view of the
  comments received and the state of each document as of that date.
- **Assignment & response loop**: every finding is routed to a reviewer
  (master data with routing rules), gets a tokenised expiring response link,
  and can be delivered via **Slack DM** (`tools/challenge_notify.py`,
  dry-run mode without a token).
- **PDF export** with the responses, cells sized to fit each answer.
- **Editable reviewers**: add or edit reviewers directly in the web UI.

## Project layout

```
tools/
  data_challenge.py       rule engine: column detection, checks, profiling
  data_challenge_pdf.py   PDF renderer (fpdf2)
  challenge_workflow.py   assignment layer: routing, tokens, request outbox
  challenge_notify.py     Slack DM notifier (Level 1, stdlib only)
  hr_turnover.py          HR / employee-turnover domain pack
  users_master.json       reviewers master data + routing rules
  webapp/app.py           Flask web app (local prototype)
sample_data/
  Kaggle_Employee_Turnover.csv   IBM HR Analytics sample (Kaggle)
```

## Run

```bash
pip install -r requirements.txt
python tools/webapp/app.py
# open http://127.0.0.1:5000
```

Runtime data (uploads, reports, tokens) is stored under `tools/webapp/data/`
and is intentionally **not** part of the repository (see `.gitignore`).

## Slack delivery (optional)

Create a Slack app with bot scopes `chat:write`, `users:read.email`, `im:write`
and export `SLACK_BOT_TOKEN` before starting the web app. Without the token the
"Send questions via Slack" button runs in dry-run mode (nothing is sent; the
flow is stamped so you can test the loop).

## Status

Local prototype: no authentication, no hosting, links resolve on localhost.
Phase 2 (hosted response form + real delivery) is the next step.
