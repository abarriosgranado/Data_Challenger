#!/usr/bin/env python3
"""
Data Challenge - Level 1 notifier (Slack DM)
============================================

Sends each pending finding to its assignee as a Slack direct message containing
the question, the evidence and the tokenised response link. The reviewer still
answers in the web app (/respond); this module only handles delivery.

Modes
-----
- LIVE     : set the SLACK_BOT_TOKEN env var (bot scopes needed:
             chat:write, users:read.email, im:write). Assignees are matched
             by email via users.lookupByEmail.
- DRY-RUN  : no token present. Nothing is sent; the message that WOULD be sent
             is printed/logged and the request is stamped notified_via="dry-run"
             so the UI can show that the loop was exercised.

Per-request fields written into the report store:
    notified_via   : "slack" | "dry-run"
    notified_utc   : ISO timestamp of the (attempted) notification
    notify_error   : last delivery error, or null

Standalone usage (outbox file produced by challenge_workflow.py or the webapp):
    python tools/challenge_notify.py path/to/REPORT_requests.json

No third-party dependencies: uses urllib from the standard library.
"""
from __future__ import annotations

import json
import os
import sys
import urllib.request
from datetime import datetime, timezone

SLACK_API = "https://slack.com/api"


# --------------------------------------------------------------------------- #
# Slack Web API (stdlib only)
# --------------------------------------------------------------------------- #
def _slack_call(method: str, token: str, payload: dict) -> dict:
    req = urllib.request.Request(
        f"{SLACK_API}/{method}",
        data=json.dumps(payload).encode("utf-8"),
        headers={
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json; charset=utf-8",
        },
        method="POST",
    )
    with urllib.request.urlopen(req, timeout=15) as resp:
        data = json.loads(resp.read().decode("utf-8"))
    if not data.get("ok"):
        raise RuntimeError(f"Slack {method} failed: {data.get('error', 'unknown_error')}")
    return data


def _lookup_user_id(token: str, email: str, cache: dict) -> str:
    if email in cache:
        return cache[email]
    data = _slack_call("users.lookupByEmail", token, {"email": email})
    uid = data["user"]["id"]
    cache[email] = uid
    return uid


# --------------------------------------------------------------------------- #
# Message
# --------------------------------------------------------------------------- #
def build_message(req: dict, source_file: str) -> str:
    return (
        f":mag: *Data Challenge - review question #{req['finding_no']}* "
        f"({req['severity']} severity)\n"
        f"*File:* {source_file}\n"
        f"*Issue area:* {req['issue_area']}\n"
        f"*What we detected:* {req['observation']}\n"
        f"*Evidence:* {req['evidence']}\n\n"
        f"*Question for you:* {req['question']}\n\n"
        f":point_right: Please answer here (link expires {req['expires_utc'][:10]}):\n"
        f"{req['response_link']}"
    )


# --------------------------------------------------------------------------- #
# Core
# --------------------------------------------------------------------------- #
def notify_report(store: dict, resend: bool = False, log=print) -> dict:
    """Send (or dry-run) a Slack DM for every pending, un-notified request.

    Mutates the store in place and returns a summary dict.
    Caller is responsible for persisting the store afterwards.
    """
    token = os.environ.get("SLACK_BOT_TOKEN", "").strip()
    dry_run = not token
    now = datetime.now(timezone.utc).isoformat()
    cache: dict = {}
    sent = skipped = failed = 0

    for req in store.get("requests", []):
        if req.get("status") == "answered":
            skipped += 1
            continue
        if req.get("notified_utc") and not resend:
            skipped += 1
            continue
        if not req.get("assignee_email"):
            req["notify_error"] = "no assignee email"
            failed += 1
            continue

        msg = build_message(req, store.get("source_file", "?"))
        if dry_run:
            log(f"[dry-run] DM to {req['assignee_name']} <{req['assignee_email']}>:\n{msg}\n")
            req["notified_via"] = "dry-run"
            req["notified_utc"] = now
            req["notify_error"] = None
            sent += 1
            continue

        try:
            uid = _lookup_user_id(token, req["assignee_email"], cache)
            channel = _slack_call("conversations.open", token, {"users": uid})["channel"]["id"]
            _slack_call("chat.postMessage", token, {"channel": channel, "text": msg})
            req["notified_via"] = "slack"
            req["notified_utc"] = now
            req["notify_error"] = None
            sent += 1
        except Exception as e:  # keep going; record the error per request
            req["notify_error"] = str(e)
            failed += 1
            log(f"[error] {req['assignee_email']}: {e}")

    return {"mode": "dry-run" if dry_run else "slack",
            "sent": sent, "skipped": skipped, "failed": failed}


# --------------------------------------------------------------------------- #
# CLI
# --------------------------------------------------------------------------- #
def main(argv=None):
    argv = argv if argv is not None else sys.argv[1:]
    if len(argv) < 1:
        print(__doc__)
        return 1
    path = argv[0]
    resend = "--resend" in argv
    with open(path, encoding="utf-8") as fh:
        store = json.load(fh)
    summary = notify_report(store, resend=resend)
    with open(path, "w", encoding="utf-8") as fh:
        json.dump(store, fh, indent=2, ensure_ascii=False)
    print(f"Mode: {summary['mode']}  sent: {summary['sent']}  "
          f"skipped: {summary['skipped']}  failed: {summary['failed']}")
    return 0


if __name__ == "__main__":
    sys.exit(main())
