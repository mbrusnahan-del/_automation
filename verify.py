"""
Kingdom Structural automation — launch verification script.

Run this ONCE before flipping the scheduled sweep to live. It checks that
every external dependency the sweep needs is actually reachable with the
credentials in your environment, then performs a dry-run of the full sweep
so you can see what would change WITHOUT mutating anything.

Usage:
    python verify.py                    # full suite
    python verify.py --skip-smtp        # skip SMTP checks (e.g. creds not set up yet)
    python verify.py --skip-dry-run     # just check creds, don't hit the full sweep

Exit codes:
    0  — all green, safe to launch
    1  — one or more checks failed; see output for what to fix

Checks performed:
    [1] NOTION_TOKEN is set and authenticates                     (mandatory)
    [2] Notion integration has read access to the 3 databases     (mandatory)
    [3] Notion integration has write access                       (mandatory;
                                                                   write to
                                                                   Automation
                                                                   Log test row)
    [4] Engineer roster in config.py covers every ENGINEER
        assigned to a current Project                             (advisory)
    [5] SMTP creds present & can authenticate                     (mandatory
                                                                   for Jobs C/D)
    [6] SMTP test send to KS_SMTP_USER (i.e. send-to-self)        (advisory)
    [7] Dry-run of sweep.py across all 6 jobs                     (advisory)
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import smtplib
import sys
import traceback
from typing import Callable

log = logging.getLogger("verify")

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, _SCRIPT_DIR)

import config  # noqa: E402
import sweep   # noqa: E402


# ─── Small UI helpers ────────────────────────────────────────────────────


class _Ctx:
    def __init__(self) -> None:
        self.failures: list[str] = []
        self.warnings: list[str] = []

    def ok(self, msg: str) -> None:
        print(f"  ✓ {msg}")

    def warn(self, msg: str) -> None:
        print(f"  ! {msg}")
        self.warnings.append(msg)

    def fail(self, msg: str) -> None:
        print(f"  ✗ {msg}")
        self.failures.append(msg)


def _step(label: str, fn: Callable[[_Ctx], None], ctx: _Ctx) -> None:
    print(f"\n[{label}]")
    try:
        fn(ctx)
    except Exception as e:
        ctx.fail(f"{type(e).__name__}: {e}")
        if os.environ.get("KS_VERIFY_TRACE"):
            traceback.print_exc()


# ─── Check 1: NOTION_TOKEN authenticates ─────────────────────────────────


def check_notion_auth(ctx: _Ctx) -> None:
    if not os.environ.get("NOTION_TOKEN"):
        ctx.fail("NOTION_TOKEN not set. Export it from your Notion integration.")
        return
    try:
        resp = sweep._notion_request("GET", "/users/me")
    except sweep.NotionError as e:
        ctx.fail(f"Notion /users/me failed: {e}")
        return
    bot_name = resp.get("name") or resp.get("bot", {}).get("owner", {}).get("user", {}).get("name") or "?"
    ctx.ok(f"Authenticated as Notion bot: {bot_name!r}")


# ─── Check 2: integration sees the 3 databases ───────────────────────────


def check_databases_visible(ctx: _Ctx) -> None:
    targets = [
        ("Projects",       config.PROJECTS_DS_ID),
        ("Proposal Brief", config.BRIEF_DS_ID),
        ("Automation Log", config.AUTOMATION_LOG_DS_ID),
    ]
    for name, ds_id in targets:
        try:
            # Querying with an empty body + page_size=1 is the cheapest way
            # to confirm read access to a data source.
            sweep._notion_request(
                "POST", f"/data_sources/{ds_id}/query", {"page_size": 1}
            )
            ctx.ok(f"{name} DB reachable ({ds_id[:8]}…)")
        except sweep.NotionError as e:
            msg = str(e)
            if "object_not_found" in msg or "not_found" in msg:
                ctx.fail(f"{name} DB invisible to this integration. Share it with "
                         "the integration in Notion → DB → ⋯ → Add connections.")
            else:
                ctx.fail(f"{name} DB query failed: {msg[:180]}")


# ─── Check 3: can write to Automation Log ────────────────────────────────


def check_automation_log_write(ctx: _Ctx) -> None:
    from datetime import datetime, timezone
    now = datetime.now(timezone.utc).isoformat()
    try:
        resp = sweep.create_page(
            {"type": "data_source_id", "data_source_id": config.AUTOMATION_LOG_DS_ID},
            {
                "Log Entry": {"title": [{"type": "text", "text": {"content":
                    "verify.py smoke-test row (safe to delete)"}}]},
                "Job":       {"select": {"name": "A"}},
                "Outcome":   {"select": {"name": "Skipped"}},
                "Timestamp": {"date": {"start": now}},
            },
        )
        log_url = resp.get("url", "?")
        ctx.ok(f"Wrote test row to Automation Log: {log_url}")
        ctx.warn("Delete that test row manually when you're done verifying.")
    except sweep.NotionError as e:
        ctx.fail(f"Could not write to Automation Log: {e}")


# ─── Check 4: engineer roster covers live projects ───────────────────────


def check_engineer_roster(ctx: _Ctx) -> None:
    try:
        projects = sweep.query_data_source(
            config.PROJECTS_DS_ID,
            filter_obj={"property": config.ProjectProp.ENGINEER,
                        "people": {"is_not_empty": True}},
            page_size=100,
        )
    except sweep.NotionError as e:
        ctx.fail(f"Could not list Projects: {e}")
        return

    seen: set[str] = set()
    for proj in projects:
        for uid in sweep.people_ids(proj["properties"], config.ProjectProp.ENGINEER):
            seen.add(uid)

    missing = sorted(uid for uid in seen if uid not in config.ENGINEER_ROSTER)
    if not missing:
        ctx.ok(f"All {len(seen)} distinct engineers on live Projects are in the roster.")
        return
    for uid in missing:
        ctx.warn(f"Engineer UUID not in ENGINEER_ROSTER: {uid} — add to config.py")


# ─── Check 5 & 6: SMTP ───────────────────────────────────────────────────


def check_smtp_auth(ctx: _Ctx) -> None:
    # Skip SMTP checks entirely when the sweep is configured to use Notion
    # automations for notifications (no SMTP involved).
    mode = os.environ.get("KS_EMAIL_MODE", "stub").lower()
    if mode == "notion":
        ctx.ok("KS_EMAIL_MODE=notion — SMTP not used (Notion handles notify)")
        return
    host = os.environ.get("KS_SMTP_HOST", "smtp.office365.com")
    port = int(os.environ.get("KS_SMTP_PORT", "587"))
    user = os.environ.get("KS_SMTP_USER")
    pw   = os.environ.get("KS_SMTP_PASS")
    if not user or not pw:
        ctx.fail("KS_SMTP_USER / KS_SMTP_PASS not set. "
                 "Generate an App Password in Microsoft account security settings.")
        return
    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo(); s.starttls(); s.ehlo()
            s.login(user, pw)
        ctx.ok(f"SMTP {host}:{port} auth OK as {user}")
    except smtplib.SMTPAuthenticationError as e:
        ctx.fail(f"SMTP auth failed ({e.smtp_code}): {e.smtp_error!r}. "
                 "Use an App Password, not the mailbox password.")
    except smtplib.SMTPException as e:
        ctx.fail(f"SMTP connection error: {e}")


def check_smtp_self_send(ctx: _Ctx) -> None:
    mode = os.environ.get("KS_EMAIL_MODE", "stub").lower()
    if mode == "notion":
        ctx.ok("skipped (KS_EMAIL_MODE=notion)")
        return
    user = os.environ.get("KS_SMTP_USER")
    if not user:
        ctx.warn("Skipping self-send test (KS_SMTP_USER not set).")
        return
    os.environ["KS_EMAIL_MODE"] = "send"
    import notifications
    try:
        notifications.send_email(
            to=[user],
            subject="[KS automation verify.py] SMTP smoke test",
            body=("This is an automated test from verify.py confirming the SMTP "
                  "path is wired. If you received this, Jobs C and D can send "
                  "engineer and admin notifications. Safe to delete."),
        )
        ctx.ok(f"Test email sent to {user} — check your inbox.")
    except Exception as e:
        ctx.fail(f"Self-send failed: {e}")


# ─── Check 7: full sweep dry-run ─────────────────────────────────────────


def check_dry_run(ctx: _Ctx) -> None:
    print("    (running sweep.py --dry-run, no writes...)")
    rc = sweep.main(["--dry-run"])
    if rc == 0:
        ctx.ok("sweep.py --dry-run exited cleanly (rc=0)")
    else:
        ctx.fail(f"sweep.py --dry-run returned rc={rc}")


# ─── Orchestrator ────────────────────────────────────────────────────────


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="KS automation pre-launch verifier")
    ap.add_argument("--skip-smtp", action="store_true",
                    help="Skip SMTP checks (e.g. creds not set up yet)")
    ap.add_argument("--skip-dry-run", action="store_true",
                    help="Skip the full sweep dry-run")
    args = ap.parse_args(argv)

    logging.basicConfig(level=logging.INFO, format="%(levelname)s %(name)s: %(message)s")
    print("Kingdom Structural automation — launch verification\n")

    ctx = _Ctx()

    _step("1/7 Notion auth",             check_notion_auth,         ctx)
    _step("2/7 Notion databases visible", check_databases_visible,   ctx)
    _step("3/7 Notion write access",     check_automation_log_write, ctx)
    _step("4/7 Engineer roster coverage", check_engineer_roster,     ctx)

    if args.skip_smtp:
        print("\n[5/7 SMTP auth] skipped (--skip-smtp)")
        print("[6/7 SMTP self-send] skipped (--skip-smtp)")
    else:
        _step("5/7 SMTP auth",        check_smtp_auth,      ctx)
        _step("6/7 SMTP self-send",   check_smtp_self_send, ctx)

    if args.skip_dry_run:
        print("\n[7/7 Dry-run sweep] skipped (--skip-dry-run)")
    else:
        _step("7/7 Dry-run sweep",    check_dry_run,        ctx)

    print("\n" + "=" * 60)
    print(f"RESULT: {len(ctx.failures)} failure(s), {len(ctx.warnings)} warning(s)")
    if ctx.failures:
        print("\nFailures to fix before launching:")
        for f in ctx.failures:
            print(f"  ✗ {f}")
        return 1
    if ctx.warnings:
        print("\nWarnings (review but non-blocking):")
        for w in ctx.warnings:
            print(f"  ! {w}")
    print("\n✓ All checks passed — safe to switch the scheduled task to sweep.py.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
