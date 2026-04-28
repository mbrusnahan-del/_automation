"""
Email notification adapter for the automation sweep (v2.1.1).

Three modes, selected via KS_EMAIL_MODE env var:

    'stub'  (default) — write the email to _pending_email/*.eml and log.
                        Safe for dry-runs and first-deploy shakedown.
    'draft'           — drop a draft into Gmail using the MCP connector.
                        Used when the sweep runs inside Cowork and the Gmail
                        connector is authenticated.
    'send'            — SMTP send via Office 365. Requires:
                            KS_SMTP_HOST (default smtp.office365.com)
                            KS_SMTP_PORT (default 587)
                            KS_SMTP_USER
                            KS_SMTP_PASS (app password)

The sweep doesn't care which mode is active — it calls send_email() and moves
on. Switch modes by flipping the env var; no code change needed.
"""
from __future__ import annotations

import email.message
import logging
import os
import smtplib
from datetime import datetime, timezone
from pathlib import Path

log = logging.getLogger("notifications")

_AUTOMATION_DIR = Path(__file__).resolve().parent
_PENDING_DIR = _AUTOMATION_DIR / "_pending_email"


def send_email(to: list[str], subject: str, body: str, *, cc: list[str] | None = None) -> None:
    mode = os.environ.get("KS_EMAIL_MODE", "stub").lower()
    if mode == "stub":
        _write_stub(to, subject, body, cc=cc)
    elif mode == "draft":
        _gmail_draft(to, subject, body, cc=cc)
    elif mode == "send":
        _smtp_send(to, subject, body, cc=cc)
    elif mode == "notion":
        # Notion automations handle notifications; this mode is a no-op.
        # The sweep still flips the state flags, which is what Notion's
        # automation watches. We just skip emitting email.
        log.info("[notion] email skipped (Notion automations handle notify) → %s", to)
    else:
        raise RuntimeError(f"Unknown KS_EMAIL_MODE={mode!r}")


def _write_stub(to: list[str], subject: str, body: str, *, cc: list[str] | None) -> None:
    _PENDING_DIR.mkdir(parents=True, exist_ok=True)
    ts = datetime.now(timezone.utc).strftime("%Y%m%dT%H%M%SZ")
    slug = "".join(c if c.isalnum() or c in "-_" else "_" for c in (to[0] if to else "unknown"))
    path = _PENDING_DIR / f"{ts}__{slug}.eml"
    msg = email.message.EmailMessage()
    msg["Subject"] = subject
    msg["From"] = "automation@kingdomstructural.com"
    msg["To"] = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.set_content(body)
    path.write_bytes(bytes(msg))
    log.info("[stub] queued email → %s", path.name)


def _smtp_send(to: list[str], subject: str, body: str, *, cc: list[str] | None) -> None:
    """Send via Office 365 SMTP (smtp.office365.com:587 STARTTLS).

    Required env vars:
      KS_SMTP_USER  - mailbox the messages come from (e.g. mbrusnahan@kingdomstructural.com)
      KS_SMTP_PASS  - an App Password generated from Microsoft 365 security settings.
                      The mailbox account's login password will NOT work if the tenant
                      has MFA enabled (which it almost certainly does).

    Optional:
      KS_SMTP_HOST  - defaults to smtp.office365.com
      KS_SMTP_PORT  - defaults to 587
      KS_SMTP_FROM  - display name + address; defaults to bare KS_SMTP_USER

    Raises RuntimeError with a user-facing fix-it message on the common
    failure modes so the Automation Log entry is actionable.
    """
    host = os.environ.get("KS_SMTP_HOST", "smtp.office365.com")
    port = int(os.environ.get("KS_SMTP_PORT", "587"))
    user = os.environ.get("KS_SMTP_USER")
    pw   = os.environ.get("KS_SMTP_PASS")
    frm  = os.environ.get("KS_SMTP_FROM") or user

    if not user or not pw:
        raise RuntimeError(
            "KS_SMTP_USER / KS_SMTP_PASS not set. Generate an App Password at "
            "https://account.microsoft.com/security and export it as KS_SMTP_PASS. "
            "The mailbox password won't work with MFA enabled."
        )

    msg = email.message.EmailMessage()
    msg["Subject"] = subject
    msg["From"]    = frm
    msg["To"]      = ", ".join(to)
    if cc:
        msg["Cc"] = ", ".join(cc)
    msg.set_content(body)

    try:
        with smtplib.SMTP(host, port, timeout=30) as s:
            s.ehlo()
            s.starttls()
            s.ehlo()
            s.login(user, pw)
            s.send_message(msg)
    except smtplib.SMTPAuthenticationError as e:
        raise RuntimeError(
            f"SMTP authentication failed ({e.smtp_code} {e.smtp_error!r}). "
            "Most common cause: KS_SMTP_PASS is a normal mailbox password, "
            "but tenant requires an App Password. Generate one and retry."
        ) from e
    except smtplib.SMTPException as e:
        raise RuntimeError(f"SMTP send failed: {e}") from e

    log.info("[smtp] sent subj=%r → %s (cc=%s)", subject, to, cc or [])


# ────────────────────────────────────────────────────────────────────────
# CLI test harness  —  python notifications.py --to someone@domain.com
# Sends one test message in whatever mode KS_EMAIL_MODE is set to. Useful
# for verifying SMTP creds before flipping the scheduled sweep to live.
# ────────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    import argparse
    import logging as _logging

    # Auto-load .env when invoked as a script so creds don't need re-exporting.
    _env_path = Path(__file__).resolve().parent / ".env"
    if _env_path.exists():
        for _raw in _env_path.read_text(encoding="utf-8").splitlines():
            _line = _raw.strip()
            if not _line or _line.startswith("#") or "=" not in _line:
                continue
            _k, _, _v = _line.partition("=")
            _k = _k.strip(); _v = _v.strip()
            if len(_v) >= 2 and _v[0] == _v[-1] and _v[0] in "\"'":
                _v = _v[1:-1]
            os.environ.setdefault(_k, _v)

    _logging.basicConfig(level=_logging.INFO,
                         format="%(levelname)s %(name)s: %(message)s")
    ap = argparse.ArgumentParser(description="Test the notifications layer.")
    ap.add_argument("--to", required=True, help="Primary recipient email")
    ap.add_argument("--subject", default="KS automation smoke test")
    ap.add_argument("--body", default=("This is a test send from the Kingdom Structural "
                                       "automation stack. If you received this, the email "
                                       "path is wired correctly."))
    args = ap.parse_args()

    mode = os.environ.get("KS_EMAIL_MODE", "stub")
    print(f"KS_EMAIL_MODE = {mode!r}")
    if mode == "send":
        for var in ("KS_SMTP_USER", "KS_SMTP_PASS"):
            if not os.environ.get(var):
                print(f"  ! {var} not set — will fail")
    try:
        send_email(to=[args.to], subject=args.subject, body=args.body)
        print(f"✓ send_email() returned without error (mode={mode})")
    except Exception as e:
        print(f"✗ {type(e).__name__}: {e}")
        raise SystemExit(1)


def _gmail_draft(to: list[str], subject: str, body: str, *, cc: list[str] | None) -> None:
    """Placeholder — Cowork-integrated Gmail draft creation would call the MCP tool here.

    For a pure-Python deployment (cron, GitHub Actions, etc.), use 'send' mode with SMTP.
    """
    raise RuntimeError(
        "KS_EMAIL_MODE=draft is Cowork-only. Set KS_EMAIL_MODE=send with "
        "KS_SMTP_* creds for headless deployment, or 'stub' for dry-runs."
    )
