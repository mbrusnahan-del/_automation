"""
PC-side HTTP server for the n8n hybrid architecture.

n8n handles pure-Notion jobs (A0, A02, C, D) itself. The two jobs that
need local filesystem access (A: create OneDrive folder, B: render Word
contract) are dispatched FROM n8n TO this server via HTTP POST.

The server is intentionally minimal:
- Python stdlib only (http.server, json)
- Token-authenticated (shared secret in header)
- Synchronous per request — fine because job volume is low
- Logs to webhook_server.log

Endpoints
---------
POST /create-folder
    body: { "project_id": "notion-page-uuid" }
    Reads Project properties from Notion (name, city, state, etc.),
    creates the OneDrive folder, populates the template tree, stamps
    the SharePoint URL on the Project, flips Folder=true and sets
    Status=Proposal Brief Pending. Returns { "ok": true, "path": "..." }
    or { "ok": false, "error": "..." }.

POST /render-contract
    body: { "brief_id": "notion-page-uuid" }
    Resolves every relation needed for the contract render, calls
    render_contract_package, stamps Rendered At and Status=Rendered on
    the Brief, bumps linked Project.Status to Contract Rendered.
    Returns { "ok": true, "contract": "path" } or an error.

POST /health
    Returns { "ok": true } — for n8n to sanity-check the server is up.

Auth
----
All requests must carry:
    X-KS-Token: <value of KS_WEBHOOK_TOKEN from .env>

Requests without a matching token return 401 and are NOT logged in detail
to avoid leaking which token attempts were made.

Usage
-----
    python n8n/webhook_server.py                      # defaults, port 8787
    python n8n/webhook_server.py --port 9000          # custom port
    python n8n/webhook_server.py --host 0.0.0.0       # listen on all ifaces

To auto-start on Windows login, drop `run_webhook_server.bat` into
shell:startup.
"""
from __future__ import annotations

import argparse
import json
import logging
import os
import sys
import traceback
from datetime import datetime, timezone
from http.server import BaseHTTPRequestHandler, HTTPServer

_SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
_PARENT = os.path.dirname(_SCRIPT_DIR)
sys.path.insert(0, _PARENT)

# Re-use the sweep module's Notion layer + job helpers. This guarantees
# the webhook server behaves identically to what sweep.py's Jobs A and B
# do — same filters, same error paths, same idempotency checks.
import config            # noqa: E402
import sweep             # noqa: E402
from render_proposal_package import (                               # noqa: E402
    render_contract_package, populate_from_template, ONEDRIVE_ROOT,
)

import urllib.parse      # noqa: E402

log = logging.getLogger("webhook_server")


# ────────────────────────────────────────────────────────────────────────
# Shared handlers — these are the same operations sweep.py's Jobs A and B
# perform, factored out so they can be called via HTTP too.
# ────────────────────────────────────────────────────────────────────────


def create_folder_for_project(project_id: str) -> dict:
    """Execute Job A for a single Project id."""
    project = sweep.get_page(project_id)
    props = project["properties"]

    # Gate re-check — in case the Project got edited between n8n firing
    # and this request arriving.
    if sweep.checkbox_val(props, config.ProjectProp.FOLDER):
        return {"ok": False, "error": "Project.Folder already true (idempotent skip)"}
    name = sweep.title_val(props, config.ProjectProp.NAME)
    city = sweep.text_val(props, config.ProjectProp.CITY)
    state = sweep.text_val(props, config.ProjectProp.STATE)
    if not (name and city and state):
        return {"ok": False, "error": "Project missing Name / City / State"}
    if not sweep.people_ids(props, config.ProjectProp.ENGINEER):
        return {"ok": False, "error": "Project has no ENGINEER"}
    if not sweep.people_ids(props, config.ProjectProp.ADMIN):
        return {"ok": False, "error": "Project has no Admin"}

    try:
        number, short = config.parse_project_name(name)
    except config.InvalidProjectNumberError as e:
        return {"ok": False, "error": str(e)}

    year = config.build_year_from_number(number)
    folder_name = config.build_folder_name(number, short, city, state)
    folder_path = os.path.join(ONEDRIVE_ROOT, year, folder_name)

    os.makedirs(folder_path, exist_ok=True)
    stats = populate_from_template(folder_path)
    if isinstance(stats, dict) and stats.get("error"):
        raise OSError(stats["error"])

    win_root = os.environ.get(
        "KS_ONEDRIVE_WIN_ROOT",
        "C:/Users/MichaelBrusnahan/OneDrive - Kingdom Structural LLC",
    )
    rel = os.path.relpath(folder_path, ONEDRIVE_ROOT).replace("\\", "/")
    folder_url = "file:///" + urllib.parse.quote(
        f"{win_root}/_Projects/{rel}", safe="/:"
    )

    sweep.update_page(project_id, {
        config.ProjectProp.FOLDER: {"checkbox": True},
        config.ProjectProp.SHAREPOINT_FOLDER: {"url": folder_url},
        config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.BRIEF_PENDING}},
    })
    sweep.log_row(
        "A", "Success",
        title=f"A · Folder created for {name} (via webhook)",
        project_id=project_id,
        details=f"Path: {folder_path}; {stats}",
    )
    return {"ok": True, "path": folder_path, "url": folder_url, "stats": stats}


def render_contract_for_brief(brief_id: str) -> dict:
    """Execute Job B for a single Brief id."""
    brief = sweep.get_page(brief_id)
    bprops = brief["properties"]
    if sweep.status_val(bprops, config.BriefProp.STATUS) == config.BriefStatus.RENDERED:
        return {"ok": False, "error": "Brief already Rendered (idempotent skip)"}

    project_rel = sweep.relation_ids(bprops, config.BriefProp.PROJECT)
    if not project_rel:
        return {"ok": False, "error": "Brief has no linked Project"}
    project_page = sweep.get_page(project_rel[0])
    if not sweep.checkbox_val(project_page["properties"], config.ProjectProp.FOLDER):
        return {"ok": False, "error": "Project.Folder=false; create folder first"}

    project_data = sweep._resolve_project_dict(project_page)
    client_data = sweep._resolve_client_dict(project_page)
    contact_data = sweep._resolve_contact_dict(project_page)
    engineer = sweep._resolve_engineer_dict(project_page)
    brief_overlay = sweep._resolve_brief_overlay_dict(brief)
    brief_body = sweep._fetch_brief_body_markdown(brief_id)
    fee_lines = sweep._fetch_fee_lines(brief_id)

    result = render_contract_package(
        project_data=project_data, client_data=client_data,
        contact_data=contact_data, engineer=engineer,
        brief=brief_overlay, brief_body=brief_body, fee_lines=fee_lines,
    )
    if isinstance(result, dict) and result.get("error"):
        raise RuntimeError(result["error"])
    contract_path = result.get("contract") if isinstance(result, dict) else result

    now = datetime.now(timezone.utc).isoformat()
    sweep.update_page(brief_id, {
        config.BriefProp.RENDERED_AT: {"date": {"start": now}},
        config.BriefProp.STATUS: {"select": {"name": config.BriefStatus.RENDERED}},
    })
    sweep.update_page(project_rel[0], {
        config.ProjectProp.STATUS: {"select": {"name": config.ProjectStatus.CONTRACT_RENDERED}},
    })
    sweep.log_row(
        "B", "Success",
        title=f"B · Contract rendered for {sweep.title_val(bprops, config.BriefProp.NAME)} (via webhook)",
        project_id=project_rel[0], brief_id=brief_id,
        details=f"Contract: {contract_path}",
    )
    return {"ok": True, "contract": contract_path}


# ────────────────────────────────────────────────────────────────────────
# HTTP handler
# ────────────────────────────────────────────────────────────────────────


class _KSHandler(BaseHTTPRequestHandler):

    server_version = "KSWebhook/1.0"

    def _send_json(self, status: int, payload: dict) -> None:
        body = json.dumps(payload).encode("utf-8")
        self.send_response(status)
        self.send_header("Content-Type", "application/json")
        self.send_header("Content-Length", str(len(body)))
        self.end_headers()
        self.wfile.write(body)

    def _check_auth(self) -> bool:
        expected = os.environ.get("KS_WEBHOOK_TOKEN")
        if not expected:
            log.error("KS_WEBHOOK_TOKEN env var not set — refusing all requests")
            self._send_json(500, {"ok": False, "error": "server misconfigured"})
            return False
        supplied = self.headers.get("X-KS-Token")
        if supplied != expected:
            self._send_json(401, {"ok": False, "error": "unauthorized"})
            return False
        return True

    def _read_json(self) -> dict | None:
        length = int(self.headers.get("Content-Length") or 0)
        if length <= 0:
            return {}
        try:
            raw = self.rfile.read(length)
            return json.loads(raw.decode("utf-8"))
        except (ValueError, UnicodeDecodeError) as e:
            self._send_json(400, {"ok": False, "error": f"invalid JSON: {e}"})
            return None

    def log_message(self, fmt: str, *args) -> None:
        # Route through logging so it ends up in webhook_server.log
        log.info("%s - %s", self.address_string(), fmt % args)

    def do_POST(self) -> None:
        if self.path == "/health":
            # Health check — no auth required, just "is the server alive"
            self._send_json(200, {"ok": True, "time": datetime.now(timezone.utc).isoformat()})
            return

        if not self._check_auth():
            return

        body = self._read_json()
        if body is None:
            return

        try:
            if self.path == "/create-folder":
                project_id = (body or {}).get("project_id")
                if not project_id:
                    self._send_json(400, {"ok": False, "error": "missing project_id"})
                    return
                result = create_folder_for_project(project_id)
                self._send_json(200 if result.get("ok") else 422, result)
                return

            if self.path == "/render-contract":
                brief_id = (body or {}).get("brief_id")
                if not brief_id:
                    self._send_json(400, {"ok": False, "error": "missing brief_id"})
                    return
                result = render_contract_for_brief(brief_id)
                self._send_json(200 if result.get("ok") else 422, result)
                return

            self._send_json(404, {"ok": False, "error": f"no handler for {self.path}"})

        except Exception as e:
            log.exception("handler crashed for %s", self.path)
            self._send_json(500, {"ok": False, "error": f"{type(e).__name__}: {e}",
                                  "trace": traceback.format_exc()[-2000:]})

    def do_GET(self) -> None:
        # GET-only handler for health checks and troubleshooting.
        if self.path == "/health":
            self._send_json(200, {"ok": True, "time": datetime.now(timezone.utc).isoformat()})
            return
        self._send_json(405, {"ok": False, "error": "method not allowed"})


def main(argv: list[str]) -> int:
    ap = argparse.ArgumentParser(description="KS webhook server for n8n dispatch.")
    ap.add_argument("--host", default="127.0.0.1",
                    help="Bind address. Use 0.0.0.0 to accept LAN connections.")
    ap.add_argument("--port", type=int, default=8787)
    args = ap.parse_args(argv)

    # Log to both stdout and a persistent file
    log_path = os.path.join(_SCRIPT_DIR, "webhook_server.log")
    logging.basicConfig(
        level=logging.INFO,
        format="%(asctime)s %(levelname)s %(name)s: %(message)s",
        handlers=[
            logging.StreamHandler(sys.stdout),
            logging.FileHandler(log_path, encoding="utf-8"),
        ],
    )

    # Auto-load .env next to sweep.py so NOTION_TOKEN + KS_WEBHOOK_TOKEN
    # are available to the request handlers.
    sweep._load_dotenv()

    if not os.environ.get("KS_WEBHOOK_TOKEN"):
        log.warning(
            "KS_WEBHOOK_TOKEN not set in .env — requests will be rejected. "
            "Add a line like KS_WEBHOOK_TOKEN=choose-a-random-string to .env."
        )

    httpd = HTTPServer((args.host, args.port), _KSHandler)
    log.info("KS webhook server listening on http://%s:%d", args.host, args.port)
    log.info("Health check:  curl http://%s:%d/health", args.host, args.port)
    try:
        httpd.serve_forever()
    except KeyboardInterrupt:
        log.info("shutting down")
        httpd.server_close()
    return 0


if __name__ == "__main__":
    raise SystemExit(main(sys.argv[1:]))
