"""
Kingdom Structural — Notion Project Automation render service.

FastAPI service deployed to Cloud Run (or Fly.io/Railway/Render).
n8n hits POST /process with { page_id }; we do everything from there:

  1. Fetch full project from Notion
  2. Decide: create_folder | render_proposal | skip
  3. For create_folder:
       - Copy template folder to OneDrive via Microsoft Graph
       - PATCH Notion page: Folder = true, SharePoint Folder URL
  4. For render_proposal:
       - Fetch linked Client + first linked Contact
       - Render contract + fee memo .docx via the existing python-docx pipeline
       - Upload both to {project folder}/Contracts/ via Microsoft Graph
       - DO NOT change Engineering Status (partner advances it manually)

The rendering code itself is copied from /mnt/_Projects/_automation/.
The service is fully stateless; it re-reads comps_cache.json each request
(baked into the container at build time; the excel-comps-sync cron still
owns updating comps in OneDrive, so at deploy time you snapshot it into
this container — or, optionally, point COMPS_CACHE_URL at a hosted copy).
"""
from __future__ import annotations

import base64
import io
import logging
import os
import re
import time
import urllib.parse
from datetime import date
from typing import Any, Dict, List, Optional, Tuple

import requests
from fastapi import FastAPI, HTTPException, Header, Request
from pydantic import BaseModel

# python-docx pipeline (vendored alongside main.py; see requirements.txt + Dockerfile)
from render_proposal_package import (
    render_proposal_package,
    project_folder_name,
)

# ---------------------------------------------------------------------------
# Config — all from env vars
# ---------------------------------------------------------------------------
NOTION_TOKEN = os.environ["NOTION_TOKEN"]
NOTION_API = "https://api.notion.com/v1"
NOTION_VERSION = "2022-06-28"

# Notion data source IDs (per CLAUDE.md)
PROJECTS_DS_ID  = "262b73dc-460e-8137-b3bb-000b62103b15"
BRIEFS_DS_ID    = "6b64658f-5fb3-4329-b02c-3ac3cb8a0828"
FEE_LINES_DS_ID = "e05398ef-c999-477b-9420-3ccd4976730a"


GRAPH_CLIENT_ID = os.environ["GRAPH_CLIENT_ID"]
GRAPH_CLIENT_SECRET = os.environ["GRAPH_CLIENT_SECRET"]
GRAPH_TENANT_ID = os.environ["GRAPH_TENANT_ID"]
GRAPH_USER_UPN = os.environ["GRAPH_USER_UPN"]  # e.g. mbrusnahan@kingdomstructural.com
GRAPH_API = "https://graph.microsoft.com/v1.0"

# Windows-facing path used to build the SharePoint Folder URL that gets
# written back to Notion. The file:/// links open the local OneDrive copy.
ONEDRIVE_HOST_ROOT = (
    "file:///C:/Users/MichaelBrusnahan/"
    "OneDrive%20-%20Kingdom%20Structural%20LLC/_Projects"
)

# Location of the canonical template folder in OneDrive, relative to the
# user's drive root. Keep in sync with the cron task.
TEMPLATE_DRIVE_PATH = "/_Projects/2026/_26001001 New 2026 Job Name - City, STATE"
PROJECTS_DRIVE_PATH = "/_Projects"

# Optional shared secret for n8n to include in X-Kingdom-Auth header.
SHARED_SECRET = os.environ.get("SHARED_SECRET", "")

# Async copy polling — how long we wait for Graph's folder-copy job to finish.
COPY_POLL_TIMEOUT_SEC = 120
COPY_POLL_INTERVAL_SEC = 3

logging.basicConfig(level=logging.INFO, format="%(asctime)s %(levelname)s %(message)s")
log = logging.getLogger("kingdom-automation")

app = FastAPI(title="Kingdom Structural Notion Automation")


# ---------------------------------------------------------------------------
# Request/response models
# ---------------------------------------------------------------------------
class ProcessRequest(BaseModel):
    page_id: str
    source: Optional[str] = "webhook"


class ProcessResponse(BaseModel):
    ok: bool
    action: str
    page_id: str
    job_number: Optional[str] = None
    folder_name: Optional[str] = None
    skip_reason: Optional[str] = None
    contract_path: Optional[str] = None
    memo_path: Optional[str] = None
    message: str


# ---------------------------------------------------------------------------
# Notion helpers
# ---------------------------------------------------------------------------
def notion_headers() -> Dict[str, str]:
    return {
        "Authorization": f"Bearer {NOTION_TOKEN}",
        "Notion-Version": NOTION_VERSION,
        "Content-Type": "application/json",
    }


def notion_get_page(page_id: str) -> Dict[str, Any]:
    r = requests.get(f"{NOTION_API}/pages/{page_id}", headers=notion_headers(), timeout=30)
    r.raise_for_status()
    return r.json()


def notion_patch_page(page_id: str, properties: Dict[str, Any]) -> None:
    r = requests.patch(
        f"{NOTION_API}/pages/{page_id}",
        headers=notion_headers(),
        json={"properties": properties},
        timeout=30,
    )
    r.raise_for_status()


def _text_plain(prop: Optional[dict]) -> str:
    if not prop:
        return ""
    if prop.get("title"):
        return "".join(t.get("plain_text", "") for t in prop["title"])
    if prop.get("rich_text"):
        return "".join(t.get("plain_text", "") for t in prop["rich_text"])
    if prop.get("url"):
        return prop["url"] or ""
    return ""


def _select_name(prop: Optional[dict]) -> str:
    if prop and prop.get("select"):
        return prop["select"].get("name", "") or ""
    return ""


def _status_name(prop: Optional[dict]) -> str:
    if prop and prop.get("status"):
        return prop["status"].get("name", "") or ""
    return ""


def _checkbox(prop: Optional[dict]) -> bool:
    return bool(prop and prop.get("checkbox"))


def _relation_ids(prop: Optional[dict]) -> List[str]:
    if not prop or not prop.get("relation"):
        return []
    return [r["id"] for r in prop["relation"] if "id" in r]


def _number(prop: Optional[dict]) -> float:
    if prop and isinstance(prop.get("number"), (int, float)):
        return float(prop["number"])
    return 0.0


def parse_project(page: Dict[str, Any]) -> Dict[str, Any]:
    """Flatten the Notion project page into the dict shape the renderer expects."""
    props = page.get("properties", {})
    project_name_raw = _text_plain(props.get("Project Name")).replace("\n", " ").strip()
    m = re.match(r"^(\d{8})[\s\-]+(.+)$", project_name_raw)
    job_number = m.group(1) if m else ""
    project_name_clean = m.group(2).strip() if m else project_name_raw

    return {
        "page_id": page.get("id"),
        "job_number": job_number,
        "year": ("20" + job_number[:2]) if job_number else "",
        "project_name_raw": project_name_raw,
        "project_name_clean": project_name_clean,
        "city": _text_plain(props.get("City")).strip(),
        "state": _text_plain(props.get("State")).strip(),
        "engineering_status": _status_name(props.get("Engineering Status")),
        "folder_checked": _checkbox(props.get("Folder")),
        "project_type": _select_name(props.get("Project Type")),
        "icc_code_year": _select_name(props.get("ICC Code Year")),
        "jurisdiction": _text_plain(props.get("Jurisdiction")),
        "scope_description": _text_plain(props.get("Scope Description")),
        "project_street": _text_plain(props.get("Project Street")),
        "reimbursables_treatment": _select_name(props.get("Reimbursables Treatment")),
        "approx_sf": _number(props.get("Approx. Structural SF")),
        "sharepoint_folder": _text_plain(props.get("SharePoint Folder")),
        "client_ids": _relation_ids(props.get("Client")),
        "contact_ids": _relation_ids(props.get("Project Contact")),
    }


def parse_client(page: Dict[str, Any]) -> Dict[str, Any]:
    props = page.get("properties", {})
    return {
        "id": page.get("id"),
        "Name": _text_plain(props.get("Name")),
        "Address": _text_plain(props.get("Address")),
        "The Guy": _select_name(props.get("The Guy")),
        "contact_ids": _relation_ids(props.get("Contacts")),
    }


def parse_contact(page: Dict[str, Any]) -> Dict[str, Any]:
    props = page.get("properties", {})
    return {
        "id": page.get("id"),
        "Contact Name": _text_plain(props.get("Contact Name")),
        "Email": _text_plain(props.get("Email")),
        "Phone": _text_plain(props.get("Phone")),
        "Phone (Mobile)": _text_plain(props.get("Phone (Mobile)")),
        "Role": _text_plain(props.get("Role")),
    }

def parse_brief(page: Dict[str, Any]) -> Dict[str, Any]:
    """Flatten a Proposal Brief page into the overlay dict the renderer expects.
    Per 2026-04-22 schema rework: City, State, ICC Code Year, Jurisdiction,
    Project Street, Reimbursables Treatment all live on the Brief now."""
    props = page.get("properties", {})
    return {
        "id": page.get("id"),
        "City": _text_plain(props.get("City")).strip(),
        "State": _text_plain(props.get("State")).strip(),
        "ICC Code Year": _select_name(props.get("ICC Code Year")),
        "Jurisdiction": _text_plain(props.get("Jurisdiction")).strip(),
        "Project Street": _text_plain(props.get("Project Street")).strip(),
        "Reimbursables Treatment": _select_name(props.get("Reimbursables Treatment")),
        "Project Type": _select_name(props.get("Project Type")),  # rollup, display only
        "Approx. Structural SF": _number(props.get("Approx. Structural SF")),
        "Scope Bullets": [],  # populated by brief body parser later
        "project_ids": _relation_ids(props.get("Parent item")) or _relation_ids(props.get("Project")),
        "Status": _status_name(props.get("Status")) or _select_name(props.get("Status")),
    }


def fetch_brief_for_project(project_page_id: str) -> Optional[Dict[str, Any]]:
    """Find the Proposal Brief linked to this project, if any. Returns parsed brief dict."""
    page = notion_get_page(project_page_id)
    brief_ids = _relation_ids(page.get("properties", {}).get("Proposal Brief"))
    if not brief_ids:
        return None
    brief_page = notion_get_page(brief_ids[0])
    return parse_brief(brief_page)


def fetch_project_for_brief(brief_page: Dict[str, Any]) -> Optional[Dict[str, Any]]:
    """Find the Project linked to this brief via Parent item / Project relation."""
    props = brief_page.get("properties", {})
    proj_ids = _relation_ids(props.get("Parent item")) or _relation_ids(props.get("Project"))
    if not proj_ids:
        return None
    project_page = notion_get_page(proj_ids[0])
    return parse_project(project_page)




# ---------------------------------------------------------------------------
# Microsoft Graph helpers (app-only / client credentials)
# ---------------------------------------------------------------------------
_token_cache: Dict[str, Any] = {"token": None, "exp": 0}


def graph_token() -> str:
    now = time.time()
    if _token_cache["token"] and _token_cache["exp"] > now + 60:
        return _token_cache["token"]
    url = f"https://login.microsoftonline.com/{GRAPH_TENANT_ID}/oauth2/v2.0/token"
    r = requests.post(
        url,
        data={
            "client_id": GRAPH_CLIENT_ID,
            "client_secret": GRAPH_CLIENT_SECRET,
            "scope": "https://graph.microsoft.com/.default",
            "grant_type": "client_credentials",
        },
        timeout=30,
    )
    r.raise_for_status()
    body = r.json()
    _token_cache["token"] = body["access_token"]
    _token_cache["exp"] = now + int(body.get("expires_in", 3600))
    return _token_cache["token"]


def graph_headers() -> Dict[str, str]:
    return {"Authorization": f"Bearer {graph_token()}"}


def _drive_root_url() -> str:
    """Root URL for the target OneDrive.

    GRAPH_USER_UPN must be a user in the tenant with an OneDrive provisioned
    and whose drive the service has been granted access to (Files.ReadWrite.All
    app permission, admin-consented).
    """
    return f"{GRAPH_API}/users/{GRAPH_USER_UPN}/drive"


def _path_segment(path: str) -> str:
    """URL-encode a drive path for /root:/<path>: style addressing."""
    # Each path component gets percent-encoded, but we preserve "/" separators.
    parts = [urllib.parse.quote(p, safe="") for p in path.strip("/").split("/")]
    return "/" + "/".join(parts) if parts else ""


def graph_path_exists(path: str) -> bool:
    url = f"{_drive_root_url()}/root:{_path_segment(path)}"
    r = requests.get(url, headers=graph_headers(), timeout=30)
    return r.status_code == 200


def graph_get_item(path: str) -> Optional[Dict[str, Any]]:
    url = f"{_drive_root_url()}/root:{_path_segment(path)}"
    r = requests.get(url, headers=graph_headers(), timeout=30)
    if r.status_code == 404:
        return None
    r.raise_for_status()
    return r.json()


def graph_copy_template(dest_parent_path: str, new_name: str) -> None:
    """Recursively copy the canonical template folder into dest_parent_path/new_name.

    This is Graph's native async copy. We submit and then poll the monitor URL
    until completion so the caller has a synchronous-feeling API.
    """
    url = f"{_drive_root_url()}/root:{_path_segment(TEMPLATE_DRIVE_PATH)}:/copy"
    body = {
        "parentReference": {"path": f"/drive/root:{_path_segment(dest_parent_path)}"},
        "name": new_name,
    }
    r = requests.post(url, headers={**graph_headers(), "Content-Type": "application/json"},
                      json=body, timeout=30)
    if r.status_code not in (200, 202):
        raise HTTPException(status_code=502,
                            detail=f"Graph copy failed: {r.status_code} {r.text}")
    monitor = r.headers.get("Location")
    if not monitor:
        # Synchronous completion (rare); we're done.
        return
    # Poll the monitor URL until complete (or we give up).
    deadline = time.time() + COPY_POLL_TIMEOUT_SEC
    while time.time() < deadline:
        m = requests.get(monitor, timeout=30)
        if m.status_code == 200:
            mj = m.json()
            status = mj.get("status")
            if status in ("completed", "succeeded"):
                return
            if status == "failed":
                raise HTTPException(
                    status_code=502,
                    detail=f"Graph copy failed: {mj.get('error', {}).get('message','?')}",
                )
        elif m.status_code in (201, 303):
            return  # done
        time.sleep(COPY_POLL_INTERVAL_SEC)
    raise HTTPException(status_code=504, detail="Graph copy timed out")


def graph_upload_file(dest_path: str, data: bytes, mime_type: str) -> Dict[str, Any]:
    """Upload a file ≤ 4MB via a single PUT. .docx are always well under this."""
    url = f"{_drive_root_url()}/root:{_path_segment(dest_path)}:/content"
    headers = {**graph_headers(), "Content-Type": mime_type,
               "Content-Length": str(len(data))}
    r = requests.put(url, headers=headers, data=data, timeout=120)
    if r.status_code not in (200, 201):
        raise HTTPException(status_code=502,
                            detail=f"Graph upload failed: {r.status_code} {r.text}")
    return r.json()


# ---------------------------------------------------------------------------
# Actions
# ---------------------------------------------------------------------------
DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"


def do_create_folder(project: Dict[str, Any]) -> Tuple[str, str]:
    """Copy template into /_Projects/{year}/{folder_name} and flip Notion flag.

    Returns (folder_name, sharepoint_url).
    """
    job = project["job_number"]
    year = project["year"]
    folder_name, _ = project_folder_name(
        job,
        project["project_name_raw"],
        project["city"],
        project["state"],
    )

    year_path = f"{PROJECTS_DRIVE_PATH}/{year}"
    new_folder_path = f"{year_path}/{folder_name}"

    # If the folder already exists on-disk, skip the copy and just reconcile
    # the Notion flag. Mirrors the "mismatch: leave alone" rule in the cron
    # task — except here we DO set the flag to true, because we're the
    # authoritative sync now and a webhook fired specifically for this page.
    if not graph_path_exists(new_folder_path):
        graph_copy_template(dest_parent_path=year_path, new_name=folder_name)

    sp_url = (
        f"{ONEDRIVE_HOST_ROOT}/{year}/"
        f"{urllib.parse.quote(folder_name, safe='')}"
    )
    notion_patch_page(
        project["page_id"],
        {
            "Folder": {"checkbox": True},
            "SharePoint Folder": {"url": sp_url},
        },
    )
    return folder_name, sp_url


def _idempotency_skip(project: Dict[str, Any], folder_name: str) -> bool:
    """True if {job#} Contract or Fee Analysis file already exists in Contracts/."""
    contracts_path = (
        f"{PROJECTS_DRIVE_PATH}/{project['year']}/{folder_name}/Contracts"
    )
    item = graph_get_item(contracts_path)
    if not item:
        return False
    # List children and look for existing
    children_url = f"{_drive_root_url()}/items/{item['id']}/children?$top=200"
    r = requests.get(children_url, headers=graph_headers(), timeout=30)
    if r.status_code != 200:
        return False
    for c in r.json().get("value", []):
        name = c.get("name", "")
        if name.startswith(project["job_number"]) and ("Contract" in name or "Fee Analysis" in name):
            return True
    return False


def do_render_proposal(project: Dict[str, Any], brief: Optional[Dict[str, Any]] = None) -> Dict[str, str]:
    """Render contract + memo, upload to OneDrive. Returns paths of uploaded files."""
    if not project["client_ids"]:
        raise HTTPException(status_code=422, detail="project has no Client relation")

    # Pull client and first contact
    client_page = notion_get_page(project["client_ids"][0])
    client = parse_client(client_page)

    contact_id = None
    if project["contact_ids"]:
        contact_id = project["contact_ids"][0]
    elif client["contact_ids"]:
        contact_id = client["contact_ids"][0]

    contact: Dict[str, Any]
    if contact_id:
        contact = parse_contact(notion_get_page(contact_id))
    else:
        contact = {"Contact Name": "", "Email": "", "Phone": "", "Phone (Mobile)": "", "Role": ""}

    folder_name, proj_short = project_folder_name(
        project["job_number"],
        project["project_name_raw"],
        project["city"],
        project["state"],
        client_name=client["Name"],
    )

    if _idempotency_skip(project, folder_name):
        return {"status": "already_rendered", "folder_name": folder_name}

    # Auto-fetch brief if not passed in
    if brief is None:
        brief = fetch_brief_for_project(project["page_id"])
        if brief:
            log.info("auto-fetched brief %s for project %s", brief["id"], project["page_id"])

    # Build dicts in the shape render_proposal_package expects.
    project_data = {
        "Project Name": project["project_name_raw"],
        "City": project["city"],
        "State": project["state"],
        "Engineering Status": project["engineering_status"],
        "Project Type": project["project_type"],
        "ICC Code Year": project["icc_code_year"],
        "Jurisdiction": project["jurisdiction"],
        "Scope Description": project["scope_description"],
        "Project Street": project["project_street"],
        "Reimbursables Treatment": project["reimbursables_treatment"],
        "Approx. Structural SF": project["approx_sf"],
        "SharePoint Folder": project["sharepoint_folder"],
    }
    client_data = {
        "Name": client["Name"],
        "Address": client["Address"],
        "The Guy": client["The Guy"],
    }
    contact_data = {
        "Contact Name": contact.get("Contact Name", ""),
        "Email": contact.get("Email", ""),
        "Phone": contact.get("Phone", ""),
        "Phone (Mobile)": contact.get("Phone (Mobile)", ""),
        "Role": contact.get("Role", ""),
    }
    project_for_memo = {
        "project_name": project["project_name_clean"],
        "ks_job_number": project["job_number"],
        "project_type": project["project_type"],
        "approx_sf": project["approx_sf"],
        "jurisdiction": project["jurisdiction"],
        "location": ", ".join([x for x in [project["city"], project["state"]] if x]),
        "scope_description": project["scope_description"],
        "partner": client["The Guy"],
        "client_company_name": client["Name"],
        "engineering_status": project["engineering_status"],
    }

    # render_proposal_package writes to a local folder. We redirect ONEDRIVE_ROOT
    # at import time (see bootstrap below) to a tmpfs path so rendering happens
    # in-container, then we ship the two resulting .docx files to Graph.
    import render_proposal_package as rpp
    import tempfile, shutil, pathlib

    with tempfile.TemporaryDirectory() as tmp:
        # Make a fake "OneDrive root" the renderer can write into. It expects
        # a project folder under {ONEDRIVE_ROOT}/{year}/{folder_name}/Contracts/.
        fake_year = pathlib.Path(tmp) / project["year"]
        fake_folder = fake_year / folder_name
        (fake_folder / "Contracts").mkdir(parents=True, exist_ok=True)
        rpp.ONEDRIVE_ROOT = str(pathlib.Path(tmp))

        # Also point the template folder at an empty dir so populate_from_template
        # no-ops inside the container (the real copy already happened in Graph).
        empty_template = pathlib.Path(tmp) / "_empty_template"
        empty_template.mkdir(exist_ok=True)
        rpp.TEMPLATE_PROJECT_FOLDER = str(empty_template)

        result = render_proposal_package(
            project_data, client_data, contact_data, project_for_memo,
            brief=brief,
        )
        if "error" in result:
            raise HTTPException(status_code=500, detail=result["error"])

        contract_local = pathlib.Path(result["contract"])
        memo_local = pathlib.Path(result["memo"])

        contract_bytes = contract_local.read_bytes()
        memo_bytes = memo_local.read_bytes()

    # Upload both to the real OneDrive under {Projects}/{year}/{folder}/Contracts/
    contracts_drive_path = f"{PROJECTS_DRIVE_PATH}/{project['year']}/{folder_name}/Contracts"
    contract_drive_path = f"{contracts_drive_path}/{contract_local.name}"
    memo_drive_path = f"{contracts_drive_path}/{memo_local.name}"

    graph_upload_file(contract_drive_path, contract_bytes, DOCX_MIME)
    graph_upload_file(memo_drive_path, memo_bytes, DOCX_MIME)

    return {
        "status": "rendered",
        "folder_name": folder_name,
        "contract_path": contract_drive_path,
        "memo_path": memo_drive_path,
    }


# ---------------------------------------------------------------------------
# HTTP endpoints
# ---------------------------------------------------------------------------
@app.get("/healthz")
def healthz():
    return {"ok": True}


@app.post("/process", response_model=ProcessResponse)
def process(req: ProcessRequest,
            x_kingdom_auth: Optional[str] = Header(default=None)):
    if SHARED_SECRET and x_kingdom_auth != SHARED_SECRET:
        raise HTTPException(status_code=401, detail="bad shared secret")

    page = notion_get_page(req.page_id)
    project = parse_project(page)

    if not project["job_number"]:
        return ProcessResponse(
            ok=True, action="skip", page_id=req.page_id,
            skip_reason="project name does not start with 8-digit job number",
            message="skipped",
        )

    if not project["folder_checked"]:
        folder_name, sp_url = do_create_folder(project)
        log.info("created folder %s for page %s", folder_name, req.page_id)

        # If the project is ALSO in Proposal Requested, render right now so the
        # partner doesn't have to wait for a second webhook fire.
        if project["engineering_status"] == "Proposal Requested":
            # Re-parse with the freshly-set Folder=true
            project["folder_checked"] = True
            project["sharepoint_folder"] = sp_url
            result = do_render_proposal(project)
            return ProcessResponse(
                ok=True, action="create_folder+render", page_id=req.page_id,
                job_number=project["job_number"], folder_name=folder_name,
                contract_path=result.get("contract_path"),
                memo_path=result.get("memo_path"),
                message=f"created folder and rendered proposal ({result['status']})",
            )

        return ProcessResponse(
            ok=True, action="create_folder", page_id=req.page_id,
            job_number=project["job_number"], folder_name=folder_name,
            message="created folder and updated Notion",
        )

    if project["engineering_status"] == "Proposal Requested":
        result = do_render_proposal(project)
        return ProcessResponse(
            ok=True, action="render_proposal", page_id=req.page_id,
            job_number=project["job_number"],
            folder_name=result.get("folder_name"),
            contract_path=result.get("contract_path"),
            memo_path=result.get("memo_path"),
            message=f"proposal {result['status']}",
        )

    return ProcessResponse(
        ok=True, action="skip", page_id=req.page_id,
        job_number=project["job_number"],
        skip_reason=f"folder already synced; status={project['engineering_status']}",
        message="no action needed",
    )


@app.post("/process/raw")
async def process_raw(request: Request,
                      x_kingdom_auth: Optional[str] = Header(default=None)):
    """Alternative endpoint that accepts Notion's native automation webhook
    payload shape directly. Useful if you wire Notion → this service
    without n8n in the middle."""
    body = await request.json()
    page_id = (
        body.get("data", {}).get("id")
        or body.get("id")
        or body.get("page", {}).get("id")
    )
    if not page_id:
        raise HTTPException(status_code=400, detail="no page_id in payload")
    return process(ProcessRequest(page_id=page_id, source="notion_direct"),
                   x_kingdom_auth=x_kingdom_auth)


class RenderContractRequest(BaseModel):
    brief_id: str
    source: Optional[str] = "webhook"


class RenderContractResponse(BaseModel):
    ok: bool
    page_id: str
    brief_id: str
    job_number: Optional[str] = None
    folder_name: Optional[str] = None
    contract_path: Optional[str] = None
    memo_path: Optional[str] = None
    message: str = ""
    skip_reason: Optional[str] = None


@app.post("/render-contract", response_model=RenderContractResponse)
def render_contract(req: RenderContractRequest,
                    x_kingdom_auth: Optional[str] = Header(default=None)):
    """Render contract from a Brief ID. Used by n8n workflow 06.
    Fetches Brief → Project, then calls do_render_proposal with brief overlay."""
    if SHARED_SECRET and x_kingdom_auth != SHARED_SECRET:
        raise HTTPException(status_code=401, detail="bad shared secret")

    brief_page = notion_get_page(req.brief_id)
    brief = parse_brief(brief_page)

    # Find linked project
    project = fetch_project_for_brief(brief_page)
    if not project:
        return RenderContractResponse(
            ok=False, page_id="", brief_id=req.brief_id,
            skip_reason="brief has no linked project (Parent item relation empty)",
            message="cannot render — brief not linked to a project",
        )

    if not project["job_number"]:
        return RenderContractResponse(
            ok=False, page_id=project["page_id"], brief_id=req.brief_id,
            skip_reason="project name does not start with 8-digit job number",
            message="skipped",
        )

    if not project["folder_checked"]:
        return RenderContractResponse(
            ok=False, page_id=project["page_id"], brief_id=req.brief_id,
            job_number=project["job_number"],
            skip_reason="project folder not yet created",
            message="cannot render — folder missing",
        )

    result = do_render_proposal(project, brief=brief)
    return RenderContractResponse(
        ok=True, page_id=project["page_id"], brief_id=req.brief_id,
        job_number=project["job_number"],
        folder_name=result.get("folder_name"),
        contract_path=result.get("contract_path"),
        memo_path=result.get("memo_path"),
        message=f"contract render {result.get('status','ok')}",
    )
