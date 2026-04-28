"""
Kingdom Structural automation — centralized constants and templates.

Everything here is Phase-2-n8n-portable: Notion property names, status
values, database IDs, and notification copy all live in this one module
so the port to n8n becomes a find-and-replace exercise, not a code audit.

Created 2026-04-22 per v2.1 Change Order, §7 (n8n migration readiness).
"""
from __future__ import annotations


# ========================================================================
# Notion database / data-source IDs
# ========================================================================
# These are `collection://` IDs (data source IDs). They are stable and safe
# to hardcode. If a DB is ever re-created, update here in one place.

PROJECTS_DS_ID       = "262b73dc-460e-8137-b3bb-000b62103b15"
BRIEF_DS_ID          = "6b64658f-5fb3-4329-b02c-3ac3cb8a0828"
CLIENT_DS_ID         = "2efb73dc-460e-80fe-860c-000bdfd01348"
CONTACT_DS_ID        = "c88189fd-5bdf-46ad-a02e-9f426fd58527"
JOBS_26KS_DS_ID      = "c58ca341-4985-45bf-a17f-19df8553a132"
AUTOMATION_LOG_DS_ID = "04d7aa20-eb8d-4d51-b302-25fb4eaaf26e"


# ========================================================================
# Property names (Projects DB)
# ========================================================================
class ProjectProp:
    NAME                = "Project Name"
    CLIENT              = "Client"
    PROJECT_CONTACT     = "Project Contact"
    ENGINEER            = "ENGINEER"
    ADMIN               = "Admin"
    DRAFTER             = "DRAFTER"
    # Address fields (restored per v2.1 §6 Option A)
    CITY                = "City"
    STATE               = "State"
    PROJECT_STREET      = "Project Street"
    # Relations
    PROPOSAL_BRIEF      = "Proposal Brief"
    # Automation state
    FOLDER              = "Folder"                    # checkbox
    SHAREPOINT_FOLDER   = "SharePoint Folder"         # url
    STATUS              = "Status"                    # select (v2.1)
    ENGINEER_NOTIFIED   = "Engineer Notified"         # checkbox (v2.1)
    INTAKE_COMPLETE     = "Intake Complete"           # checkbox (v2.1.4):
                                                      # admin ticks this after
                                                      # confirming all onboarding
                                                      # fields are populated. When
                                                      # true + Engineering Status
                                                      # == "Proposal Requested",
                                                      # Jobs A0/A/C fire.
    # Legacy status the team already uses — kept for reporting
    ENGINEERING_STATUS  = "Engineering Status"


class ProjectStatus:
    """Values of Projects.Status (v2.1). Drives automation state only;
    Engineering Status is the team-facing human status and is untouched."""
    DRAFT                 = "Draft"
    FOLDER_PENDING        = "Folder Creation Pending"
    BRIEF_PENDING         = "Proposal Brief Pending"
    ADMIN_REVIEW          = "Admin Review"
    CONTRACT_RENDERED     = "Contract Rendered"
    FOLDER_FAILED         = "⚠ Folder Creation Failed"


# ========================================================================
# Property names (Proposal Brief DB)
# ========================================================================
class BriefProp:
    NAME              = "Name"
    PROJECT           = "Project"                     # relation to Projects
    PROJECT_CONTACT   = "Project Contact"
    # Context/address (engineer fills these in on the Brief)
    CITY              = "City"
    STATE             = "State"
    JURISDICTION      = "Jurisdiction"
    PROJECT_ADDRESS   = "Project Address"
    PROJECT_TYPE      = "Project Type"
    ICC_CODE_YEAR     = "ICC Code Year"
    # Render gate + state
    READY_TO_RENDER   = "Ready to Render"             # button (legacy trigger)
    RENDERED_AT       = "Rendered At"                 # date
    STATUS            = "Status"                      # select (v2.1)
    ADMINS_NOTIFIED   = "Admins Notified"             # checkbox (v2.1)


class BriefStatus:
    """Values of Proposal Brief.Status (v2.1)."""
    NOT_STARTED            = "Not Started"
    IN_PROGRESS            = "In Progress"
    READY_FOR_ADMIN_REVIEW = "Ready for Admin Review"
    APPROVED               = "Approved"
    RENDERED               = "Rendered"


# ========================================================================
# Property names (Automation Log DB, v2.1)
# ========================================================================
class LogProp:
    LOG_ENTRY  = "Log Entry"                          # title
    TIMESTAMP  = "Timestamp"
    JOB        = "Job"
    OUTCOME    = "Outcome"
    PROJECT    = "Project"
    BRIEF      = "Brief"
    ERROR      = "Error"
    DETAILS    = "Details"


class LogJob:
    """Job identifiers written to the Automation Log."""
    A0  = "A0"       # Auto-create Proposal Brief
    A02 = "A02"      # Auto-provision child Fee Schedule DB
    A   = "A"        # Folder sync (OneDrive tree)
    B   = "B"        # Contract render
    C   = "C"        # Engineer notification
    D   = "D"        # Admin notification


class LogOutcome:
    SUCCESS = "Success"
    SKIPPED = "Skipped"        # idempotency hit — nothing to do
    FAILED  = "Failed"


# ========================================================================
# Notification templates (v2.1 §5)
# ========================================================================
# Jinja-style {placeholders} so n8n Set-node templates are drop-in copies.
# Fill with str.format(**context) or equivalent in Python; n8n will use
# expression syntax {{ $json.x }} — the text itself is portable.

ENGINEER_EMAIL_SUBJECT = (
    "[{project_number}] — Folders ready for your Proposal Brief"
)

ENGINEER_EMAIL_BODY = """Hey {engineer_first_name} — the OneDrive folder for {project_number} {project_name} has been set up and is ready to go.

OneDrive: {folder_url}

Proposal Brief: {brief_url}

When you have a minute, please complete the Proposal Brief (City, State, ICC Code Year, Jurisdiction, scope checkboxes, fee schedule). Once you're done, flip the Brief Status to "Ready for Admin Review" — that'll notify the partners to look it over.

Thanks!
"""

ENGINEER_NOTION_COMMENT = """@{engineer_mention} — folders are ready for {project_number} {project_name}.

- OneDrive: {folder_url}
- Proposal Brief: {brief_url}

When done, flip the Brief Status to "Ready for Admin Review" to loop in the partners.
"""


ADMIN_EMAIL_SUBJECT = (
    "[{project_number}] — Proposal Brief ready for your review"
)

ADMIN_EMAIL_BODY = """The Proposal Brief for {project_number} {project_name} has been completed by {engineer_name} and is ready for partner review.

Brief: {brief_url}

OneDrive: {folder_url}

Once you've reviewed and made any adjustments, check "Ready to Render" on the Brief to kick off the contract build.
"""

ADMIN_NOTION_COMMENT = """{admin_mentions} — Proposal Brief for {project_number} {project_name} is ready for your review.

Completed by {engineer_name}.

- Brief: {brief_url}
- OneDrive: {folder_url}

Tick "Ready to Render" on the Brief to kick off the contract build.
"""


# ========================================================================
# Error-path notification (v2.1 §5, Job A failure)
# ========================================================================
ADMIN_ERROR_COMMENT = """{admin_mentions} — ⚠ Folder Creation Failed for {project_number} {project_name}.

Error: {error_message}

The Project.Status has been set to "⚠ Folder Creation Failed" and the error is recorded in Project.Last Automation Error. No folder was created and the Engineer has NOT been notified. Please investigate and re-run once resolved.
"""


# ========================================================================
# Helpers
# ========================================================================
def format_engineer_email(context: dict) -> tuple[str, str]:
    """Returns (subject, body)."""
    return (
        ENGINEER_EMAIL_SUBJECT.format(**context),
        ENGINEER_EMAIL_BODY.format(**context),
    )


def format_admin_email(context: dict) -> tuple[str, str]:
    """Returns (subject, body)."""
    return (
        ADMIN_EMAIL_SUBJECT.format(**context),
        ADMIN_EMAIL_BODY.format(**context),
    )


# ========================================================================
# Engineer & Admin roster (v2.1.1 — cached, no runtime Notion user lookup)
# ========================================================================
# Keyed by Notion user UUID. Values: (display_name, email). Used by Job C
# to avoid an MCP round-trip per engineer per sweep.
#
# Sync process: when a new engineer is onboarded, add their Notion user ID
# here manually. The sweep script does NOT mutate this dict.

ENGINEER_ROSTER: dict[str, tuple[str, str]] = {
    "249d872b-594c-8166-8a39-0002e2633684": ("Jonathan Brusnahan", "jbrusnahan@kingdomstructural.com"),
    "314d872b-594c-811b-8018-000251c1efe5": ("Gabriel O'reilly",   "goreilly@kingdomstructural.com"),
    "249d872b-594c-817d-b942-0002cf0617f8": ("Chris Valdez",       "cvaldez@kingdomstructural.com"),
    "2bad872b-594c-817d-8702-000294b684d0": ("Michael Brusnahan",  "mbrusnahan@kingdomstructural.com"),
    "261d872b-594c-8114-9aea-0002b83ff5ab": ("Partner (Admin)",    "admin@kingdomstructural.com"),  # TODO confirm
}


def engineer_first_name(user_id: str) -> str:
    """Return 'Jonathan' from 'Jonathan Brusnahan'. Falls back to 'there'."""
    full = ENGINEER_ROSTER.get(user_id, (None, None))[0]
    return full.split(" ", 1)[0] if full else "there"


def engineer_email(user_id: str) -> str | None:
    return ENGINEER_ROSTER.get(user_id, (None, None))[1]


# ========================================================================
# Job A folder-naming helper (v2.1.2)
# ========================================================================
import re as _re


class InvalidProjectNumberError(ValueError):
    """Raised when Project Name doesn't start with an 8-digit KS job number.

    Carries the raw name plus a user-facing message the automation can drop
    straight into the Automation Log so partners see a clean rename prompt
    instead of a Python traceback.
    """

    def __init__(self, project_name: str):
        self.project_name = project_name
        super().__init__(
            f'Project Name must start with an 8-digit KS job number '
            f'(e.g. "26107020 Project Name"), got {project_name!r}. '
            f'Rename the Project in Notion to include a full 8-digit prefix, '
            f'then re-run the sweep.'
        )


def parse_project_name(project_name: str) -> tuple[str, str]:
    """Split '26107020 Jacksonville First Assembly' into ('26107020', 'Jacksonville First Assembly').

    Also handles '26107020 - Jacksonville...' and '26107020-Jacksonville...'
    The number prefix MUST be exactly 8 digits — Job A refuses to build a
    folder for anything shorter because (a) year parsing depends on the full
    prefix and (b) KS's internal job-number convention is 8 digits.

    Returns (number, short_name). Raises InvalidProjectNumberError on mismatch.
    """
    # Anchor on word boundary after digits so '2999887' doesn't match as 7-of-8.
    m = _re.match(r"^(\d{8})(?=\D|$)\s*[-\u2014]?\s*(.*)$", project_name.strip())
    if not m:
        raise InvalidProjectNumberError(project_name)
    number, rest = m.group(1), m.group(2).strip()
    rest = _re.sub(r"<br\s*/?>", "", rest, flags=_re.IGNORECASE).strip()
    rest = rest.rstrip(" -\u2014")
    return number, rest


def build_folder_name(number: str, short_name: str, city: str, state: str) -> str:
    """v2.1.7 naming: '26107020 Jacksonville First Assembly - Jacksonville, AR'.

    No dash between job number and project name (the dash was added in
    v2.1.1, removed 2026-04-27 to match the team's pre-v2.1.1 convention
    and keep all OneDrive folders visually consistent).

    The location suffix still uses ' - ' as the separator before
    '{City}, {State}', because that's a logical boundary between
    project name and project location.
    """
    return f"{number} {short_name} - {city}, {state}"


def build_year_from_number(number: str) -> str:
    """'26107020' → '2026'."""
    return "20" + number[:2]


# ========================================================================
# Retry policy for transient network errors (v2.1.5)
# ========================================================================
# Notion API + Microsoft Graph reads/writes occasionally die on a Windows
# DNS hiccup (`getaddrinfo failed [Errno 11001]`), a brief connection
# reset, or a 5xx from the upstream service. The sweep should retry these
# transparently a few times before giving up and writing a Failed log row.
#
# When a sweep job hits a retryable error pattern, wait BACKOFF_SECONDS[i]
# before attempt i+1 (zero-indexed). With MAX_ATTEMPTS = 3 and the default
# tuple, the timeline is:
#   attempt 1 → fail → sleep 0.5s → attempt 2 → fail → sleep 1.5s → attempt 3
# Total wall time on a fully-failed call: ~2 seconds.

class RetryPolicy:
    MAX_ATTEMPTS    = 3
    BACKOFF_SECONDS = (0.5, 1.5, 4.0)   # sleep BEFORE attempt 2, 3, 4

    # Substring match against the error string. Order doesn't matter; we
    # just check `any(p in err_str for p in RETRYABLE_PATTERNS)`.
    RETRYABLE_PATTERNS = (
        "getaddrinfo failed",     # Windows DNS lookup failed
        "Errno 11001",            # same, numeric code
        "Connection reset",
        "Connection aborted",
        "ConnectionResetError",
        "RemoteDisconnected",
        "timed out",
        "ETIMEDOUT",
        "ECONNRESET",
        "TemporaryFailure",
        "ServiceUnavailable",
        " 502 ",
        " 503 ",
        " 504 ",
        "Too Many Requests",      # Notion 429 — retry with backoff
        " 429 ",
    )


def is_retryable_error(err: BaseException | str) -> bool:
    """Pattern-match the error string against RetryPolicy.RETRYABLE_PATTERNS.

    Accepts either an exception or a pre-formatted string. Used by the
    scheduled-task prompt's retry wrapper so policy lives in one place.
    """
    s = str(err) if isinstance(err, BaseException) else err
    return any(p in s for p in RetryPolicy.RETRYABLE_PATTERNS)


# ========================================================================
# A02 short-circuit list (v2.1.5)
# ========================================================================
# A02 (Auto-provision child Fee Schedule DB) used to fetch every Brief's
# child blocks every sweep to check the gate, even for Briefs whose
# workflow had already advanced past needing one. That created ~120
# unnecessary network calls per sweep and a corresponding failure surface.
#
# Briefs in any of these statuses are guaranteed to either already have a
# Fee Schedule (because the engineer has filled it in / admin has approved
# / job has rendered) or to be past the point where one would be useful.
# A02 SHOULD filter them out at the initial Brief query — never fetch
# their children.

A02_SKIP_STATUSES: tuple[str, ...] = (
    BriefStatus.READY_FOR_ADMIN_REVIEW,
    BriefStatus.APPROVED,
    BriefStatus.RENDERED,
)
