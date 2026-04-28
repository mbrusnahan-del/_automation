# Starter prompts for Claude Code

Copy/paste one of these after you `cd` into `_automation_n8n/` and run `claude`.
CLAUDE.md gives Claude the project context automatically — these prompts just
point at what you want to do first.

---

## A. Get oriented (safe first prompt)

```
Read CLAUDE.md and README.md, then give me a one-screen summary of:
1. what's deployable today vs. what's still stubbed,
2. any TODOs or REPLACE-ME markers in the code or n8n JSON,
3. the shortest path from a cold machine to a working end-to-end test
   (Cloud Run deploy + n8n import + one Notion automation wired up).

Don't change any files yet.
```

## B. Actually deploy it

```
I want to ship the render service to Cloud Run. My GCP project is
<PROJECT_ID>, region us-central1. Secrets live in Secret Manager under
notion-token, graph-client-secret, render-shared-secret (create them if
missing; I'll paste values into the prompts). GRAPH_TENANT_ID is <tenant>,
GRAPH_CLIENT_ID is <client>, GRAPH_USER_UPN is
michaelbrusnahan@kingdomstructural.com.

Walk me through the deploy, running each gcloud command one at a time so I
can approve. After the service is live, hit /healthz to confirm, then do a
smoke test against a known Notion page id I'll provide.
```

## C. Add tests before you touch anything

```
Add unit tests for the pure helpers in render_service/main.py:
parse_project, parse_client, parse_contact, and _idempotency_skip.
Use pytest + fixture JSON captured from real Notion responses (I'll paste
a sample when you ask). Don't mock Notion HTTP; just test the parsers.
Target: green `pytest` run. Don't refactor production code unless a test
reveals a real bug.
```

## D. Put this under git

```
Initialize this folder as a git repo. .gitignore should exclude
__pycache__, .env, *.pyc, and any local secret files. Stage everything
else, make one initial commit ("initial import: notion→n8n→render service"),
then ask me before pushing — I'll give you the remote URL.
```

## E. Add the third Notion trigger (Brief Ready to Render)

```
README section 5 mentions an optional third automation: trigger when a
Proposal Brief's "Ready to Render" checkbox flips true, resolve the linked
Project, and call /process with that Project's page_id. Extend main.py
with a /process/brief endpoint that accepts a Brief page_id, fetches the
Brief, reads its Project relation, and calls the existing do_*() helpers
on the Project. Update the n8n JSON with a second webhook path
kingdom-notion-brief that targets the new endpoint. Don't remove the
existing /process path.
```

---

## How to start Claude Code here

**First time on this machine** — run the bootstrap script (installs Claude
Code if missing, adds Notion/GitHub/Sentry/Linear MCPs at user scope):

```
cd "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_automation_n8n"
./setup-claude-code.ps1
```

**Every time after that**, just:

```
cd "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_automation_n8n"
claude
```

On your very first `claude` run in this folder you'll see an approval prompt
for the Notion MCP listed in `.mcp.json` — approve it, then run `/mcp` inside
the session to kick off the Notion OAuth flow in your browser. After that
it's cached.

(Requires Node 18+. First time you launch `claude` itself, it asks you to
log in with your Anthropic account.)
