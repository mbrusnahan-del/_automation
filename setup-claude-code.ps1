# setup-claude-code.ps1
# One-time setup on this Windows machine: installs Claude Code if missing,
# registers useful MCP connectors at user scope (so they follow you into
# every project, not just this one).
#
# Run from PowerShell:
#   cd "C:\Users\MichaelBrusnahan\OneDrive - Kingdom Structural LLC\_automation_n8n"
#   ./setup-claude-code.ps1
#
# Idempotent: safe to run multiple times. `claude mcp add` will no-op or
# tell you the server already exists.

$ErrorActionPreference = 'Stop'

function Test-CommandExists($name) {
    $null -ne (Get-Command $name -ErrorAction SilentlyContinue)
}

Write-Host ""
Write-Host "=== Kingdom Structural — Claude Code bootstrap ===" -ForegroundColor Cyan
Write-Host ""

# --- 1. Claude Code CLI ---
if (-not (Test-CommandExists 'claude')) {
    Write-Host "Claude Code CLI not found. Installing via npm..." -ForegroundColor Yellow
    if (-not (Test-CommandExists 'npm')) {
        Write-Host "ERROR: npm not found. Install Node.js 18+ from https://nodejs.org first, then re-run." -ForegroundColor Red
        exit 1
    }
    npm install -g '@anthropic-ai/claude-code'
    Write-Host "Installed Claude Code." -ForegroundColor Green
} else {
    $ver = (claude --version) 2>$null
    Write-Host "Claude Code already installed ($ver)." -ForegroundColor Green
}

Write-Host ""

# --- 2. MCP connectors at user scope ---
# Format: name, transport, url, why
$connectors = @(
    @{ name = 'notion'; transport = 'http'; url = 'https://mcp.notion.com/mcp';         why = 'Notion databases — projects, clients, contacts' }
    @{ name = 'github'; transport = 'http'; url = 'https://api.githubcopilot.com/mcp/'; why = 'Repos, PRs, issues' }
    @{ name = 'sentry'; transport = 'http'; url = 'https://mcp.sentry.dev/mcp';         why = 'Error monitoring (optional)' }
    @{ name = 'linear'; transport = 'http'; url = 'https://mcp.linear.com/mcp';         why = 'Issue tracking (optional)' }
)

Write-Host "Registering MCP connectors at user scope..." -ForegroundColor Cyan
foreach ($c in $connectors) {
    Write-Host "  - $($c.name)  [$($c.why)]"
    try {
        claude mcp add --transport $c.transport --scope user $c.name $c.url 2>&1 | Out-Null
        Write-Host "    added" -ForegroundColor Green
    } catch {
        Write-Host "    already present or add failed — continuing" -ForegroundColor DarkYellow
    }
}

Write-Host ""
Write-Host "Current MCP servers:" -ForegroundColor Cyan
claude mcp list

Write-Host ""
Write-Host "=== Done ===" -ForegroundColor Cyan
Write-Host ""
Write-Host "Next steps:" -ForegroundColor White
Write-Host "  1. Stay in this folder and run:   claude"
Write-Host "  2. Inside the session, run:       /mcp"
Write-Host "  3. Approve the OAuth prompts for Notion (and GitHub if you use it)."
Write-Host "  4. Open FIRST_PROMPT.md and paste Prompt A to get oriented."
Write-Host ""
