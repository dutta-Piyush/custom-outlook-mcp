# Outlook MCP Setup
# This script configures the Outlook MCP server for use with VS Code Copilot.

$ErrorActionPreference = "Stop"
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path

Write-Host "Outlook MCP Setup" -ForegroundColor Cyan
Write-Host "================================`n"

# ── Collect OUTLOOK_TOKEN ────────────────────────────────────────────────────
Write-Host "You need a Microsoft Graph API bearer token (OUTLOOK_TOKEN)."
Write-Host "To obtain one:"
Write-Host "  1. Open https://developer.microsoft.com/en-us/graph/graph-explorer"
Write-Host "  2. Sign in and click 'Access token' to copy the token."
Write-Host "  OR use your organisation's app registration / delegated OAuth flow."
Write-Host ""
$token = Read-Host "Paste your OUTLOOK_TOKEN"
if (-not $token) {
    Write-Host "ERROR: OUTLOOK_TOKEN is required." -ForegroundColor Red
    exit 1
}

# ── Optional proxy ───────────────────────────────────────────────────────────
$proxy = Read-Host "Optional: enter corporate proxy URL (leave blank to skip)"

# ── Write .vscode/mcp.json ───────────────────────────────────────────────────
$vscodedir = Join-Path $ScriptDir ".vscode"
if (-not (Test-Path $vscodedir)) {
    New-Item -ItemType Directory -Path $vscodedir | Out-Null
}

$env_block = [ordered]@{
    OUTLOOK_TOKEN = $token
}
if ($proxy) {
    $env_block["OUTLOOK_PROXY"] = $proxy
}

$config = [ordered]@{
    servers = [ordered]@{
        "outlook-mcp" = [ordered]@{
            type    = "stdio"
            command = "python"
            args    = @("$ScriptDir\server.py")
            env     = $env_block
        }
    }
}

$json = $config | ConvertTo-Json -Depth 10
$outPath = Join-Path $vscodedir "mcp.json"
Set-Content -Path $outPath -Value $json -Encoding UTF8
Write-Host "`nConfiguration written to: $outPath" -ForegroundColor Green
Write-Host "Restart VS Code and enable the 'outlook-mcp' server in Copilot settings." -ForegroundColor Cyan
