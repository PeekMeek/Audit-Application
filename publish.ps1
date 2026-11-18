<#
  Publishes the Additive Applications app to GitHub Pages.

  Usage (from this folder in PowerShell):
      .\publish.ps1 "What you changed"

  It bumps the offline-cache version in sw.js (so phones pick up the new files),
  then commits everything and pushes to GitHub. Pages updates about a minute later.
#>
param(
  [Parameter(Position = 0)]
  [string]$Message = "Update app"
)

$ErrorActionPreference = 'Stop'
Set-Location -Path $PSScriptRoot

if (-not (Get-Command git -ErrorAction SilentlyContinue)) {
  Write-Error "Git isn't installed. Install it with:  winget install --id Git.Git -e   then open a new PowerShell window."
}

# Stop early if nothing changed
$changes = git status --porcelain
if ($LASTEXITCODE -ne 0) { Write-Error "This folder isn't a git repository yet. See 'Pushing updates from PowerShell' in README.md." }
if (-not $changes) { Write-Host "Nothing has changed since the last publish." -ForegroundColor Yellow; exit 0 }

# Bump the cache version in sw.js: aa-v2 -> aa-v3
$sw = Join-Path $PSScriptRoot 'sw.js'
$text = [System.IO.File]::ReadAllText($sw)
$match = [regex]::Match($text, "const VERSION = 'aa-v(\d+)';")
if ($match.Success) {
  $old = [int]$match.Groups[1].Value
  $new = $old + 1
  $text = $text.Replace($match.Value, "const VERSION = 'aa-v$new';")
  [System.IO.File]::WriteAllText($sw, $text)   # UTF-8, no BOM
  Write-Host "Cache version: aa-v$old -> aa-v$new"
} else {
  Write-Warning "Couldn't find VERSION in sw.js, so phones may keep showing the old version."
}

git add -A
if ($LASTEXITCODE -ne 0) { Write-Error "git add failed." }

git commit -m $Message
if ($LASTEXITCODE -ne 0) { Write-Error "git commit failed." }

git push
if ($LASTEXITCODE -ne 0) { Write-Error "git push failed. Check your internet connection and GitHub sign-in, then run the script again." }

$remote = git remote get-url origin
if ($remote -match 'github\.com[:/]([^/]+)/([^/.]+)') {
  Write-Host ""
  Write-Host "Published. Live in about a minute at: https://$($Matches[1].ToLower()).github.io/$($Matches[2])/" -ForegroundColor Green
} else {
  Write-Host "Published." -ForegroundColor Green
}
