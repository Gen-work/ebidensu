# Export-DailyPatch.ps1
# ---------------------------------------------------------
# Extracts today's git diff (patch) and copies it to clipboard.
# Lightweight sync tool to avoid feeding the full codebase to an LLM.
# ---------------------------------------------------------

Write-Host "Computing latest commit and current changes..." -ForegroundColor Cyan

# 1. Check for uncommitted changes; if any, stage a temporary commit.
$status = git status --porcelain
if ($status) {
    git add .
    git commit -m "chore: temp WIP commit for home sync" | Out-Null
    Write-Host "[Info] Uncommitted changes saved as a WIP commit." -ForegroundColor Yellow
}

# 2. Get the full diff of the most recent commit.
# (To include multiple commits, change HEAD~1 to HEAD~3 etc.)
$patchData = git diff HEAD~1 HEAD

if (-not $patchData) {
    Write-Warning "No diff found. No code changes detected."
    exit
}

# 3. Build LLM prompt
$output = "The following is the git diff (patch) for today's work.`n"
$output += "Learn from this diff to understand the current implementation. You do not need to output the full code.`n`n"
$output += "``````diff`n"
$output += $patchData
$output += "`n`````````n"

# Copy to clipboard
Set-Clipboard -Value $output
Write-Host "[OK] Today's diff (patch) copied to clipboard!" -ForegroundColor Green
Write-Host ("Character count: approx {0}. Paste into your AI Web UI." -f $output.Length) -ForegroundColor Gray
