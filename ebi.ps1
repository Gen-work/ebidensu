#Requires -Version 5.1
# ============================================================
#  ebi.ps1 - the ebi-dance command-line entry point.
#
#  P0-07 / P0-08 surface only. P1-07..P1-10 add help / lint / explain /
#  doctor and the real run options (--resume, --only, --operator, see
#  docs/ebi-dance/Plan.md section 9).
#
#    .\ebi.ps1 run    workflows\spike.capture_window.json -WorkDir C:\work\x
#    .\ebi.ps1 dryrun workflows\spike.capture_window.json -WorkDir C:\work\x
#
#  Exit codes: 0 ok, 1 a step failed, 3 the operator quit, 2 usage.
# ============================================================
param(
    [Parameter(Position = 0)] [string]$Command  = 'help',
    [Parameter(Position = 1)] [string]$Workflow = '',
    [string]$WorkDir = '',
    [string]$RunId   = '',
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
# Capture switches BEFORE dot-sourcing anything (CLAUDE.md switch pattern).
$dryRunFlag  = [bool]$DryRun.IsPresent

. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Runner.ps1')

function Show-EbiUsage {
    Write-Host ''
    Write-Host 'ebi - declarative evidence workflows (P0 spike surface)' -ForegroundColor Cyan
    Write-Host '  ebi.ps1 run    <workflow.json> [-WorkDir <dir>] [-RunId <id>]   run setup + teardown'
    Write-Host '  ebi.ps1 dryrun <workflow.json> [-WorkDir <dir>]                 same, every step in DryRun'
    Write-Host '  ebi.ps1 help'
    Write-Host ''
    Write-Host '  WorkDir defaults to the current directory. Outputs land in <WorkDir>\capture, trace in <WorkDir>\run\<runId>\trace.jsonl.'
}

$cmd = $Command.ToLowerInvariant()
if ($cmd -eq 'help' -or $cmd -eq '-h' -or $cmd -eq '--help') { Show-EbiUsage; exit 0 }
if ($cmd -ne 'run' -and $cmd -ne 'dryrun') {
    Write-Host ('[ERROR] unknown command: ' + $Command) -ForegroundColor Red
    Show-EbiUsage
    exit 2
}
if ([string]::IsNullOrWhiteSpace($Workflow)) {
    Write-Host '[ERROR] a workflow file is required' -ForegroundColor Red
    Show-EbiUsage
    exit 2
}
if ($cmd -eq 'dryrun') { $dryRunFlag = $true }

$wd = if ([string]::IsNullOrWhiteSpace($WorkDir)) { (Get-Location).Path } else { $WorkDir }
if (-not (Test-Path -LiteralPath $wd)) { New-Item -ItemType Directory -Path $wd -Force | Out-Null }
$wd = (Resolve-Path -LiteralPath $wd).Path

$wfPath = $Workflow
if (-not [System.IO.Path]::IsPathRooted($wfPath)) {
    $candidate = Join-Path (Get-Location).Path $wfPath
    if (-not (Test-Path -LiteralPath $candidate)) { $candidate = Join-Path $PSScriptRoot $wfPath }
    $wfPath = $candidate
}

$summary = Invoke-EbiWorkflow -Path $wfPath -WorkDir $wd -RunId $RunId -DryRun:$dryRunFlag
if ([string]$summary['failure'] -eq 'operator_quit') { exit 3 }
if ([bool]$summary['ok']) { exit 0 }
exit 1
