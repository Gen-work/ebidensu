#Requires -Version 5.1
<#
.SYNOPSIS
    VerifyTool — main entry for the GIFT→GFIX migration evidence workflow.
.DESCRIPTION
    Routes to phase-specific scripts based on -Phase.
    Remembers WorkDir, Owner, window geometry, CursorCell, CloneSourceDir
    between runs via verify_session.json.
.EXAMPLE
    .\VerifyTool.ps1
    .\VerifyTool.ps1 -Phase Status
    .\VerifyTool.ps1 -Phase Validate
    .\VerifyTool.ps1 -Phase GiftHmSnap -TargetIds JIGPL48S
#>
param(
    [string]  $Phase          = 'Menu',
    [string]  $WorkDir        = '',
    [string]  $Owner          = '',
    [string[]]$TargetIds      = @(),
    [string]  $CloneSourceDir = '',
    [string[]]$BizCodes       = @(),
    [switch]  $Force,
    [switch]  $Interactive,
    [switch]  $RefreshUrls,
    [switch]  $DryRun,
    [switch]  $NoResize,
    [int]     $WindowWidth    = 0,
    [int]     $WindowHeight   = 0,
    [int]     $CropPx         = -1,
    [string]  $CursorCell     = '',
    [switch]  $Help
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir   = Split-Path $MyInvocation.MyCommand.Path
$sessionFile = Join-Path $scriptDir 'verify_session.json'
$cfgFile     = Join-Path $scriptDir 'VerifyConfig.psd1'

# ── load config ───────────────────────────────────────────────────────────────
function Load-VerifyConfig {
    $raw = Import-PowerShellDataFile $cfgFile

    # resolve script paths relative to scriptDir
    $scripts = @{}
    foreach ($k in $raw.Scripts.Keys) {
        $scripts[$k] = Join-Path $scriptDir $raw.Scripts[$k]
    }
    $raw.Scripts = $scripts
    return $raw
}

$cfg = Load-VerifyConfig

# ── session persistence ───────────────────────────────────────────────────────
function Load-Session {
    if (Test-Path $sessionFile) {
        try { return Get-Content $sessionFile -Raw | ConvertFrom-Json }
        catch { }
    }
    return [pscustomobject]@{
        WorkDir       = $cfg.DefaultWorkDir
        Owner         = $cfg.DefaultOwner
        WindowWidth   = $cfg.Window.Width
        WindowHeight  = $cfg.Window.Height
        CropPx        = $cfg.Window.CropPx
        CursorCell    = $cfg.Review.CursorCell
        CloneSourceDir= ''
        EvidenceDir   = ''
    }
}

function Save-Session([pscustomobject]$s) {
    $s | ConvertTo-Json | Set-Content $sessionFile -Encoding UTF8
}

$session = Load-Session

# ── merge CLI → session ───────────────────────────────────────────────────────
if ($WorkDir)        { $session.WorkDir        = $WorkDir }
if ($Owner)          { $session.Owner          = $Owner }
if ($CloneSourceDir) { $session.CloneSourceDir = $CloneSourceDir }
if ($CursorCell)     { $session.CursorCell     = $CursorCell }
if ($WindowWidth  -gt 0) { $session.WindowWidth  = $WindowWidth }
if ($WindowHeight -gt 0) { $session.WindowHeight = $WindowHeight }
if ($CropPx -ge 0)       { $session.CropPx       = $CropPx }

$WorkDir  = $session.WorkDir
$Owner    = $session.Owner

# ── help ──────────────────────────────────────────────────────────────────────
if ($Help -or $Phase -eq 'Help') {
    Get-Help $MyInvocation.MyCommand.Path -Detailed
    exit 0
}

# ── resolve phase alias ───────────────────────────────────────────────────────
function Resolve-Phase([string]$raw) {
    $key = $raw.Trim()
    if ($cfg.Aliases.ContainsKey($key)) { return $cfg.Aliases[$key] }
    # case-insensitive fallback
    foreach ($a in $cfg.Aliases.Keys) {
        if ($a -ieq $key) { return $cfg.Aliases[$a] }
    }
    return $key
}

# ── mapping helpers ───────────────────────────────────────────────────────────
function Get-MappingPath {
    return Join-Path $WorkDir ($cfg.Paths.MappingPattern -f $Owner)
}

function Read-Mapping {
    $mp = Get-MappingPath
    if (-not (Test-Path $mp)) { return @() }
    return @(Import-Csv $mp -Encoding UTF8BOM)
}

function Write-Mapping([object[]]$rows) {
    $mp = Get-MappingPath
    $rows | Export-Csv $mp -NoTypeInformation -Encoding UTF8BOM
}

# ── ensure phase columns exist (auto-repair on startup) ───────────────────────
function Ensure-PhaseColumns {
    $mp = Get-MappingPath
    if (-not (Test-Path $mp)) { return }
    $rows = @(Import-Csv $mp -Encoding UTF8BOM)
    if ($rows.Count -eq 0) { return }

    $fields = $cfg.PhaseOrder | Where-Object { $_.Field -ne '' } | ForEach-Object { $_.Field } | Select-Object -Unique
    $sample = $rows[0]
    $missing = $fields | Where-Object { -not ($sample.PSObject.Properties.Name -contains $_) }

    if ($missing.Count -eq 0) { return }

    Write-Host "[RepairMapping] Adding missing columns: $($missing -join ', ')" -ForegroundColor Cyan
    foreach ($row in $rows) {
        foreach ($col in $missing) {
            $row | Add-Member -NotePropertyName $col -NotePropertyValue '' -Force
        }
    }
    $rows | Export-Csv $mp -NoTypeInformation -Encoding UTF8BOM
}

# ── status display ────────────────────────────────────────────────────────────
function Show-Status {
    $rows = Read-Mapping
    if ($rows.Count -eq 0) {
        Write-Host "(no mapping)" -ForegroundColor Gray
        return
    }

    $total = $rows.Count
    Write-Host "`n=== Status: $WorkDir ===" -ForegroundColor Cyan
    Write-Host "  Mapping: $(Get-MappingPath)  ($total rows)"

    foreach ($entry in $cfg.PhaseOrder) {
        $field = $entry.Field
        if (-not $field) { continue }
        if ($entry.Status -eq 'planned') { continue }

        $label    = $entry.Label
        $bitValue = $entry.BitValue

        if ($bitValue) {
            $done = ($rows | Where-Object { ($null -ne $_.$field) -and (([int]$_.$field -band $bitValue) -eq $bitValue) }).Count
        } else {
            $done = ($rows | Where-Object { $_.$field -eq '1' }).Count
        }

        $pct   = if ($total -gt 0) { [int](100 * $done / $total) } else { 0 }
        $color = if ($done -eq $total) { 'Green' } elseif ($done -gt 0) { 'Yellow' } else { 'Gray' }
        Write-Host ("  {0,-36} {1,4}/{2}  ({3}%)" -f $label, $done, $total, $pct) -ForegroundColor $color
    }
    Write-Host ''
}

# ── recommend next phase ───────────────────────────────────────────────────────
function Get-RecommendPhase {
    $rows = Read-Mapping
    if ($rows.Count -eq 0) { return 'Mapping' }

    foreach ($entry in $cfg.PhaseOrder) {
        $field  = $entry.Field
        $status = $entry.Status
        if (-not $field -or $status -eq 'planned' -or $status -eq 'legacy') { continue }

        $bitValue = $entry.BitValue
        if ($bitValue) {
            $pending = $rows | Where-Object { ($null -eq $_.$field) -or (([int]$_.$field -band $bitValue) -ne $bitValue) }
        } else {
            $pending = $rows | Where-Object { $_.$field -ne '1' }
        }
        if ($pending.Count -gt 0) { return $entry.Key }
    }
    return 'ReviewEvidence'
}

# ── interactive menu ──────────────────────────────────────────────────────────
function Show-Menu {
    Show-Status

    $recommend = Get-RecommendPhase
    Write-Host "  Recommended: $recommend" -ForegroundColor Magenta
    Write-Host ''
    Write-Host "  Type a phase name (or Enter for '$recommend', Q to quit):" -ForegroundColor White
    $input = Read-Host '  Phase'
    $input = $input.Trim()
    if ($input -eq '' )                  { return $recommend }
    if ($input -ieq 'q' -or $input -ieq 'quit') { exit 0 }
    return $input
}

# ── build common child args ───────────────────────────────────────────────────
function Get-CommonArgs {
    $a = @{
        WorkDir      = $WorkDir
        Owner        = $Owner
        WindowWidth  = $session.WindowWidth
        WindowHeight = $session.WindowHeight
        CropPx       = $session.CropPx
    }
    if ($NoResize)              { $a['NoResize']   = $true }
    if ($DryRun)                { $a['DryRun']     = $true }
    if ($Force)                 { $a['Force']      = $true }
    if ($TargetIds.Count -gt 0) { $a['TargetIds']  = $TargetIds }
    return $a
}

# ── phase dispatcher ──────────────────────────────────────────────────────────
function Invoke-ToolPhase([string]$resolvedPhase) {
    $common = Get-CommonArgs

    switch ($resolvedPhase) {

        'Mapping' {
            $args2 = @{ WorkDir = $WorkDir; Owner = $Owner }
            if ($Force) { $args2['Force'] = $true }
            & $cfg.Scripts.GenerateMapping @args2
        }

        'RepairMapping' {
            Ensure-PhaseColumns
        }

        'ExcelSnap' {
            & $cfg.Scripts.Excel @common
        }

        'GiftHmSnap' {
            & $cfg.Scripts.Hm @common -Mode GiftRecv
        }

        'GfixHmSnap' {
            & $cfg.Scripts.Hm @common -Mode GfixRecv
        }

        'GiftMqSnap' {
            & $cfg.Scripts.Mq @common -Mode GiftRecv
        }

        'GiftJenkins' {
            $a = $common.Clone()
            if ($RefreshUrls) { $a['RefreshUrls'] = $true }
            & $cfg.Scripts.Jenkins @a -Mode GiftRecv
        }

        'GfixJenkins' {
            $a = $common.Clone()
            if ($RefreshUrls) { $a['RefreshUrls'] = $true }
            & $cfg.Scripts.Jenkins @a -Mode GfixRecv
        }

        'GiftJenkinsNoFile' {
            & $cfg.Scripts.Jenkins @common -Mode NoGfix
        }

        'Clone' {
            $a = @{ WorkDir = $WorkDir; Owner = $Owner }
            if ($CloneSourceDir -or $session.CloneSourceDir) {
                $a['CloneSourceDir'] = if ($CloneSourceDir) { $CloneSourceDir } else { $session.CloneSourceDir }
            }
            if ($BizCodes.Count -gt 0) { $a['BizCodes'] = $BizCodes }
            if ($Force)                { $a['Force']     = $true }
            if ($TargetIds.Count -gt 0){ $a['TargetIds'] = $TargetIds }
            if ($DryRun)               { $a['DryRun']    = $true }
            & $cfg.Scripts.Clone @a
        }

        'ReplaceGift' {
            & $cfg.Scripts.Replace @common -Mode Gift
        }

        'ReplaceGfix' {
            & $cfg.Scripts.Replace @common -Mode Gfix
        }

        'ReplaceDf' {
            & $cfg.Scripts.Replace @common -Mode Df
        }

        'MarkGift' {
            & $cfg.Scripts.Mark @common -Mode Gift
        }

        'MarkGfix' {
            & $cfg.Scripts.Mark @common -Mode Gfix
        }

        'MarkDf' {
            & $cfg.Scripts.Mark @common -Mode Df
        }

        'ReviewGift' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner; CursorCell=$session.CursorCell; Mode='Gift' }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            if ($Force)                 { $a['Force']     = $true }
            & $cfg.Scripts.Review @a
        }

        'ReviewGfix' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner; CursorCell=$session.CursorCell; Mode='Gfix' }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            if ($Force)                 { $a['Force']     = $true }
            & $cfg.Scripts.Review @a
        }

        'ReviewDf' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner; CursorCell=$session.CursorCell; Mode='Df' }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            if ($Force)                 { $a['Force']     = $true }
            & $cfg.Scripts.Review @a
        }

        'ReviewEvidence' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner; CursorCell=$session.CursorCell; Mode='All' }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            if ($Force)                 { $a['Force']     = $true }
            & $cfg.Scripts.Review @a
        }

        'GfixLodDownload' {
            Write-Host '[GfixLodDownload] Not implemented yet.' -ForegroundColor Yellow
        }

        'DfSnap' {
            Write-Host '[DfSnap] Not implemented yet.' -ForegroundColor Yellow
        }

        'Validate' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            & $cfg.Scripts.Validate @a
        }

        'ProbeShapes' {
            $a = @{ WorkDir=$WorkDir; Owner=$Owner }
            if ($TargetIds.Count -gt 0) { $a['TargetIds'] = $TargetIds }
            & $cfg.Scripts.Probe @a
        }

        'Crop' {
            & $cfg.Scripts.Crop @common
        }

        'Status' {
            Show-Status
        }

        'Menu' {
            # should not reach here — handled in main loop
        }

        default {
            Write-Warning "Unknown phase: $resolvedPhase"
        }
    }
}

# ── guard: WorkDir required for most phases ───────────────────────────────────
function Assert-WorkDir([string]$phase) {
    $noWorkDirNeeded = @('Menu','Status','Help')
    if ($phase -in $noWorkDirNeeded) { return }
    if (-not $WorkDir) {
        throw "WorkDir is not set. Run with -WorkDir <path> first."
    }
    if (-not (Test-Path $WorkDir)) {
        throw "WorkDir does not exist: $WorkDir"
    }
}

# ── main ──────────────────────────────────────────────────────────────────────

# auto-repair mapping columns on every startup (silent if nothing to add)
if ($WorkDir -and (Test-Path $WorkDir)) {
    try { Ensure-PhaseColumns } catch { }
}

if ($Phase -eq 'Menu') {
    # interactive mode
    while ($true) {
        $rawPhase = Show-Menu
        $resolved = Resolve-Phase $rawPhase
        Write-Host "`n→ $resolved" -ForegroundColor Cyan

        try {
            Assert-WorkDir $resolved
            Invoke-ToolPhase $resolved
        } catch {
            Write-Host "ERROR: $_" -ForegroundColor Red
        }

        Save-Session $session
        Write-Host ''
    }
} else {
    $resolved = Resolve-Phase $Phase
    Assert-WorkDir $resolved
    Invoke-ToolPhase $resolved
    Save-Session $session
}
