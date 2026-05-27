# ============================================================
#  Run-Snap.ps1 - one human-friendly entry point for snap phases
#
#  Daily usage:
#    .\Run-Snap.ps1
#
#  Direct examples:
#    .\Run-Snap.ps1 -Phase HmGift -Force
#    .\Run-Snap.ps1 -Phase JenkinsGift -TargetIds JIGPL48S
#    .\Run-Snap.ps1 -Phase MqGift -WindowWidth 1050 -WindowHeight 761 -CropPx 6
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner   = '',
    [ValidateSet('Menu','Mapping','Excel','HmGift','MqGift','JenkinsGift','NoGfix','HmGfix','JenkinsGfix','Crop','Status','Df')]
    [string]$Phase = 'Menu',

    [int]$WindowWidth  = 0,
    [int]$WindowHeight = 0,
    [int]$CropPx       = -1,

    [string[]]$TargetIds = @(),
    [switch]$Force,
    [switch]$Interactive,
    [switch]$NoResize,
    [switch]$RefreshUrls,
    [switch]$DryRun,

    [string]$ConfigPath = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

function Read-Choice([string]$Prompt, [string]$Default = '') {
    if ([string]::IsNullOrWhiteSpace($Default)) {
        return (Read-Host $Prompt)
    }
    $v = Read-Host ("{0} [{1}]" -f $Prompt, $Default)
    if ([string]::IsNullOrWhiteSpace($v)) { return $Default }
    return $v
}

function To-BoolText([bool]$v) {
    if ($v) { return 'ON' }
    return 'off'
}

function Resolve-ToolPath([hashtable]$Config, [string]$ScriptKey) {
    $name = $Config.Scripts[$ScriptKey]
    if ([string]::IsNullOrWhiteSpace($name)) { throw "Script key not configured: $ScriptKey" }
    $p = Join-Path $PSScriptRoot $name
    if (-not (Test-Path -LiteralPath $p)) { throw "Script not found: $p" }
    return (Resolve-Path -LiteralPath $p).Path
}

function Load-Session([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @{} }
    try {
        $obj = Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json
        $h = @{}
        foreach ($p in $obj.PSObject.Properties) { $h[$p.Name] = $p.Value }
        return $h
    } catch { return @{} }
}

function Save-Session([string]$Path, [hashtable]$Session) {
    try { $Session | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $Path -Encoding UTF8 } catch {}
}

function Get-MappingPath([string]$Dir, [string]$OwnerName) {
    return (Join-Path $Dir ("mapping_{0}.csv" -f $OwnerName))
}

function Load-MappingSafe([string]$Path) {
    if (-not (Test-Path -LiteralPath $Path)) { return @() }
    return @(Import-Csv -LiteralPath $Path -Encoding UTF8)
}

function Get-FieldStats([array]$Rows, [string]$Field) {
    $total = $Rows.Count
    if ($total -eq 0 -or [string]::IsNullOrWhiteSpace($Field)) {
        return @{ Total=$total; Done=0; Pending=$total; Missing=$true }
    }
    $first = $Rows | Select-Object -First 1
    if ($null -eq $first -or -not ($first.PSObject.Properties.Name -contains $Field)) {
        return @{ Total=$total; Done=0; Pending=$total; Missing=$true }
    }
    $done = @($Rows | Where-Object { $_.$Field -eq '1' }).Count
    return @{ Total=$total; Done=$done; Pending=($total - $done); Missing=$false }
}

function Show-Status([hashtable]$Config, [array]$Rows) {
    Write-Host ''
    Write-Host '===== Current mapping status =====' -ForegroundColor Cyan
    if ($Rows.Count -eq 0) {
        Write-Host '  mapping csv not found yet, or it has no rows.' -ForegroundColor Yellow
        return
    }
    Write-Host ("  Rows: {0}" -f $Rows.Count)
    foreach ($p in $Config.PhaseOrder) {
        if ([string]::IsNullOrWhiteSpace($p.Field)) { continue }
        $s = Get-FieldStats $Rows $p.Field
        if ($s.Missing) {
            Write-Host ("  {0,-22} : column missing ({1})" -f $p.Key, $p.Field) -ForegroundColor DarkYellow
        } else {
            Write-Host ("  {0,-22} : {1}/{2} done, {3} pending  [{4}]" -f $p.Key, $s.Done, $s.Total, $s.Pending, $p.Field)
        }
    }

    $toCounts = @($Rows | Group-Object TO_code | Sort-Object Name)
    if ($toCounts.Count -gt 0) {
        Write-Host ''
        Write-Host '  TO_code groups:' -ForegroundColor DarkGray
        foreach ($g in $toCounts) { Write-Host ("    {0,-8} {1}" -f $g.Name, $g.Count) -ForegroundColor DarkGray }
    }
}

function Get-RecommendPhase([hashtable]$Config, [array]$Rows) {
    if ($Rows.Count -eq 0) { return 'Mapping' }
    foreach ($p in $Config.PhaseOrder) {
        if ([string]::IsNullOrWhiteSpace($p.Field)) { continue }
        if ($p.Key -eq 'Df') { continue }
        $s = Get-FieldStats $Rows $p.Field
        if (-not $s.Missing -and $s.Pending -gt 0) { return $p.Key }
    }
    return 'Status'
}

function Ask-RunOptions([hashtable]$State) {
    Write-Host ''
    Write-Host 'Options for this run:' -ForegroundColor Cyan
    Write-Host ("  Force       : {0}" -f (To-BoolText $State.Force))
    Write-Host ("  Interactive : {0}" -f (To-BoolText $State.Interactive))
    Write-Host ("  NoResize    : {0}" -f (To-BoolText $State.NoResize))
    Write-Host ("  RefreshUrls : {0}" -f (To-BoolText $State.RefreshUrls))
    Write-Host ("  Window      : {0}x{1}" -f $State.WindowWidth, $State.WindowHeight)
    Write-Host ("  CropPx      : {0}" -f $State.CropPx)
    if ($State.TargetIds.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f ($State.TargetIds -join ', ')) }
    Write-Host ''
    Write-Host '  f=toggle Force, i=toggle Interactive, n=toggle NoResize, r=toggle RefreshUrls'
    Write-Host '  w=window size, c=crop px, t=target IDs, Enter=continue'
    while ($true) {
        $x = Read-Host 'option'
        if ([string]::IsNullOrWhiteSpace($x)) { break }
        switch -Regex ($x.Trim().ToLower()) {
            '^f$' { $State.Force = -not $State.Force }
            '^i$' { $State.Interactive = -not $State.Interactive }
            '^n$' { $State.NoResize = -not $State.NoResize }
            '^r$' { $State.RefreshUrls = -not $State.RefreshUrls }
            '^w$' {
                $v = Read-Choice 'Window size, e.g. 1050x761' ("{0}x{1}" -f $State.WindowWidth, $State.WindowHeight)
                if ($v -match '^\s*(\d+)\s*[xX, ]\s*(\d+)\s*$') {
                    $State.WindowWidth  = [int]$Matches[1]
                    $State.WindowHeight = [int]$Matches[2]
                } else { Write-Host '  invalid size' -ForegroundColor Yellow }
            }
            '^c$' {
                $v = Read-Choice 'CropPx' ([string]$State.CropPx)
                if ($v -match '^\d+$') { $State.CropPx = [int]$v }
            }
            '^t$' {
                $v = Read-Choice 'Target IDs, comma-separated. Empty = all' ($State.TargetIds -join ',')
                if ([string]::IsNullOrWhiteSpace($v)) { $State.TargetIds = @() }
                else { $State.TargetIds = @($v -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ }) }
            }
            default { Write-Host '  unknown option' -ForegroundColor Yellow }
        }
        Write-Host ("  Force={0}, Interactive={1}, NoResize={2}, RefreshUrls={3}, Window={4}x{5}, CropPx={6}, TargetIds={7}" -f `
            (To-BoolText $State.Force), (To-BoolText $State.Interactive), (To-BoolText $State.NoResize), (To-BoolText $State.RefreshUrls), `
            $State.WindowWidth, $State.WindowHeight, $State.CropPx, ($(if ($State.TargetIds.Count -gt 0) { $State.TargetIds -join ',' } else { 'all' })))
    }
}

function Invoke-ToolPhase([string]$PhaseKey, [hashtable]$Config, [hashtable]$State) {
    $common = Resolve-ToolPath $Config 'Common'
    $base = @{
        WorkDir      = $State.WorkDir
        Owner        = $State.Owner
    }

    if ($PhaseKey -eq 'Mapping') {
        $p = Resolve-ToolPath $Config 'GenerateMapping'
        $args = $base.Clone()
        if ($State.Force) { $args.Force = $true }
        Write-Host ("[RUN] {0}" -f (Split-Path $p -Leaf)) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Excel') {
        $p = Resolve-ToolPath $Config 'Excel'
        $args = $base.Clone()
        if ($State.Force) { $args.Force = $true }
        Write-Host ("[RUN] {0}" -f (Split-Path $p -Leaf)) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'HmGift' -or $PhaseKey -eq 'HmGfix') {
        $p = Resolve-ToolPath $Config 'Hm'
        $stage = if ($PhaseKey -eq 'HmGift') { 'GIFT' } else { 'GFIX' }
        $args = $base.Clone()
        $args.Stage = $stage
        $args.CropPx = $State.CropPx
        $args.WindowWidth = $State.WindowWidth
        $args.WindowHeight = $State.WindowHeight
        $args.ActionWaitMs = $Config.Timing.ActionWaitMs
        $args.ResultWaitSec = $Config.Timing.ResultWaitSec
        $args.TabsToCorrelid = $Config.Hm.TabsToCorrelid
        $args.TabsBackFromSearch = $Config.Hm.TabsBackFromSearch
        $args.TabsBackToInput = $Config.Hm.TabsBackToInput
        $args.CommonScript = $common
        if ($State.TargetIds.Count -gt 0) { $args.TargetIds = $State.TargetIds }
        if ($State.Force) { $args.Force = $true }
        if ($State.Interactive) { $args.Interactive = $true }
        if ($State.NoResize) { $args.NoResize = $true }
        Write-Host ("[RUN] HmSnap {0}" -f $stage) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'MqGift') {
        $p = Resolve-ToolPath $Config 'Mq'
        $args = $base.Clone()
        $args.CropPx = $State.CropPx
        $args.WindowWidth = $State.WindowWidth
        $args.WindowHeight = $State.WindowHeight
        $args.ActionWaitMs = $Config.Timing.ActionWaitMs
        $args.ResultWaitSec = $Config.Timing.ResultWaitSec
        $args.TabsToInquiry = $Config.Mq.TabsToInquiry
        $args.TabsToCorrelid = $Config.Mq.TabsToCorrelid
        $args.CommonScript = $common
        if ($State.TargetIds.Count -gt 0) { $args.TargetIds = $State.TargetIds }
        if ($State.Force) { $args.Force = $true }
        if ($State.Interactive) { $args.Interactive = $true }
        if ($State.NoResize) { $args.NoResize = $true }
        Write-Host '[RUN] MqSnap GIFT' -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'JenkinsGift' -or $PhaseKey -eq 'JenkinsGfix' -or $PhaseKey -eq 'NoGfix') {
        $p = Resolve-ToolPath $Config 'Jenkins'
        $mode = switch ($PhaseKey) {
            'JenkinsGift' { 'GiftRecv' }
            'JenkinsGfix' { 'GfixRecv' }
            'NoGfix' { 'NoGfix' }
        }
        $args = $base.Clone()
        $args.Mode = $mode
        $args.CropPx = $State.CropPx
        $args.WindowWidth = $State.WindowWidth
        $args.WindowHeight = $State.WindowHeight
        $args.ActionWaitMs = $Config.Timing.ActionWaitMs
        $args.ResultWaitMs = $Config.Timing.ResultWaitMs
        $args.CommonScript = $common
        if ($State.TargetIds.Count -gt 0) { $args.TargetIds = $State.TargetIds }
        if ($State.Force) { $args.Force = $true }
        if ($State.Interactive) { $args.Interactive = $true }
        if ($State.NoResize) { $args.NoResize = $true }
        if ($State.RefreshUrls) { $args.RefreshUrls = $true }
        Write-Host ("[RUN] JenkinsSnap {0}" -f $mode) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Crop') {
        $p = Resolve-ToolPath $Config 'Crop'
        $dir = Join-Path $State.WorkDir 'snap'
        $args = @{ Dir = $dir; CropPx = $State.CropPx; Recurse = $true }
        if ($State.Force) { $args.Force = $true }
        Write-Host ("[RUN] Crop-Snap {0}" -f $dir) -ForegroundColor Green
        if ($State.DryRun) { $args; return }
        & $p @args
        return
    }

    if ($PhaseKey -eq 'Df') {
        Write-Host '[INFO] DF phase is not implemented in this launcher yet.' -ForegroundColor Yellow
        return
    }

    if ($PhaseKey -eq 'Status') { return }
    throw "Unknown phase: $PhaseKey"
}

# Load config
if ([string]::IsNullOrWhiteSpace($ConfigPath)) { $ConfigPath = Join-Path $PSScriptRoot 'SnapConfig.psd1' }
if (-not (Test-Path -LiteralPath $ConfigPath)) { throw "Config not found: $ConfigPath" }
$Config = Import-PowerShellDataFile -LiteralPath $ConfigPath

$sessionPath = Join-Path $PSScriptRoot 'snap_session.json'
$session = Load-Session $sessionPath

if ([string]::IsNullOrWhiteSpace($Owner)) {
    if ($session.ContainsKey('Owner') -and -not [string]::IsNullOrWhiteSpace([string]$session.Owner)) { $Owner = [string]$session.Owner }
    else { $Owner = [string]$Config.DefaultOwner }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $candidate = ''
    if ($session.ContainsKey('WorkDir')) { $candidate = [string]$session.WorkDir }
    if ([string]::IsNullOrWhiteSpace($candidate)) { $candidate = [string]$Config.DefaultWorkDir }
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
        $WorkDir = Read-Choice 'WorkDir path' $candidate
    } else {
        $WorkDir = Read-Host 'WorkDir path'
    }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { throw 'WorkDir is empty.' }
if (-not (Test-Path -LiteralPath $WorkDir)) { throw "WorkDir not found: $WorkDir" }

if ($WindowWidth -le 0)  { $WindowWidth  = [int]$Config.Window.Width }
if ($WindowHeight -le 0) { $WindowHeight = [int]$Config.Window.Height }
if ($CropPx -lt 0)       { $CropPx       = [int]$Config.Window.CropPx }

$flatTargets = @()
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $flatTargets += $v }
    }
}
$TargetIds = @($flatTargets | Select-Object -Unique)

$state = @{
    WorkDir      = $WorkDir
    Owner        = $Owner
    WindowWidth  = $WindowWidth
    WindowHeight = $WindowHeight
    CropPx       = $CropPx
    TargetIds    = $TargetIds
    Force        = [bool]$Force.IsPresent
    Interactive  = [bool]$Interactive.IsPresent
    NoResize     = [bool]$NoResize.IsPresent
    RefreshUrls  = [bool]$RefreshUrls.IsPresent
    DryRun       = [bool]$DryRun.IsPresent
}

$session.WorkDir = $WorkDir
$session.Owner = $Owner
$session.WindowWidth = $WindowWidth
$session.WindowHeight = $WindowHeight
$session.CropPx = $CropPx
Save-Session $sessionPath $session

$mappingPath = Get-MappingPath $WorkDir $Owner
$mappingRows = Load-MappingSafe $mappingPath

Write-Host ''
Write-Host '===== Snap Launcher =====' -ForegroundColor Green
Write-Host ("  WorkDir : {0}" -f $WorkDir)
Write-Host ("  Owner   : {0}" -f $Owner)
Write-Host ("  Mapping : {0}" -f $mappingPath)
Write-Host ("  Window  : {0}x{1}, CropPx={2}" -f $WindowWidth, $WindowHeight, $CropPx)

if ($Phase -ne 'Menu') {
    if ($Phase -eq 'Status') { Show-Status $Config $mappingRows; return }
    Invoke-ToolPhase $Phase $Config $state
    return
}

while ($true) {
    $mappingRows = Load-MappingSafe $mappingPath
    Show-Status $Config $mappingRows
    $rec = Get-RecommendPhase $Config $mappingRows

    Write-Host ''
    Write-Host ("Recommended next: {0}" -f $rec) -ForegroundColor Green
    Write-Host ''
    Write-Host 'Choose phase:' -ForegroundColor Cyan
    Write-Host '  1  Mapping        mapping 生成 / 更新'
    Write-Host '  2  Excel          Excel 証跡'
    Write-Host '  3  HmGift         GIFT HM 証跡'
    Write-Host '  4  MqGift         GIFT MQ 証跡'
    Write-Host '  5  JenkinsGift    GIFT Jenkins 証跡 + DL'
    Write-Host '  6  NoGfix         no-GFIX 証跡'
    Write-Host '  7  HmGfix         GFIX HM 証跡'
    Write-Host '  8  JenkinsGfix    GFIX Jenkins 証跡 + DL'
    Write-Host '  9  Crop           old png crop batch'
    Write-Host '  s  Status only'
    Write-Host '  q  Quit'
    $ans = Read-Choice 'phase' $rec
    $key = switch ($ans.Trim().ToLower()) {
        '1' { 'Mapping' }
        '2' { 'Excel' }
        '3' { 'HmGift' }
        '4' { 'MqGift' }
        '5' { 'JenkinsGift' }
        '6' { 'NoGfix' }
        '7' { 'HmGfix' }
        '8' { 'JenkinsGfix' }
        '9' { 'Crop' }
        's' { 'Status' }
        'q' { 'Quit' }
        default { $ans.Trim() }
    }
    if ($key -eq 'Quit') { break }
    if ($key -eq 'Status') { continue }

    Ask-RunOptions $state
    Invoke-ToolPhase $key $Config $state

    $session.WorkDir = $state.WorkDir
    $session.Owner = $state.Owner
    $session.WindowWidth = $state.WindowWidth
    $session.WindowHeight = $state.WindowHeight
    $session.CropPx = $state.CropPx
    Save-Session $sessionPath $session

    Write-Host ''
    Write-Host 'Back to launcher menu. Enter to refresh / q to quit : ' -ForegroundColor Magenta -NoNewline
    $again = Read-Host
    if ($again -eq 'q') { break }
}
