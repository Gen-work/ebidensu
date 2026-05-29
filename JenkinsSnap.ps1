#Requires -Version 5.1
param(
    [ValidateSet('GiftRecv','GfixRecv','NoGfix')]
    [string]$Mode           = 'GiftRecv',
    [string]$WorkDir        = '',
    [string]$Owner          = '厳',
    [string[]]$TargetIds    = @(),
    [switch]$RefreshUrls,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$Interactive,
    [switch]$NoResize,
    [int]$WindowWidth       = 1050,
    [int]$WindowHeight      = 761,
    [int]$CropPx            = 6,
    [int]$ActionWaitMs      = 500,
    [int]$ResultWaitMs      = 500,
    [string]$CommonScript   = ''
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$forceFlag    = [bool]$Force.IsPresent
$dryFlag      = [bool]$DryRun.IsPresent
$noResizeFlag = [bool]$NoResize.IsPresent
$refreshFlag  = [bool]$RefreshUrls.IsPresent

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CommonScript)) {
    $CommonScript = Join-Path $scriptDir 'Common.ps1'
}
. $CommonScript

# ── timing globals for Send-Tab / Send-Enter / etc. ──────────────────────────
$Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitMs = $ResultWaitMs }

if (-not $WorkDir) { throw '-WorkDir is required' }

# ── mode config ───────────────────────────────────────────────────────────────
$modeCfg = switch ($Mode) {
    'GiftRecv' { @{ Field='GIFT_Jenkins_snap'; Folder='GIFT_Jenkins';    UrlKind='GIFT'; Download=$true;  DataRoot='DATA\GIFT' } }
    'GfixRecv' { @{ Field='GFIX_Jenkins_snap'; Folder='GFIX_Jenkins';    UrlKind='GFIX'; Download=$true;  DataRoot='DATA\GFIX' } }
    'NoGfix'   { @{ Field='GIFT_noGfixfile_snap'; Folder='GIFT_noGfixfile'; UrlKind='GFIX'; Download=$false; DataRoot='' } }
}

$snapField = $modeCfg.Field
$snapFolder = $modeCfg.Folder

# ── mapping ───────────────────────────────────────────────────────────────────
$mappingFile = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path $mappingFile)) { throw "Mapping not found: $mappingFile" }

$allRows = @(Import-Csv $mappingFile -Encoding UTF8)

if ($TargetIds.Count -gt 0) {
    $allRows = @($allRows | Where-Object {
        $_.Correl_ID_S -in $TargetIds -or
        $_.Correl_ID_M -in $TargetIds -or
        $_.JOB_NAME    -in $TargetIds -or
        $_.Excel_NAME  -in $TargetIds
    })
}

$pending = @($allRows | Where-Object {
    $cur = [string]$_.$snapField
    $forceFlag -or -not $cur -or $cur -eq '0' -or $cur -eq ''
})

if ($pending.Count -eq 0) {
    Write-Host "[$Mode] No pending rows." -ForegroundColor Green
    exit 0
}

Write-Host "`n===== JenkinsSnap $Mode =====" -ForegroundColor Green
Write-Host "Pending rows: $($pending.Count)" -ForegroundColor Cyan

# ── paths ─────────────────────────────────────────────────────────────────────
$snapRoot = Join-Path $WorkDir 'snap'
$snapDir  = Join-Path $snapRoot $snapFolder
Ensure-Dir $snapDir

# ── URL cache (PS5.1-safe: no -AsHashtable) ──────────────────────────────────
$urlCacheFile = Join-Path $WorkDir 'jenkins_urls.json'
$urlCache = @{}
if (Test-Path $urlCacheFile) {
    try {
        $parsed = Get-Content $urlCacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
        # Convert PSCustomObject to hashtable manually (PS 5.1 compat)
        $parsed.PSObject.Properties | ForEach-Object { $urlCache[$_.Name] = $_.Value }
    } catch {}
}
$urlDirty = $false

# ── inline helpers ────────────────────────────────────────────────────────────
function Invoke-CropPng([string]$path, [int]$crop) {
    if ($crop -le 0) { return }
    if (-not (Test-Path -LiteralPath $path)) { return }
    try {
        $orig = [System.Drawing.Image]::FromFile($path)
        $w = $orig.Width  - $crop * 2
        $h = $orig.Height - $crop * 2
        if ($w -le 0 -or $h -le 0) { $orig.Dispose(); return }
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.DrawImage($orig, -$crop, -$crop)
        $g.Dispose()
        $orig.Dispose()
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    } catch {
        Write-Host ("  [WARN] crop failed: {0}" -f $_) -ForegroundColor Yellow
    }
}

function Move-EdgeToWorkPos([IntPtr]$hWnd, [int]$w, [int]$h) {
    [WinAPI]::MoveWindow($hWnd, 0, 0, $w, $h, $true) | Out-Null
    Start-Sleep -Milliseconds 200
}

function Bring-ConsoleToFront {
    $hwnd = (Get-Process -Id $PID).MainWindowHandle
    if ($hwnd -ne [IntPtr]::Zero) {
        [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
        [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
    }
}

function Get-CurrentEdgeUrl {
    # Ctrl+L copies address bar, then Ctrl+C
    Send-Key '^l' 300
    Send-Key '^a' 150
    Send-Key '^c' 300
    Send-Key '{ESC}' 200
    $url = [System.Windows.Forms.Clipboard]::GetText()
    if ($url -match '^https?://') { return $url.Trim() }
    return ''
}

function Update-MappingField([string]$correlIdS, [string]$field, [string]$val) {
    # Update in $allRows (the in-memory slice) AND re-read full CSV to patch
    $full = @(Import-Csv $mappingFile -Encoding UTF8)
    foreach ($r in $full) {
        if ([string]$r.Correl_ID_S -eq $correlIdS) {
            $r.$field = $val
        }
    }
    $full | Export-Csv $mappingFile -NoTypeInformation -Encoding UTF8
}

# ── Get Edge hwnd (activate then read foreground) ─────────────────────────────
function Get-EdgeHwndNow {
    $hWnd = Activate-EdgeWindow
    if ($hWnd -eq [IntPtr]::Zero) {
        Write-Host '  [WARN] Edge not found.' -ForegroundColor Yellow
    }
    return $hWnd
}

# ── Group pending rows by TO_code for URL collection ─────────────────────────
# TO_code = the application code portion used to filter Jenkins folder
# We use Excel_NAME or JOB_NAME prefix (first 8 chars) as the group key.
# Actually we group by the Correl_ID_S prefix (BIZ code = first 3 chars e.g. JRV).
# User navigates to the Jenkins folder for each group once.
$groups = @{}
foreach ($row in $pending) {
    # Try to derive a group key — use first 3 chars of JOB_NAME as BIZ code
    $jobName = [string]$row.JOB_NAME
    $grpKey  = if ($jobName.Length -ge 3) { $jobName.Substring(0,3) } else { 'ALL' }
    if (-not $groups.ContainsKey($grpKey)) { $groups[$grpKey] = [System.Collections.Generic.List[object]]::new() }
    $groups[$grpKey].Add($row)
}

Write-Host ''
Write-Host 'Prepare Edge now:' -ForegroundColor Yellow
Write-Host "  - Open Jenkins folder for $Mode (e.g. GIFT or GFIX receive folder)"
Write-Host '  - Make sure the job list page is visible'
Write-Host 'Then press Enter. (q to quit)' -ForegroundColor Magenta
$resp = Read-Host
if ($resp -eq 'q') { exit 0 }

$edgeHwnd = Get-EdgeHwndNow
if ($edgeHwnd -eq [IntPtr]::Zero) {
    Write-Host '[ERROR] Could not get Edge window. Aborting.' -ForegroundColor Red
    exit 1
}

if (-not $noResizeFlag) {
    Move-EdgeToWorkPos $edgeHwnd $WindowWidth $WindowHeight
}

# ── per-group URL resolution ──────────────────────────────────────────────────
$resolvedUrls = @{}   # grpKey -> base URL

foreach ($grpKey in $groups.Keys) {
    $cacheKey = "${Mode}_${grpKey}"
    if (-not $refreshFlag -and $urlCache.ContainsKey($cacheKey)) {
        $resolvedUrls[$grpKey] = $urlCache[$cacheKey]
        Write-Host ("  [URL cached] {0} -> {1}" -f $grpKey, $resolvedUrls[$grpKey]) -ForegroundColor DarkGray
        continue
    }

    Write-Host ''
    Write-Host ("Navigate Edge to the Jenkins folder for BIZ=$grpKey ($Mode)") -ForegroundColor Yellow
    Write-Host 'Then press Enter to capture URL. (q to quit, s to skip)' -ForegroundColor Magenta
    $r = Read-Host
    if ($r -eq 'q') { exit 0 }
    if ($r -eq 's') { continue }

    $edgeHwnd = Get-EdgeHwndNow
    $baseUrl = Get-CurrentEdgeUrl
    if ($baseUrl) {
        $resolvedUrls[$grpKey] = $baseUrl
        $urlCache[$cacheKey]   = $baseUrl
        $urlDirty = $true
        Write-Host ("  [URL saved] {0}" -f $baseUrl) -ForegroundColor Green
    } else {
        Write-Host '  [WARN] Could not read URL from Edge address bar.' -ForegroundColor Yellow
    }
}

# ── main loop ─────────────────────────────────────────────────────────────────
$cntDone = 0
$cntSkip = 0
$cntFail = 0

foreach ($row in $pending) {
    $correl  = [string]$row.Correl_ID_S
    $jobName = [string]$row.JOB_NAME
    $grpKey  = if ($jobName.Length -ge 3) { $jobName.Substring(0,3) } else { 'ALL' }

    Write-Host ''
    Write-Host ("[$Mode] $correl ($jobName)") -ForegroundColor Yellow

    if ($dryFlag) {
        Write-Host '  [DryRun] skip' -ForegroundColor DarkGray
        $cntSkip++; continue
    }

    $snapPath = Join-Path $snapDir "$correl.png"

    # Get or navigate to correct Jenkins page
    $baseUrl = if ($resolvedUrls.ContainsKey($grpKey)) { $resolvedUrls[$grpKey] } else { '' }

    $edgeHwnd = Get-EdgeHwndNow
    if ($edgeHwnd -eq [IntPtr]::Zero) {
        Write-Host "  [FAIL] Edge not found — skip $correl" -ForegroundColor Red
        $cntFail++; continue
    }

    if (-not $noResizeFlag) {
        Move-EdgeToWorkPos $edgeHwnd $WindowWidth $WindowHeight
    }

    # Search for the correl ID using Ctrl+F
    Click-PageBody
    Send-CtrlF
    Paste-Replace $correl
    Start-Sleep -Milliseconds $ResultWaitMs
    Send-Key '{ESC}' 200

    # Take screenshot
    $edgeHwnd = [WinAPI]::GetForegroundWindow()
    if ($edgeHwnd -eq [IntPtr]::Zero) {
        Write-Host "  [FAIL] No foreground window — skip $correl" -ForegroundColor Red
        $cntFail++; continue
    }

    try {
        Take-WindowScreenshot $edgeHwnd $snapPath
        if ($CropPx -gt 0) { Invoke-CropPng $snapPath $CropPx }
        Write-Host ("  Saved: {0}" -f $snapPath) -ForegroundColor Green
    } catch {
        Write-Host ("  [FAIL] screenshot: {0}" -f $_.Exception.Message) -ForegroundColor Red
        $cntFail++; continue
    }

    # Optional: download data files (GiftRecv / GfixRecv only)
    if ($modeCfg.Download -and $baseUrl) {
        $dataDir = Join-Path $WorkDir $modeCfg.DataRoot
        Ensure-Dir $dataDir
        $fileUrl = $baseUrl.TrimEnd('/') + '/' + $correl
        try {
            $destFile = Join-Path $dataDir $correl
            if (-not (Test-Path -LiteralPath $destFile) -or $forceFlag) {
                Invoke-WebRequest -Uri $fileUrl -OutFile $destFile -UseBasicParsing -ErrorAction Stop
                Write-Host ("  DL: {0}" -f $correl) -ForegroundColor Gray
            } else {
                Write-Host ("  DL: already exists — {0}" -f $correl) -ForegroundColor DarkGray
            }
        } catch {
            Write-Host ("  [WARN] DL failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    # Update mapping
    try {
        Update-MappingField $correl $snapField '1'
        $cntDone++
    } catch {
        Write-Host ("  [WARN] mapping update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        $cntFail++
    }
}

# ── save URL cache ────────────────────────────────────────────────────────────
if ($urlDirty) {
    $urlCache | ConvertTo-Json -Depth 3 | Set-Content $urlCacheFile -Encoding UTF8
    Write-Host "[$Mode] URL cache saved." -ForegroundColor DarkGray
}

# ── summary ───────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host "===== JenkinsSnap $Mode Done =====" -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })

# Return focus to console
Bring-ConsoleToFront
