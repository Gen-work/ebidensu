#Requires -Version 5.1
param(
    [ValidateSet('GiftRecv','GfixRecv','NoGfix')]
    [string]$Mode           = 'GiftRecv',
    [string]$WorkDir        = '',
    [string]$Owner          = ([char]0x53B3),
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
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag    = [bool]$Force.IsPresent
$dryFlag      = [bool]$DryRun.IsPresent
$noResizeFlag = [bool]$NoResize.IsPresent
$refreshFlag  = [bool]$RefreshUrls.IsPresent

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CommonScript)) {
    $CommonScript = Join-Path $scriptDir 'Common.ps1'
}
. $CommonScript
. (Join-Path $scriptDir 'MappingStore.ps1')
. (Join-Path $scriptDir 'ProgressLog.ps1')
. (Join-Path $scriptDir 'JenkinsDownload.ps1')

$Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitMs = $ResultWaitMs }

if (-not $WorkDir) { throw '-WorkDir is required' }

# ── mode config ───────────────────────────────────────────────────────────────
# Field    : CSV column that tracks completion for this mode
# Folder   : subfolder under snap\ for screenshots
# GroupCol : CSV column to group rows by (navigate to Jenkins once per group)
# SearchCol: CSV column value to Ctrl+F search on the Jenkins page
$modeCfg = switch ($Mode) {
    'GiftRecv' { @{
        Field     = 'GIFT_Jenkins_snap'
        Folder    = 'GIFT_Jenkins'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
    'GfixRecv' { @{
        Field     = 'GFIX_Jenkins_snap'
        Folder    = 'GFIX_Jenkins'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
    'NoGfix'   { @{
        Field     = 'GIFT_noGfixfile_snap'
        Folder    = 'GIFT_noGfixfile'
        GroupCol  = 'TO_code'
        SearchCol = 'CORREL_ID_S'
        JOB       = 'JOB_NAME'
    }}
}

$snapField  = $modeCfg.Field
$snapFolder = $modeCfg.Folder
$groupCol   = $modeCfg.GroupCol
$searchCol  = $modeCfg.SearchCol
$job        = $modeCfg.JOB

# ── mapping ───────────────────────────────────────────────────────────────────
$mappingFile = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)

# MappingStore: single source of truth for read/filter/write.
# $allRows is the FULL set (so we never drop non-target rows on write);
# $pending holds references INTO $allRows, so mutating a pending row and
# then Export-MappingAtomic $allRows persists exactly that change.
$allRows = Import-Mapping $mappingFile
Ensure-MappingColumns -Rows $allRows -Extra @(@{ Name = $snapField; Default = '0' }) | Out-Null

$targets = ConvertTo-TargetIdList $TargetIds
$pending = @(Get-PendingRows -Rows $allRows -Field $snapField -Force $forceFlag -Targets $targets)

Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -Action 'start' -Status 'info' `
    -Message ("pending={0} force={1} targets=[{2}]" -f $pending.Count, $forceFlag, ($targets -join ','))

if ($pending.Count -eq 0) {
    Write-Host "[$Mode] No pending rows." -ForegroundColor Green
    exit 0
}

Write-Host "`n===== JenkinsSnap $Mode =====" -ForegroundColor Green
Write-Host "Pending rows: $($pending.Count)" -ForegroundColor Cyan

# ── paths ─────────────────────────────────────────────────────────────────────
$snapDir = Join-Path (Join-Path $WorkDir 'snap') $snapFolder
$dataRoot = Join-Path $WorkDir 'DATA'
Ensure-Dir $snapDir
if ($Mode -in 'GiftRecv','GfixRecv') { Ensure-Dir $dataRoot }

# ── URL cache (PS5.1-safe: no -AsHashtable) ──────────────────────────────────
$urlCacheFile = Join-Path $WorkDir 'jenkins_urls.json'
$urlCache = @{}
if (Test-Path $urlCacheFile) {
    try {
        $parsed = Get-Content $urlCacheFile -Raw -Encoding UTF8 | ConvertFrom-Json
        $parsed.PSObject.Properties | ForEach-Object { $urlCache[$_.Name] = $_.Value }
    } catch {}
}
$urlDirty = $false

# ── inline helpers ────────────────────────────────────────────────────────────
function Invoke-CropPng([string]$path, [int]$crop) {
    if ($crop -le 0 -or -not (Test-Path -LiteralPath $path)) { return }
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

function Get-EdgeMainWindowHandle {
    $edge = @(Get-Process -Name 'msedge' -ErrorAction SilentlyContinue |
        Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
        Select-Object -First 1)
    if ($edge.Count -gt 0) { return [IntPtr]$edge[0].MainWindowHandle }
    return [IntPtr]::Zero
}

function Activate-JenkinsEdgeWindow {
    $hWnd = Get-EdgeMainWindowHandle
    if ($hWnd -eq [IntPtr]::Zero) {
        [void]$Shell.AppActivate('Microsoft Edge')
        Start-Sleep -Milliseconds 700
        $hWnd = Get-EdgeMainWindowHandle
    }
    if ($hWnd -eq [IntPtr]::Zero) {
        Write-Host '  [WARN] Edge window not found.' -ForegroundColor Yellow
        return [IntPtr]::Zero
    }

    [WinAPI]::ShowWindowAsync($hWnd, 9) | Out-Null
    [WinAPI]::SetForegroundWindow($hWnd) | Out-Null
    Start-Sleep -Milliseconds 400
    return $hWnd
}

function Move-EdgeToWorkPos([IntPtr]$hWnd) {
    if ($hWnd -eq [IntPtr]::Zero) { return }
    [WinAPI]::MoveWindow($hWnd, 0, 0, $WindowWidth, $WindowHeight, $true) | Out-Null
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
    Send-Key '^l' 300
    Send-Key '^a' 150
    Send-Key '^c' 300
    Send-Key '{ESC}' 200
    $url = [System.Windows.Forms.Clipboard]::GetText()
    if ($url -match '^https?://') { return $url.Trim() }
    return ''
}

# ── Group pending rows by TO_code ─────────────────────────────────────────────
# Build an ordered list of distinct TO_code values (preserving first-seen order)
$groupOrder = [System.Collections.Generic.List[string]]::new()
$groupMap   = @{}   # TO_code -> list of rows
foreach ($row in $pending) {
    $grp = [string]$row.$groupCol
    if ([string]::IsNullOrWhiteSpace($grp)) { $grp = '(unknown)' }
    if (-not $groupMap.ContainsKey($grp)) {
        $groupMap[$grp] = [System.Collections.Generic.List[object]]::new()
        $groupOrder.Add($grp)
    }
    $groupMap[$grp].Add($row)
}

Write-Host ''
Write-Host 'Groups by TO_code:' -ForegroundColor DarkGray
foreach ($g in $groupOrder) {
    Write-Host ("  {0,-12} : {1} rows" -f $g, $groupMap[$g].Count) -ForegroundColor DarkGray
}

# ── Process each TO_code group ────────────────────────────────────────────────
$cntDone = 0
$cntSkip = 0
$cntFail = 0

foreach ($toCode in $groupOrder) {
    $rows = $groupMap[$toCode]

    Write-Host ''
    Write-Host ("===== TO_code: {0}  ({1} rows) =====" -f $toCode, $rows.Count) -ForegroundColor Cyan

    # ── DryRun: report the plan only; never open Edge or prompt. ──────────────
    if ($dryFlag) {
        foreach ($row in $rows) {
            $correl = [string]$row.Correl_ID_S
            Write-Host ("  [DryRun] would snap {0} (search: {1}) -> snap\{2}\{0}.png" -f `
                $correl, [string]$row.$searchCol, $snapFolder) -ForegroundColor DarkGray
            $cntSkip++
        }
        continue
    }

    # ── Navigate Edge to this system's Jenkins folder ─────────────────────────
    $cacheKey  = "{0}_{1}" -f $Mode, $toCode
    $cachedUrl = if ($urlCache.ContainsKey($cacheKey)) { $urlCache[$cacheKey] } else { '' }

    if (-not $refreshFlag -and $cachedUrl) {
        Write-Host ("  [cached URL] {0}" -f $cachedUrl) -ForegroundColor DarkGray
        Write-Host ("  Navigating Edge to cached Jenkins folder URL for {0}." -f $toCode) -ForegroundColor Yellow
        Write-Host '  Press Enter to confirm, r+Enter to refresh URL, q=quit.' -ForegroundColor Magenta
        $resp = Read-Host
        if ($resp -eq 'q') { exit 0 }

        if ($resp -eq 'r') {
            # force re-navigate
            $cachedUrl = ''
        } else {
            $edgeHwnd = Activate-JenkinsEdgeWindow
            if (-not $noResizeFlag) { Move-EdgeToWorkPos $edgeHwnd }
            # Navigate to cached URL
            Send-Key '^l' 300
            [System.Windows.Forms.Clipboard]::SetText($cachedUrl)
            Send-Key '^v' 200
            Send-Key '{ENTER}' ($ResultWaitMs * 4)
        }
    }

    if (-not $cachedUrl) {
        Write-Host ''
        Write-Host (">>> Open Edge to the [{0}] Jenkins folder page." -f $toCode) -ForegroundColor Yellow
        Write-Host '    (e.g. for IDS: the IDS Jenkins job list page)' -ForegroundColor Yellow
        Write-Host '    Press Enter when ready. (q=quit)' -ForegroundColor Magenta
        $resp = Read-Host
        if ($resp -eq 'q') { exit 0 }

        $edgeHwnd = Activate-JenkinsEdgeWindow
        if (-not $noResizeFlag) { Move-EdgeToWorkPos $edgeHwnd }

        # Capture and cache the URL
        $capturedUrl = Get-CurrentEdgeUrl
        if ($capturedUrl) {
            $urlCache[$cacheKey] = $capturedUrl
            $urlDirty = $true
            Write-Host ("  [URL saved] {0}" -f $capturedUrl) -ForegroundColor Green
        } else {
            Write-Host '  [WARN] Could not capture URL. Continuing without cache.' -ForegroundColor Yellow
        }
    }

    # ── Per-row screenshot loop ───────────────────────────────────────────────
    foreach ($row in $rows) {
        $correl     = [string]$row.Correl_ID_S
        $searchJob        = [string]$row.$job
        $searchTerm = [string]$row.$searchCol

        Write-Host ''
        Write-Host ("  [$correl] search: $searchTerm") -ForegroundColor White

        $snapPath = Join-Path $snapDir "$correl.png"

        $edgeHwnd = Activate-JenkinsEdgeWindow
        if ($edgeHwnd -eq [IntPtr]::Zero) {
            Write-Host "    [FAIL] Edge not found" -ForegroundColor Red
            $cntFail++; continue
        }

        # Ctrl+F search for CORREL_ID_S; leave find bar open so the screenshot
        # shows the highlighted match (ESC is sent after the screenshot below).
        Click-PageBody
        Send-CtrlF
        Paste-Replace $searchTerm
        Start-Sleep -Milliseconds $ResultWaitMs

        # Take screenshot from the Edge handle we just activated; never use the
        # foreground handle here because the console can remain foreground after
        # Read-Host on some terminals.
        if ($edgeHwnd -eq [IntPtr]::Zero) {
            Write-Host "    [FAIL] No Edge window handle" -ForegroundColor Red
            $cntFail++; continue
        }

        try {
            Take-WindowScreenshot $edgeHwnd $snapPath
            if ($CropPx -gt 0) { Invoke-CropPng $snapPath $CropPx }
            # Dismiss find bar now that the screenshot is captured.
            Send-Key '{ESC}' 200
            Write-Host ("    Saved: snap\{0}\{1}.png" -f $snapFolder, $correl) -ForegroundColor Green
        } catch {
            Write-Host ("    [FAIL] screenshot: {0}" -f $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                -JobName $searchJob -Action 'screenshot' -Status 'fail' -Message $_.Exception.Message
            $cntFail++; continue
        }

        # Download the matching Jenkins receive file into DATA\GIFT or DATA\GFIX.
        if ($Mode -in 'GiftRecv','GfixRecv') {
            try {
                Click-PageBody
                $pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'
                $pageText = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ResultWaitMs
                $folderUrl = Get-CurrentEdgeUrl
                if ([string]::IsNullOrWhiteSpace($folderUrl)) { $folderUrl = $cachedUrl }

                if ([string]::IsNullOrWhiteSpace($pageText) -or [string]::IsNullOrWhiteSpace($folderUrl)) {
                    Write-Host '    [WARN] download skipped: page text or Jenkins folder URL is empty' -ForegroundColor Yellow
                    Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                        -JobName $searchJob -Action 'download' -Status 'warn' -Message 'page text or folder URL is empty'
                } else {
                    $dl = Invoke-JenkinsFileDownload -WorkDir $WorkDir -Mode $Mode -FolderUrl $folderUrl `
                        -PageText $pageText -CorrelId $correl -JobName $searchJob -Force:$forceFlag `
                        -ParserScript (Join-Path $scriptDir 'Parse-JenkinsList.ps1')
                    Write-Host ("    Jenkins files: found={0} matched={1} downloaded={2} skipped={3} failed={4} -> DATA\{5}" -f `
                        $dl.Found, $dl.Matched, $dl.Downloaded, $dl.Skipped, $dl.Failed, $dl.DataKind) -ForegroundColor DarkGray
                    foreach ($f in @($dl.Files)) {
                        $color = if ($f.Status -eq 'ok') { 'Gray' } elseif ($f.Status -eq 'skip') { 'DarkGray' } else { 'Yellow' }
                        Write-Host ("      [{0}] {1}" -f $f.Status, $f.Name) -ForegroundColor $color
                    }
                    $dlStatus = if ($dl.Failed -gt 0) { 'fail' } elseif ($dl.Matched -eq 0) { 'warn' } else { 'ok' }
                    $dlMessage = "DATA\$($dl.DataKind): found=$($dl.Found) matched=$($dl.Matched) downloaded=$($dl.Downloaded) skipped=$($dl.Skipped) failed=$($dl.Failed)"
                    Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                        -JobName $searchJob -Action 'download' -Status $dlStatus -Message $dlMessage
                }
            } catch {
                Write-Host ("    [WARN] download failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                    -JobName $searchJob -Action 'download' -Status 'fail' -Message $_.Exception.Message
            }
        }

        # Mark ONLY this correl's snap column done, then persist atomically.
        # ($row is a reference into $allRows, so this updates the full set.)
        try {
            $row.$snapField = '1'
            Export-MappingAtomic -Rows $allRows -Path $mappingFile | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                -JobName $searchJob -Action 'snap' -Status 'ok' `
                -Message ("snap\{0}\{1}.png" -f $snapFolder, $correl)
            $cntDone++
        } catch {
            Write-Host ("    [WARN] mapping update failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            Write-ProgressEvent -WorkDir $WorkDir -Phase "Jenkins:$Mode" -CorrelIdS $correl `
                -JobName $searchJob -Action 'mapping' -Status 'fail' -Message $_.Exception.Message
        }
    }
}

# ── save URL cache ────────────────────────────────────────────────────────────
if ($urlDirty) {
    $urlCache | ConvertTo-Json -Depth 3 | Set-Content $urlCacheFile -Encoding UTF8
    Write-Host ''
    Write-Host "[$Mode] URL cache saved: jenkins_urls.json" -ForegroundColor DarkGray
}

# ── summary ───────────────────────────────────────────────────────────────────
Write-Host ''
Write-Host ("===== JenkinsSnap $Mode Done =====") -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })

Bring-ConsoleToFront
