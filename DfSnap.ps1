# ============================================================
#  DfSnap.ps1   (Phase: DfSnap)   -- UTF-8, NO BOM, ASCII source.
#
#  For each pending mapping row:
#    1. Locate the GIFT and GFIX data files for the correl id.
#    2. Launch df.exe with both file paths.
#    3. Wait until df.exe's MainWindowHandle actually appears (not just a
#       single WaitForInputIdle), bring it to the front, then capture.
#    4. Save snap\DF\<Correl_ID_S>.png ; optional per-direction crop.
#    5. Close df.exe, continue to the next row.
#    6. Only on a real capture: DF_snap = 1 (atomic mapping write).
#
#  Capture strategy (-CaptureMode):
#    region     : grab a fixed screen rectangle (RECOMMENDED). df.exe is a
#                 legacy tool whose MainWindowHandle / PrintWindow are
#                 unreliable, so a fixed region is the most stable.
#    window     : capture df.exe's window rect via CopyFromScreen. If the
#                 rect looks wrong, fall back to REGION (never fullscreen).
#    fullscreen : whole primary screen.
#
#  The window shadow is not symmetric, so crop is per-direction
#  (-CropLeft/-CropTop/-CropRight/-CropBottom), not a single value.
# ============================================================

param(
    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),     # default owner (0x53B3)
    [string[]]$TargetIds = @(),
    [switch]$Force,

    [string]$DfExePath = '',
    [string]$GiftDataDir = '',
    [string]$GfixDataDir = '',
    [string]$FilePattern = '{0}*',

    [int]$LoadWaitSec = 8,
    [int]$WindowWaitSec = 15,            # how long to wait for MainWindowHandle

    [ValidateSet('region','window','fullscreen')]
    [string]$CaptureMode = 'region',

    # Fixed region (recommended on the ~1980x1020 target).
    [int]$RegionX = 120,
    [int]$RegionY = 280,
    [int]$RegionWidth  = 1250,
    [int]$RegionHeight = 657,

    # Per-direction crop in pixels (shadow is asymmetric).
    [int]$CropLeft   = 0,
    [int]$CropTop    = 0,
    [int]$CropRight  = 0,
    [int]$CropBottom = 0,

    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'Common.ps1')          # WinAPI, Take-WindowScreenshot, Ensure-Dir
. (Join-Path $scriptDir 'MappingStore.ps1')
. (Join-Path $scriptDir 'ProgressLog.ps1')
. (Join-Path $scriptDir 'ScreenRegion.ps1')

$forceFlag = [bool]$Force.IsPresent
$dryFlag   = [bool]$DryRun.IsPresent

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($GiftDataDir)) { $GiftDataDir = Join-Path $WorkDir 'DATA\GIFT' }
if ([string]::IsNullOrWhiteSpace($GfixDataDir)) { $GfixDataDir = Join-Path $WorkDir 'DATA\GFIX' }

$snapDir     = Join-Path $WorkDir 'snap\DF'
$mappingPath = Join-Path $WorkDir ('mapping_{0}.csv' -f $Owner)

Write-Host ''
Write-Host '===== DfSnap =====' -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Mapping     : {0}" -f $mappingPath)
Write-Host ("  CaptureMode : {0}" -f $CaptureMode)
if ($CaptureMode -eq 'region' -or $CaptureMode -eq 'window') {
    Write-Host ("  Region      : x={0} y={1} w={2} h={3}" -f $RegionX, $RegionY, $RegionWidth, $RegionHeight)
}
Write-Host ("  Crop L/T/R/B: {0}/{1}/{2}/{3}" -f $CropLeft, $CropTop, $CropRight, $CropBottom)
Write-Host ("  SnapDir     : {0}" -f $snapDir)
Write-Host ("  Force       : {0}   DryRun : {1}" -f $forceFlag, $dryFlag)
Write-Host ''

if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}
if (-not $dryFlag -and [string]::IsNullOrWhiteSpace($DfExePath)) {
    Write-Host '  [NOTE] DfExePath not set. To persist: VerifyConfig.psd1 -> Df.ExePath' -ForegroundColor Yellow
    $DfExePath = (Read-Host '  Path to df.exe').Trim()
}
if (-not $dryFlag -and -not [string]::IsNullOrWhiteSpace($DfExePath) -and -not (Test-Path -LiteralPath $DfExePath)) {
    Write-Host ("[ERROR] df.exe not found: {0}" -f $DfExePath) -ForegroundColor Red; exit 1
}
if (-not $dryFlag -and [string]::IsNullOrWhiteSpace($DfExePath)) {
    Write-Host '[ERROR] DfExePath required.' -ForegroundColor Red; exit 1
}

# -- mapping (MappingStore) -----------------------------------
$allRows = Import-Mapping $mappingPath
Ensure-MappingColumns -Rows $allRows | Out-Null   # DF_snap is a standard status column
$targets = ConvertTo-TargetIdList $TargetIds
$pending = @(Get-PendingRows -Rows $allRows -Field 'DF_snap' -Force $forceFlag -Targets $targets)

Write-ProgressEvent -WorkDir $WorkDir -Phase 'DfSnap' -Action 'start' -Status 'info' `
    -Message ("pending={0} mode={1}" -f $pending.Count, $CaptureMode)

if ($pending.Count -eq 0) {
    Write-Host '[DfSnap] No pending rows.' -ForegroundColor Green
    exit 0
}

# -- screen / capture helpers ---------------------------------
function Get-PrimaryBounds { return [System.Windows.Forms.Screen]::PrimaryScreen.Bounds }

# Clamp a region into the primary screen (delegates the math to the pure,
# unit-tested Resolve-ScreenRegion in ScreenRegion.ps1).
function Resolve-Region([int]$x, [int]$y, [int]$w, [int]$h) {
    $b = Get-PrimaryBounds
    $r = Resolve-ScreenRegion -X $x -Y $y -W $w -H $h `
            -BoundsX $b.X -BoundsY $b.Y -BoundsW $b.Width -BoundsH $b.Height
    return @{ X = $r.X; Y = $r.Y; W = $r.W; H = $r.H; Warn = $r.Warn }
}

function Save-RegionPng([string]$out, [int]$x, [int]$y, [int]$w, [int]$h) {
    if ($w -le 0 -or $h -le 0) { throw ("invalid region size {0}x{1}" -f $w, $h) }
    $bmp = New-Object System.Drawing.Bitmap($w, $h)
    $g   = [System.Drawing.Graphics]::FromImage($bmp)
    try {
        $g.CopyFromScreen($x, $y, 0, 0, (New-Object System.Drawing.Size($w, $h)))
        $bmp.Save($out, [System.Drawing.Imaging.ImageFormat]::Png)
    } finally { $g.Dispose(); $bmp.Dispose() }
}

# Poll until df.exe exposes a MainWindowHandle (or timeout).
function Wait-MainWindow($proc, [int]$timeoutSec) {
    $sw = [System.Diagnostics.Stopwatch]::StartNew()
    while ($sw.Elapsed.TotalSeconds -lt $timeoutSec) {
        try { $proc.Refresh() } catch {}
        $h = [IntPtr]::Zero
        try { $h = $proc.MainWindowHandle } catch {}
        if ($h -ne [IntPtr]::Zero) { return $h }
        Start-Sleep -Milliseconds 300
    }
    return [IntPtr]::Zero
}

function Bring-ToFront($hwnd) {
    if ($hwnd -eq [IntPtr]::Zero) { return }
    [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null     # SW_RESTORE
    [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 400
}

# Per-direction crop. Shadow margins differ per side, so callers pass
# explicit L/T/R/B. No-op when all zero.
function Crop-Directional([string]$path, [int]$l, [int]$t, [int]$r, [int]$b) {
    if (($l + $t + $r + $b) -le 0) { return }
    if (-not (Test-Path -LiteralPath $path)) { return }
    try {
        $orig = [System.Drawing.Image]::FromFile($path)
        $w = $orig.Width  - $l - $r
        $h = $orig.Height - $t - $b
        if ($w -le 0 -or $h -le 0) {
            $orig.Dispose()
            Write-Host '    [WARN] crop exceeds image size; skipped' -ForegroundColor Yellow
            return
        }
        $bmp = New-Object System.Drawing.Bitmap($w, $h)
        $g   = [System.Drawing.Graphics]::FromImage($bmp)
        $g.DrawImage($orig, -$l, -$t)
        $g.Dispose(); $orig.Dispose()
        $bmp.Save($path, [System.Drawing.Imaging.ImageFormat]::Png)
        $bmp.Dispose()
    } catch {
        Write-Host ("    [WARN] crop failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

# Returns the capture mode actually used ('region'/'window'/'fullscreen').
function Invoke-Capture([string]$out, $proc, $region) {
    switch ($CaptureMode) {
        'region' {
            Save-RegionPng $out $region.X $region.Y $region.W $region.H
            return 'region'
        }
        'fullscreen' {
            $b = Get-PrimaryBounds
            Save-RegionPng $out $b.X $b.Y $b.Width $b.Height
            return 'fullscreen'
        }
        'window' {
            $hwnd = [IntPtr]::Zero
            try { $hwnd = $proc.MainWindowHandle } catch {}
            if ($hwnd -eq [IntPtr]::Zero) {
                Write-Host '    [WARN] no window handle; falling back to region' -ForegroundColor Yellow
                Save-RegionPng $out $region.X $region.Y $region.W $region.H
                return 'region'
            }
            $rect = New-Object WinAPI+RECT
            [WinAPI]::GetWindowRect($hwnd, [ref]$rect) | Out-Null
            $w = $rect.Right - $rect.Left
            $h = $rect.Bottom - $rect.Top
            $valid = ($w -gt 50 -and $h -gt 50 -and $rect.Left -gt -30000 -and $rect.Top -gt -30000)
            if (-not $valid) {
                Write-Host ("    [WARN] window rect invalid ({0}x{1} @ {2},{3}); falling back to region" -f `
                    $w, $h, $rect.Left, $rect.Top) -ForegroundColor Yellow
                Save-RegionPng $out $region.X $region.Y $region.W $region.H
                return 'region'
            }
            Take-WindowScreenshot $hwnd $out
            return 'window'
        }
    }
}

function Find-DataFile([string]$baseDir, [string]$correlIdS) {
    if (-not (Test-Path -LiteralPath $baseDir)) { return $null }
    $pattern = $FilePattern -f $correlIdS
    $hits = @(Get-ChildItem -LiteralPath $baseDir -Filter $pattern -File -ErrorAction SilentlyContinue |
              Sort-Object LastWriteTime -Descending)
    if ($hits.Count -eq 0) { return $null }
    if ($hits.Count -gt 1) {
        Write-Host ("    [WARN] {0} files match '{1}'; using newest" -f $hits.Count, $pattern) -ForegroundColor Yellow
    }
    return $hits[0].FullName
}

# Resolve + warn about the region once up front.
$region = Resolve-Region $RegionX $RegionY $RegionWidth $RegionHeight
if ($region.Warn -ne '') {
    Write-Host ("  [WARN] region out of bounds -> {0}; using x={1} y={2} w={3} h={4}" -f `
        $region.Warn, $region.X, $region.Y, $region.W, $region.H) -ForegroundColor Yellow
}

if (-not $dryFlag) { Ensure-Dir $snapDir }

# -- main loop ------------------------------------------------
$cntDone = 0; $cntSkip = 0; $cntFail = 0

foreach ($row in $pending) {
    $correlIdS = [string]$row.Correl_ID_S
    if ([string]::IsNullOrWhiteSpace($correlIdS)) { continue }
    $jobName = ''
    if ($row.PSObject.Properties.Name -contains 'JOB_NAME') { $jobName = [string]$row.JOB_NAME }

    Write-Host ''
    Write-Host ("----- {0} -----" -f $correlIdS) -ForegroundColor Cyan

    $giftFile = Find-DataFile $GiftDataDir $correlIdS
    $gfixFile = Find-DataFile $GfixDataDir $correlIdS
    if ($null -eq $giftFile -or $null -eq $gfixFile) {
        $which = if ($null -eq $giftFile) { 'GIFT' } else { 'GFIX' }
        Write-Host ("  [FAIL] {0} data file not found for {1}" -f $which, $correlIdS) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'DfSnap' -CorrelIdS $correlIdS -JobName $jobName `
            -Action 'find-file' -Status 'fail' -Message ("{0} file missing" -f $which)
        $cntFail++; continue
    }
    Write-Host ("  GIFT: {0}" -f $giftFile) -ForegroundColor DarkGray
    Write-Host ("  GFIX: {0}" -f $gfixFile) -ForegroundColor DarkGray

    $outPng = Join-Path $snapDir ("{0}.png" -f $correlIdS)
    if ($dryFlag) {
        Write-Host ("  [DRY] would capture ({0}) -> {1}" -f $CaptureMode, $outPng) -ForegroundColor DarkGray
        $cntSkip++; continue
    }

    $proc = $null
    try {
        $proc = Start-Process -FilePath $DfExePath `
                              -ArgumentList ("`"{0}`"" -f $giftFile), ("`"{0}`"" -f $gfixFile) `
                              -PassThru

        $hwnd = Wait-MainWindow $proc $WindowWaitSec
        if ($hwnd -eq [IntPtr]::Zero) {
            Write-Host ("    [WARN] MainWindowHandle did not appear in {0}s" -f $WindowWaitSec) -ForegroundColor Yellow
        } else {
            Bring-ToFront $hwnd
        }
        if ($LoadWaitSec -gt 0) { Start-Sleep -Seconds $LoadWaitSec }

        $usedMode = Invoke-Capture $outPng $proc $region
        Crop-Directional $outPng $CropLeft $CropTop $CropRight $CropBottom
        Write-Host ("  [OK] saved ({0}): {1}" -f $usedMode, $outPng) -ForegroundColor Green

        $row.DF_snap = '1'
        Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'DfSnap' -CorrelIdS $correlIdS -JobName $jobName `
            -Action 'snap' -Status 'ok' -Message ("mode={0} snap\DF\{1}.png" -f $usedMode, $correlIdS)
        $cntDone++
    } catch {
        Write-Host ("  [FAIL] {0}" -f $_.Exception.Message) -ForegroundColor Red
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'DfSnap' -CorrelIdS $correlIdS -JobName $jobName `
            -Action 'snap' -Status 'fail' -Message $_.Exception.Message
        $cntFail++
    } finally {
        if ($null -ne $proc) {
            try {
                if (-not $proc.HasExited) {
                    $proc.CloseMainWindow() | Out-Null
                    if (-not $proc.WaitForExit(3000)) { $proc.Kill() }
                }
            } catch {
                Write-Host ("  [WARN] could not close df.exe: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
            }
            try { $proc.Dispose() } catch {}
        }
        Start-Sleep -Milliseconds 800
    }
}

Write-Host ''
Write-Host '===== DfSnap Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
