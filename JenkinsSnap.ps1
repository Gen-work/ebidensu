# ============================================================
#  JenkinsSnap.ps1  (Phase 5, v2 — clean rewrite)
#
#  3 modes (-Mode):
#    GiftRecv   : snap GIFT recv folder, download -> DATA/GIFT/
#                 mapping: GIFT_Jenkins_snap
#    GfixRecv   : snap GFIX recv folder, download -> DATA/GFIX/
#                 mapping: GFIX_Jenkins_snap
#    NoGfix     : snap GFIX recv folder, NO download
#                 mapping: GIFT_noGfixfile_snap
#
#  URL handling: per (GIFT|GFIX, appl) pair, captures URL from Edge address
#  bar via Ctrl+L + Ctrl+C. Cached in $work\jenkins_urls.json.
#
#  Usage:
#    .\JenkinsSnap.ps1 -Mode GiftRecv
#    .\JenkinsSnap.ps1 -Mode GiftRecv -Force
#    .\JenkinsSnap.ps1 -Mode GiftRecv -RefreshUrls
#    .\JenkinsSnap.ps1 -Mode GfixRecv -NoResize
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('GiftRecv','GfixRecv','NoGfix')]
    [string]$Mode,

    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),

    [int]$CropPx       = 6,
    [int]$WindowWidth  = 1050,
    [int]$WindowHeight = 761,
    [switch]$NoResize,

    [int]$ActionWaitMs = 500,
    [int]$ResultWaitMs = 500,

    [switch]$RefreshUrls,
    [string[]]$TargetIds = @(),
    [string]$CommonScript = "",
    [switch]$Interactive,
    [switch]$Force
)

$ErrorActionPreference = 'Stop'

# Unblock UNC files
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter "*.ps1" -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host "WorkDir path" }

# Resolve switch flags BEFORE any dot-source can clobber them
$forceFlag       = [bool]$Force.IsPresent
$refreshFlag     = [bool]$RefreshUrls.IsPresent
$noResizeFlag    = [bool]$NoResize.IsPresent
$interactiveFlag = [bool]$Interactive.IsPresent

# Optional narrow run. Accepts: -TargetIds JIGPL48S or -TargetIds JIGPL48S,JIDSL48S
$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}
function Test-TargetRow($row) {
    if ($targetSet.Count -eq 0) { return $true }
    return ($targetSet.ContainsKey([string]$row.Correl_ID_S) -or
            $targetSet.ContainsKey([string]$row.Correl_ID_M) -or
            $targetSet.ContainsKey([string]$row.JOB_NAME))
}

# Mode-specific config
$cfg = switch ($Mode) {
    'GiftRecv' { @{ Field='GIFT_Jenkins_snap';    Folder='GIFT_Jenkins';    UrlKind='GIFT'; Download=$true;  DataRoot='DATA\GIFT' } }
    'GfixRecv' { @{ Field='GFIX_Jenkins_snap';    Folder='GFIX_Jenkins';    UrlKind='GFIX'; Download=$true;  DataRoot='DATA\GFIX' } }
    'NoGfix'   { @{ Field='GIFT_noGfixfile_snap'; Folder='GIFT_noGfixfile'; UrlKind='GFIX'; Download=$false; DataRoot=$null      } }
}

Write-Host ""
Write-Host "===== JenkinsSnap (Phase 5 v2) =====" -ForegroundColor Green
Write-Host ("  Mode        : {0}" -f $Mode)
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Owner       : {0}" -f $Owner)
Write-Host ("  Field       : mapping.{0}" -f $cfg.Field)
Write-Host ("  URL kind    : {0}" -f $cfg.UrlKind)
Write-Host ("  Download    : {0}{1}" -f $cfg.Download, $(if ($cfg.Download) { " -> $($cfg.DataRoot)\<Correl_ID_S>" } else { "" }))
Write-Host ("  Window      : {0}" -f $(if ($noResizeFlag) { "no resize" } else { "${WindowWidth}x${WindowHeight}" }))
Write-Host ("  CropPx      : {0}" -f $CropPx)
Write-Host ("  Force       : {0}, RefreshUrls : {1}, Interactive : {2}" -f $forceFlag, $refreshFlag, $interactiveFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f (($targetSet.Keys | Sort-Object) -join ", ")) }
Write-Host ""

# ── Validate paths ──
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found." -ForegroundColor Red; exit 1 }
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) { Write-Host ("[ERROR] mapping not found: {0}" -f $mappingPath) -ForegroundColor Red; exit 1 }

$snapDir = Join-Path $WorkDir ("snap\{0}" -f $cfg.Folder)
if (-not (Test-Path -LiteralPath $snapDir)) { New-Item -ItemType Directory -Path $snapDir -Force | Out-Null }
$dataRoot = $null
if ($cfg.Download) {
    $dataRoot = Join-Path $WorkDir $cfg.DataRoot
    if (-not (Test-Path -LiteralPath $dataRoot)) { New-Item -ItemType Directory -Path $dataRoot -Force | Out-Null }
}

# ============================================================
# Dot-source Common.ps1 (project shared helpers)
# ============================================================
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($CommonScript)) { $candidates += $CommonScript }
$candidates += @(
    (Join-Path $PSScriptRoot "Common.ps1"),
    (Join-Path $PSScriptRoot "..\VerifyTool\Common.ps1"),
    (Join-Path (Split-Path $PSScriptRoot -Parent) "VerifyTool\Common.ps1")
)
$commonPath = $null
foreach ($c in $candidates) {
    if (Test-Path -LiteralPath $c) { $commonPath = (Resolve-Path -LiteralPath $c).Path; break }
}
if (-not $commonPath) { Write-Host "[ERROR] Common.ps1 not found." -ForegroundColor Red; exit 1 }
Write-Host ("[INFO] Loading Common.ps1 : {0}" -f $commonPath)

$savedEAP = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
. $commonPath
$ErrorActionPreference = $savedEAP
if (-not (Get-Command -Name 'Wait-PagePrepared' -ErrorAction SilentlyContinue)) {
    Write-Host "[ERROR] Common.ps1 dot-source failed." -ForegroundColor Red; exit 1
}

# Globals used by Common helpers
$Global:Timing = @{ ActionWaitMs = $ActionWaitMs; ResultWaitSec = 2 }
if (-not $Global:Shell) { $Global:Shell = New-Object -ComObject WScript.Shell }

Add-Type -AssemblyName System.Windows.Forms -ErrorAction SilentlyContinue
Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# ============================================================
# Inline helpers (no dot-source = no param-block side effects)
# ============================================================
function Invoke-CropPng {
    param([Parameter(Mandatory=$true)][string]$path, [int]$cropPx = 12)
    if (-not (Test-Path -LiteralPath $path)) { throw "File not found: $path" }
    $bytes   = [System.IO.File]::ReadAllBytes($path)
    $ms      = New-Object System.IO.MemoryStream(, $bytes)
    $tmpPath = "$path.crop.tmp"
    try {
        $orig = [System.Drawing.Image]::FromStream($ms)
        try {
            $newW = $orig.Width  - 2 * $cropPx
            $newH = $orig.Height - 2 * $cropPx
            if ($newW -le 0 -or $newH -le 0) { throw ("Image too small ({0}x{1})" -f $orig.Width, $orig.Height) }
            $bmp = New-Object System.Drawing.Bitmap($newW, $newH)
            try {
                $gfx = [System.Drawing.Graphics]::FromImage($bmp)
                try {
                    $src = New-Object System.Drawing.Rectangle($cropPx, $cropPx, $newW, $newH)
                    $dst = New-Object System.Drawing.Rectangle(0, 0, $newW, $newH)
                    $gfx.DrawImage($orig, $dst, $src, [System.Drawing.GraphicsUnit]::Pixel)
                } finally { $gfx.Dispose() }
                $bmp.Save($tmpPath, [System.Drawing.Imaging.ImageFormat]::Png)
            } finally { $bmp.Dispose() }
        } finally { $orig.Dispose() }
    } finally { $ms.Dispose() }
    Move-Item -LiteralPath $tmpPath -Destination $path -Force
}

function Move-EdgeAwayFromBorder {
    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) { return }
    try {
        if ($noResizeFlag) {
            $rect = New-Object WinAPI+RECT
            [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null
            $curW = $rect.Right  - $rect.Left
            $curH = $rect.Bottom - $rect.Top
            [WinAPI]::MoveWindow($hWnd, 40, 40, $curW, $curH, $true) | Out-Null
        } else {
            [WinAPI]::MoveWindow($hWnd, 40, 40, $WindowWidth, $WindowHeight, $true) | Out-Null
        }
        Start-Sleep -Milliseconds 400
    } catch {}
}

function Bring-ShellToFront {
    try {
        $hwnd = (Get-Process -Id $PID).MainWindowHandle
        if ($hwnd -ne [IntPtr]::Zero) {
            [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
            [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
            Start-Sleep -Milliseconds 200
        }
    } catch {}
}

function Activate-EdgeMainWindow {
    # Find Edge proc whose MainWindowHandle is non-zero — that's the actual
    # top-level browser window. Find-bar popup has no MainWindowHandle.
    $edgeProc = Get-Process msedge -ErrorAction SilentlyContinue |
                Where-Object { $_.MainWindowHandle -ne [IntPtr]::Zero } |
                Select-Object -First 1
    if (-not $edgeProc) { return [IntPtr]::Zero }
    $hwnd = $edgeProc.MainWindowHandle
    [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
    [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
    Start-Sleep -Milliseconds 300
    return $hwnd
}

function Update-MappingSnapField([string]$correlIdM, [string]$field, [array]$map, [string]$path) {
    foreach ($r in $map) {
        if ($r.Correl_ID_M -eq $correlIdM) { $r.$field = "1"; break }
    }
    $map | Export-Csv -LiteralPath $path -Encoding UTF8 -NoTypeInformation -Force
}

# ============================================================
# Load mapping & filter pending
# ============================================================
Write-Host ""
Write-Host "[Step 1] Loading mapping..." -ForegroundColor Cyan
$mapping = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Write-Host ("  Total rows : {0}" -f $mapping.Count)

$pendingItems = @()
$doneCount = 0
foreach ($r in $mapping) {
    if (-not (Test-TargetRow $r)) { continue }
    $cur = $r.($cfg.Field)
    if (-not $forceFlag) {
        if ($cur -eq "1") { $doneCount++; continue }
    }
    $pendingItems += $r
}
Write-Host ("  Pending    : {0}, Already done : {1}" -f $pendingItems.Count, $doneCount)
if ($pendingItems.Count -eq 0) {
    Write-Host "[INFO] Nothing to do. Use -Force to redo." -ForegroundColor Yellow; return
}

$applsNeeded = @($pendingItems | ForEach-Object { $_.TO_code } | Sort-Object -Unique)
Write-Host ("  Appls needed: {0}" -f ($applsNeeded -join ", "))

# ============================================================
# Step 2: URL collection / cache load
# ============================================================
$urlsPath = Join-Path $WorkDir "jenkins_urls.json"
$allUrls  = @{}

if ((Test-Path -LiteralPath $urlsPath) -and -not $refreshFlag) {
    try {
        $jsonObj = Get-Content -LiteralPath $urlsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($k in $jsonObj.PSObject.Properties.Name) {
            $allUrls[$k] = @{}
            foreach ($a in $jsonObj.$k.PSObject.Properties.Name) {
                $allUrls[$k][$a] = $jsonObj.$k.$a
            }
        }
        Write-Host ("[INFO] Loaded cached URLs from: {0}" -f $urlsPath)
    } catch {
        Write-Host ("  [WARN] URL cache load failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}
if (-not $allUrls.ContainsKey($cfg.UrlKind)) { $allUrls[$cfg.UrlKind] = @{} }

if ($refreshFlag) {
    $applsToCollect = $applsNeeded
} else {
    $applsToCollect = @($applsNeeded | Where-Object { -not $allUrls[$cfg.UrlKind].ContainsKey($_) })
}

if ($applsToCollect.Count -gt 0) {
    Write-Host ""
    Write-Host ("[Step 2] Collecting URLs for {0} appl(s): {1}" -f $applsToCollect.Count, ($applsToCollect -join ", ")) -ForegroundColor Cyan

    foreach ($appl in $applsToCollect) {
        $applItems = @($pendingItems | Where-Object { $_.TO_code -eq $appl })
        $sample    = $applItems[0].Correl_ID_S

        Bring-ShellToFront
        Write-Host ""
        Write-Host ("----- {0} / {1} -----" -f $cfg.UrlKind, $appl) -ForegroundColor Yellow
        Write-Host ("  Open the {0} {1} Jenkins folder (sample file: {2})" -f $cfg.UrlKind, $appl, $sample) -ForegroundColor Yellow
        Wait-PagePrepared ("Press Enter when {0}/{1} folder is loaded in Edge..." -f $cfg.UrlKind, $appl)
        Switch-ToEdge
        Start-Sleep -Milliseconds 500

        Send-Key "^l" 300
        Send-Key "^c" 200
        Start-Sleep -Milliseconds 250

        $clipText = ""
        try {
            $clip = Get-Clipboard -Format Text -ErrorAction SilentlyContinue
            if ($clip -is [array]) { $clip = $clip[0] }
            if ($null -ne $clip) { $clipText = $clip.ToString().Trim() }
        } catch { $clipText = "" }

        Bring-ShellToFront
        if ([string]::IsNullOrWhiteSpace($clipText) -or ($clipText -notmatch '^https?://')) {
            Write-Host ("  [ERROR] URL capture failed. Got: '{0}'" -f $clipText) -ForegroundColor Red
            return
        }
        if (-not $clipText.EndsWith('/')) { $clipText += '/' }
        $allUrls[$cfg.UrlKind][$appl] = $clipText
        Write-Host ("  Captured: {0}" -f $clipText) -ForegroundColor Green

        Send-Key "{ESC}" 100
    }

    $allUrls | ConvertTo-Json -Depth 5 | Set-Content -LiteralPath $urlsPath -Encoding UTF8
    Write-Host ""
    Write-Host ("  Saved URL cache: {0}" -f $urlsPath) -ForegroundColor DarkGray
}

foreach ($appl in $applsNeeded) {
    if (-not $allUrls[$cfg.UrlKind].ContainsKey($appl)) {
        Write-Host ("[ERROR] No URL for {0}/{1}. Use -RefreshUrls." -f $cfg.UrlKind, $appl) -ForegroundColor Red
        return
    }
}

# ============================================================
# Step 3: Main snap loop (grouped by TO_code)
# ============================================================
$grouped = $pendingItems | Group-Object TO_code | Sort-Object Name
$totalDone    = 0
$totalSkipped = 0
$totalDlFail  = 0

foreach ($g in $grouped) {
    $appl    = $g.Name
    $items   = $g.Group
    $baseUrl = $allUrls[$cfg.UrlKind][$appl]

    Bring-ShellToFront
    Write-Host ""
    Write-Host ("####################################################################") -ForegroundColor Cyan
    Write-Host ("##  Jenkins {0} - appl: {1}  ({2} item(s))" -f $Mode, $appl, $items.Count) -ForegroundColor Cyan
    Write-Host ("####################################################################") -ForegroundColor Cyan
    Write-Host ("    Base URL : {0}" -f $baseUrl) -ForegroundColor DarkGray
    Wait-PagePrepared ("Press Enter when {0}/{1} Jenkins folder is in Edge." -f $cfg.UrlKind, $appl)
    Switch-ToEdge
    Move-EdgeAwayFromBorder
    Click-PageBody

    $idx = 0
    foreach ($item in $items) {
        $idx++
        Write-Host ""
        Write-Host ("=" * 72) -ForegroundColor White
        Write-Host ("  [{0} {1}/{2}] JOB:{3} | Correl_ID_S:{4} | TO:{5}" -f `
            $Mode, $idx, $items.Count, $item.JOB_NAME, $item.Correl_ID_S, $appl) -ForegroundColor White
        Write-Host ("=" * 72) -ForegroundColor White

        if ($interactiveFlag) {
            Bring-ShellToFront
            Write-Host "  Enter=run / s=skip / q=quit : " -ForegroundColor Magenta -NoNewline
            $resp = Read-Host
            if ($resp -eq 'q') { return }
            if ($resp -eq 's') { Write-Host "  -> skipped" -ForegroundColor DarkYellow; $totalSkipped++; continue }
            Switch-ToEdge
            Start-Sleep -Milliseconds 300
        }

        # Open Edge find bar, paste correl id, search
        Send-CtrlF
        Paste-Replace $item.Correl_ID_S
        Send-Enter
        Start-Sleep -Milliseconds $ResultWaitMs

        # Capture Edge MAIN window (find-bar popup remains overlaid on it visually)
        $mainHwnd = Activate-EdgeMainWindow
        $outPath  = Join-Path $snapDir ("{0}.png" -f $item.Correl_ID_S)

        if ($mainHwnd -eq [IntPtr]::Zero) {
            Write-Host "    [WARN] Edge main window not found, fallback to foreground" -ForegroundColor Yellow
            Take-ForegroundScreenshot $outPath
        } else {
            $rectCap = New-Object WinAPI+RECT
            [WinAPI]::GetWindowRect($mainHwnd, [ref]$rectCap) | Out-Null
            $wCap = $rectCap.Right  - $rectCap.Left
            $hCap = $rectCap.Bottom - $rectCap.Top
            Write-Host ("    capture: hwnd={0} size={1}x{2}" -f $mainHwnd, $wCap, $hCap) -ForegroundColor DarkGray
            Take-WindowScreenshot $mainHwnd $outPath
        }

        # Crop border
        try {
            Invoke-CropPng -path $outPath -cropPx $CropPx
        } catch {
            Write-Host ("    [WARN] Crop failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        }

        # Download
        if ($cfg.Download) {
            $fileUrl = $baseUrl + $item.Correl_ID_S
            $dataDir = Join-Path $dataRoot $appl
            if (-not (Test-Path -LiteralPath $dataDir)) { New-Item -ItemType Directory -Path $dataDir -Force | Out-Null }
            $dlPath  = Join-Path $dataDir $item.Correl_ID_S
            try {
                Invoke-WebRequest -Uri $fileUrl -OutFile $dlPath -UseBasicParsing -ErrorAction Stop
                Write-Host ("    DL  : {0}" -f (Split-Path -Leaf $dlPath)) -ForegroundColor DarkGreen
            } catch {
                Write-Host ("    [WARN] DL failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                $totalDlFail++
            }
        }

        Update-MappingSnapField $item.Correl_ID_M $cfg.Field $mapping $mappingPath
        Write-Host ("    -> {0} = 1, saved {1}" -f $cfg.Field, (Split-Path -Leaf $outPath)) -ForegroundColor Green
        $totalDone++
    }

    Bring-ShellToFront
    Write-Host ""
    Write-Host ("  Appl {0} complete." -f $appl) -ForegroundColor Green
}

Bring-ShellToFront
Write-Host ""
Write-Host "===== JenkinsSnap ($Mode) Done =====" -ForegroundColor Green
Write-Host ("  Snapped : {0}" -f $totalDone)
if ($totalSkipped -gt 0) { Write-Host ("  Skipped : {0}" -f $totalSkipped) -ForegroundColor DarkGray }
if ($totalDlFail -gt 0)  { Write-Host ("  DL fail : {0}" -f $totalDlFail) -ForegroundColor Yellow }
