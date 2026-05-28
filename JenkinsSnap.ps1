#Requires -Version 5.1
param(
    [ValidateSet('GiftRecv','GfixRecv','NoGfix')]
    [string]$Mode        = 'GiftRecv',
    [string]$WorkDir     = '',
    [string]$Owner       = '厳',
    [string[]]$TargetIds = @(),
    [switch]$RefreshUrls,
    [switch]$Force,
    [switch]$DryRun,
    [int]$WindowWidth    = 1050,
    [int]$WindowHeight   = 761,
    [int]$CropPx         = 6,
    [switch]$NoResize
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# ── config ──────────────────────────────────────────────────────────────────
$scriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'Common.ps1')

$cfgPath = Join-Path $scriptDir 'VerifyConfig.psd1'
$cfg = Import-PowerShellDataFile $cfgPath

if (-not $WorkDir) { throw '-WorkDir is required' }

$mappingFile = Join-Path $WorkDir ($cfg.Paths.MappingPattern -f $Owner)
if (-not (Test-Path $mappingFile)) { throw "Mapping not found: $mappingFile" }

$snapRoot = Join-Path $WorkDir $cfg.Paths.SnapDir
$dataRoot = Join-Path $WorkDir $cfg.Paths.FileDir

# folder names per mode
$snapFolderMap = @{
    GiftRecv = 'GIFT_Jenkins'
    GfixRecv = 'GFIX_Jenkins'
    NoGfix   = 'GIFT_noGfixfile'
}
$snapFolder = $snapFolderMap[$Mode]

# CSV field updated per mode
$fieldMap = @{
    GiftRecv = 'GIFT_Jenkins_snap'
    GfixRecv = 'GFIX_Jenkins_snap'
    NoGfix   = 'GIFT_noGfixfile_snap'
}
$snapField = $fieldMap[$Mode]

$urlCacheFile = Join-Path $WorkDir 'jenkins_urls.json'

# ── helpers ─────────────────────────────────────────────────────────────────
function Load-UrlCache {
    if (Test-Path $urlCacheFile) {
        try { return Get-Content $urlCacheFile -Raw | ConvertFrom-Json -AsHashtable }
        catch { return @{} }
    }
    return @{}
}

function Save-UrlCache([hashtable]$cache) {
    $cache | ConvertTo-Json -Depth 3 | Set-Content $urlCacheFile -Encoding UTF8
}

function Get-SnapPath([string]$folder, [string]$correl) {
    return Join-Path $snapRoot $folder "$correl.png"
}

function Ensure-SnapDir([string]$folder) {
    $dir = Join-Path $snapRoot $folder
    if (-not (Test-Path $dir)) { New-Item $dir -ItemType Directory | Out-Null }
    return $dir
}

# ── read mapping ─────────────────────────────────────────────────────────────
$rows = Import-Csv $mappingFile -Encoding UTF8
if ($TargetIds.Count -gt 0) {
    $rows = $rows | Where-Object {
        $_.Correl_ID_S -in $TargetIds -or
        $_.Correl_ID_M -in $TargetIds -or
        $_.JOB_NAME    -in $TargetIds -or
        $_.Excel_NAME  -in $TargetIds
    }
}

if ($Mode -eq 'NoGfix') {
    # NoGfix: rows that have no GFIX receive file (GIFT_noGfixfile_snap field)
    $pending = $rows | Where-Object {
        $Force -or -not $_.$snapField -or $_.$snapField -eq '0' -or $_.$snapField -eq ''
    }
} else {
    $pending = $rows | Where-Object {
        $Force -or -not $_.$snapField -or $_.$snapField -eq '0' -or $_.$snapField -eq ''
    }
}

if ($pending.Count -eq 0) {
    Write-Host "[$Mode] No pending rows." -ForegroundColor Green
    exit 0
}

Write-Host "[$Mode] $($pending.Count) row(s) pending." -ForegroundColor Cyan

$urlCache = Load-UrlCache
$dirty    = $false

# ── main loop ────────────────────────────────────────────────────────────────
foreach ($row in $pending) {
    $correl  = $row.Correl_ID_S
    $jobName = $row.JOB_NAME

    Write-Host "`n[$Mode] $correl ($jobName)" -ForegroundColor Yellow

    if ($DryRun) {
        Write-Host "  [DryRun] would snap $correl"
        continue
    }

    $snapPath = Get-SnapPath $snapFolder $correl
    Ensure-SnapDir $snapFolder | Out-Null

    # ── find Edge window ──────────────────────────────────────────────────────
    $edgeHwnd = Get-EdgeHwnd
    if (-not $edgeHwnd) {
        Write-Warning "  Edge not found — skip $correl"
        continue
    }

    # ── navigate / refresh URL ────────────────────────────────────────────────
    $cacheKey  = "${Mode}_${correl}"
    $cachedUrl = $urlCache[$cacheKey]

    $needNav = $RefreshUrls -or -not $cachedUrl

    if ($needNav) {
        Write-Host "  Searching Jenkins for $correl …"

        # activate Edge, open address bar, navigate to Jenkins search
        Activate-Window $edgeHwnd
        Start-Sleep -Milliseconds 300

        # Use Ctrl+F to search for the correl ID on the Jenkins page
        [System.Windows.Forms.SendKeys]::SendWait('^f')
        Start-Sleep -Milliseconds $cfg.Timing.ActionWaitMs
        [System.Windows.Forms.SendKeys]::SendWait($correl)
        Start-Sleep -Milliseconds $cfg.Timing.ResultWaitMs
    } else {
        Write-Host "  Using cached URL: $cachedUrl"
        Activate-Window $edgeHwnd
        Start-Sleep -Milliseconds 300

        # Navigate to cached URL
        [System.Windows.Forms.SendKeys]::SendWait('^l')
        Start-Sleep -Milliseconds 300
        [System.Windows.Forms.SendKeys]::SendWait($cachedUrl)
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait('{ENTER}')
        Start-Sleep -Milliseconds ($cfg.Timing.ResultWaitSec * 1000)
    }

    # ── resize window ─────────────────────────────────────────────────────────
    if (-not $NoResize) {
        Resize-Window $edgeHwnd $WindowWidth $WindowHeight
        Start-Sleep -Milliseconds 200
    }

    # ── take screenshot ───────────────────────────────────────────────────────
    $bmp = Capture-Window $edgeHwnd $CropPx
    if ($null -eq $bmp) {
        Write-Warning "  Screenshot failed — skip $correl"
        continue
    }
    $bmp.Save($snapPath, [System.Drawing.Imaging.ImageFormat]::Png)
    $bmp.Dispose()
    Write-Host "  Saved: $snapPath" -ForegroundColor Green

    # ── cache URL ─────────────────────────────────────────────────────────────
    if ($needNav) {
        # read current URL from address bar
        Activate-Window $edgeHwnd
        [System.Windows.Forms.SendKeys]::SendWait('^l')
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait('^c')
        Start-Sleep -Milliseconds 200
        [System.Windows.Forms.SendKeys]::SendWait('{ESC}')
        $currentUrl = [System.Windows.Forms.Clipboard]::GetText()
        if ($currentUrl -match '^https?://') {
            $urlCache[$cacheKey] = $currentUrl
            $dirty = $true
        }
    }

    # ── download data files (GiftRecv / GfixRecv only) ────────────────────────
    if ($Mode -in 'GiftRecv','GfixRecv') {
        $dataFolderName = if ($Mode -eq 'GiftRecv') { 'GIFT' } else { 'GFIX' }
        $dataDir = Join-Path $dataRoot $dataFolderName $correl
        if (-not (Test-Path $dataDir)) {
            New-Item $dataDir -ItemType Directory | Out-Null
        }

        # parse Jenkins file list from page text
        $pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'
        $pageText = & $pageTextScript
        if ($pageText) {
            $parseScript = Join-Path $scriptDir 'Parse-JenkinsList.ps1'
            $files = & $parseScript -Text $pageText
            Write-Host "  Found $($files.Count) file(s) on Jenkins page"
            foreach ($f in $files) {
                $destPath = Join-Path $dataDir $f.Name
                if (-not (Test-Path $destPath) -or $Force) {
                    try {
                        Invoke-WebRequest -Uri $f.Url -OutFile $destPath -UseBasicParsing
                        Write-Host "    DL: $($f.Name)" -ForegroundColor Gray
                    } catch {
                        Write-Warning "    DL failed: $($f.Name) — $_"
                    }
                }
            }
        }
    }

    # ── update mapping ────────────────────────────────────────────────────────
    $row.$snapField = '1'
}

# ── save mapping ──────────────────────────────────────────────────────────────
$rows | Export-Csv $mappingFile -NoTypeInformation -Encoding UTF8
Write-Host "`n[$Mode] Mapping saved." -ForegroundColor Green

if ($dirty) {
    Save-UrlCache $urlCache
    Write-Host "[$Mode] URL cache saved." -ForegroundColor Gray
}
