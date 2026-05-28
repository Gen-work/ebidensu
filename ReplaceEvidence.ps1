# ============================================================
#  ReplaceEvidence.ps1
#
#  Phase: ReplaceGift / ReplaceGfix / ReplaceDf
#
#  Process per unique Excel_NAME (groups all Correl_ID_S sharing it):
#    1. Open work\evidence\<Excel_NAME>.xlsx
#    2. Find target sheet by -Mode:
#         Gift -> GIFT受信結果
#         Gfix -> GFIX受信結果
#         Df   -> GIFTデータvsGFIXデータ
#    3. Reset row 3 downward (delete shapes, clear values/format/highlight)
#    4. Insert images stacked at column B with blank rows between, picture
#       z-order = msoSendToBack so later Mark rectangles stay visible.
#    5. Tail per mode:
#         Gift -> label "GFIX Jenkins フォルダ受信ファイルなし" then
#                 GIFT_noGfixfile snaps stacked, label not repeated.
#         Gfix -> per-correl "GFIX受信log" + log paste stub
#                 (currently writes "<<TODO: GFIX 受信 log>>" placeholder).
#         Df   -> nothing extra.
#    6. Save workbook.
#    7. On all OK: isReplaced |= bit (1=Gift, 2=Gfix, 4=Df) on every row
#                  in the group, then save mapping.
#
#  Usage:
#    .\ReplaceEvidence.ps1 -Mode Gift
#    .\ReplaceEvidence.ps1 -Mode Gfix -TargetIds JIGPL48S
#    .\ReplaceEvidence.ps1 -Mode Df   -Force
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Gift','Gfix','Df')]
    [string]$Mode,

    [string]$WorkDir,
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
    [switch]$Force,

    [string]$CommonScript = '',
    [string]$ExcelHelpersScript = '',

    [int]$BlankRowsBetween = 1,

    # Labels (defaults filled below if empty). Override via VerifyConfig.psd1.
    [string]$GiftNoGfixLabel = '',
    [string]$GfixLogLabel = '',
    [string]$GfixLogTodoText = '<<TODO: GFIX 受信 log>>'
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

# Unblock UNC files
try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

# Resolve switch BEFORE dot-source
$forceFlag = [bool]$Force.IsPresent

# ============================================================
# Dot-source ExcelHelpers.ps1
# ============================================================
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'),
    (Join-Path (Split-Path $PSScriptRoot -Parent) 'VerifyTool\ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path
        break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
if (-not (Get-Command -Name 'New-ExcelApp' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed.' -ForegroundColor Red; exit 1
}

# ============================================================
# Filter targets
# ============================================================
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
            $targetSet.ContainsKey([string]$row.JOB_NAME) -or
            $targetSet.ContainsKey([string]$row.Excel_NAME))
}

# ============================================================
# Mode config (sheet names via [char] for encoding safety)
# ============================================================
$sheetGiftRecv = "GIFT" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GIFT受信結果
$sheetGfixRecv = "GFIX" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GFIX受信結果
$sheetDfDiff   = "GIFT" + [char]0x30C7 + [char]0x30FC + [char]0x30BF +              # GIFTデータ
                 "vs" +
                 "GFIX" + [char]0x30C7 + [char]0x30FC + [char]0x30BF                # GFIXデータ

# Default label fallbacks
$defaultGiftNoGfix = "GFIX Jenkins " +
                     [char]0x30D5 + [char]0x30A9 + [char]0x30EB + [char]0x30C0 +    # フォルダ
                     [char]0x53D7 + [char]0x4FE1 +                                   # 受信
                     [char]0x30D5 + [char]0x30A1 + [char]0x30A4 + [char]0x30EB +    # ファイル
                     [char]0x306A + [char]0x3057                                     # なし
$defaultGfixLog    = "GFIX" + [char]0x53D7 + [char]0x4FE1 + "log"                    # GFIX受信log

if ([string]::IsNullOrWhiteSpace($GiftNoGfixLabel)) { $GiftNoGfixLabel = $defaultGiftNoGfix }
if ([string]::IsNullOrWhiteSpace($GfixLogLabel))    { $GfixLogLabel    = $defaultGfixLog }

$modeCfg = switch ($Mode) {
    'Gift' { @{
        Sheet            = $sheetGiftRecv
        Bit              = 1
        UseExcelSnap     = $true
        PerCorrelFolders = @('GIFT_HM', 'GIFT_MQ', 'GIFT_Jenkins')
        TailKind         = 'NoGfix'
        TailLabel        = $GiftNoGfixLabel
        TailFolder       = 'GIFT_noGfixfile'
    } }
    'Gfix' { @{
        Sheet            = $sheetGfixRecv
        Bit              = 2
        UseExcelSnap     = $true
        PerCorrelFolders = @('GFIX_HM', 'GFIX_Jenkins')
        TailKind         = 'Log'
        TailLabel        = $GfixLogLabel
        TailFolder       = $null
    } }
    'Df'   { @{
        Sheet            = $sheetDfDiff
        Bit              = 4
        UseExcelSnap     = $false
        PerCorrelFolders = @('DF')
        TailKind         = 'None'
        TailLabel        = $null
        TailFolder       = $null
    } }
}

# ============================================================
# Header
# ============================================================
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'
$snapBase    = Join-Path $WorkDir 'snap'
$folderExcel = Join-Path $snapBase 'excel'

Write-Host ''
Write-Host ("===== ReplaceEvidence ({0}) =====" -f $Mode) -ForegroundColor Green
Write-Host ("  WorkDir   : {0}" -f $WorkDir)
Write-Host ("  Mapping   : {0}" -f $mappingPath)
Write-Host ("  Evidence  : {0}" -f $evDir)
Write-Host ("  Sheet     : {0}" -f $modeCfg.Sheet)
Write-Host ("  Bit       : {0}" -f $modeCfg.Bit)
Write-Host ("  Force     : {0}" -f $forceFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }
Write-Host ''

if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}
if (-not (Test-Path -LiteralPath $evDir)) {
    Write-Host "[ERROR] evidence dir missing: $evDir  (run Clone first)" -ForegroundColor Red; exit 1
}

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Ensure-Column $allRows 'isReplaced' '0'

$workRows = @($allRows | Where-Object { Test-TargetRow $_ })
if ($workRows.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

$groups = $workRows | Group-Object Excel_NAME | Sort-Object Name
Write-Host ("Groups (Excel_NAME): {0}" -f $groups.Count) -ForegroundColor Cyan

# ============================================================
# Helpers
# ============================================================
function Get-SnapPath([string]$folder, [string]$key) {
    return (Join-Path (Join-Path $snapBase $folder) ("{0}.png" -f $key))
}

function Get-GfixLogLines([string]$correlIdS, [string]$workDir, [array]$mappingRows) {
    $logDir = Join-Path $workDir 'log'
    if (-not (Test-Path -LiteralPath $logDir)) {
        Write-Host ("  [WARN] log dir not found: {0}" -f $logDir) -ForegroundColor Yellow
        return @(("<<WARN: log dir not found>>"))
    }
    $matchRow = $null
    foreach ($r in $mappingRows) {
        if ([string]$r.Correl_ID_S -eq $correlIdS) { $matchRow = $r; break }
    }
    if ($null -eq $matchRow) {
        Write-Host ("  [WARN] no mapping row for {0}" -f $correlIdS) -ForegroundColor Yellow
        return @(("<<WARN: no mapping row for {0}>>" -f $correlIdS))
    }
    $jobName = [string]$matchRow.JOB_NAME
    if ([string]::IsNullOrWhiteSpace($jobName)) {
        Write-Host ("  [WARN] JOB_NAME empty for {0}" -f $correlIdS) -ForegroundColor Yellow
        return @(("<<WARN: JOB_NAME empty for {0}>>" -f $correlIdS))
    }
    $logFile = Join-Path $logDir ("{0}.log" -f $jobName)
    if (-not (Test-Path -LiteralPath $logFile)) {
        Write-Host ("  [WARN] log file not found: {0}.log" -f $jobName) -ForegroundColor Yellow
        return @(("<<WARN: log file not found: {0}.log>>" -f $jobName))
    }
    $raw = [System.IO.File]::ReadAllBytes($logFile)
    $hasBom = ($raw.Length -ge 3 -and $raw[0] -eq 0xEF -and $raw[1] -eq 0xBB -and $raw[2] -eq 0xBF)
    $enc = if ($hasBom) { [System.Text.UTF8Encoding]::new($true) } else { [System.Text.UTF8Encoding]::new($false) }
    $content = $enc.GetString($raw)
    $lines = @($content -split "`r?`n")
    if ($lines.Count -gt 0 -and [string]::IsNullOrEmpty($lines[-1])) {
        $lines = $lines[0..($lines.Count - 2)]
    }
    if ($lines.Count -gt 100) {
        Write-Host ("  [WARN] {0}: log has {1} lines (>100), pasting all" -f $jobName, $lines.Count) -ForegroundColor Yellow
    }
    return $lines
}

# ============================================================
# Main loop
# ============================================================
$excel = New-ExcelApp
$cntDone = 0
$cntSkip = 0
$cntFail = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }

        Write-Host ''
        Write-Host ("=" * 72) -ForegroundColor White
        Write-Host ("  {0}   ({1} correl id)" -f $excelName, $g.Count) -ForegroundColor White
        Write-Host ("=" * 72) -ForegroundColor White

        # Workbook path
        $wbPath = Join-Path $evDir ("{0}.xlsx" -f $excelName)
        if (-not (Test-Path -LiteralPath $wbPath)) {
            Write-Host ("  [SKIP] workbook missing: {0}" -f $wbPath) -ForegroundColor Yellow
            $cntSkip++
            continue
        }

        # Already-done bit check
        $sampleBits = Get-BitValue $first 'isReplaced'
        if (-not $forceFlag -and (($sampleBits -band $modeCfg.Bit) -eq $modeCfg.Bit)) {
            Write-Host ("  [SKIP] bit {0} already set (isReplaced={1})" -f $modeCfg.Bit, $sampleBits) -ForegroundColor DarkGray
            $cntSkip++
            continue
        }

        $wb = $null
        try {
            $wb = Open-Workbook $excel $wbPath
        } catch {
            Write-Host ("  [FAIL] open: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $cntFail++
            continue
        }

        $allOk = $true
        try {
            Unhide-AllSheets $wb

            $ws = Get-SheetByName $wb $modeCfg.Sheet
            if ($null -eq $ws) {
                Write-Host ("  [FAIL] sheet not found: {0}" -f $modeCfg.Sheet) -ForegroundColor Red
                $allOk = $false
            } else {
                Reset-SheetBelowRow $ws 3 20

                $anchorRow = 3
                $correlIds = @($g.Group | ForEach-Object { [string]$_.Correl_ID_S } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
                $jobName   = [string]$first.JOB_NAME

                # ── ExcelSnap (Gift / Gfix) ──
                if ($modeCfg.UseExcelSnap) {
                    $excelPng = Join-Path $folderExcel ("{0}.png" -f $jobName)
                    if (Test-Path -LiteralPath $excelPng) {
                        $pic = Insert-PictureSendToBack $ws $anchorRow 2 $excelPng
                        Set-ShapeMetadata $pic 'excel' $jobName
                        $prevRow = $anchorRow
                        $anchorRow = Get-NextAnchorRow $ws $pic $BlankRowsBetween
                        Write-Host ("  [OK]   B{0}  excel\{1}.png" -f $prevRow, $jobName) -ForegroundColor Green
                    } else {
                        Write-Host ("  [WARN] excel\{0}.png missing" -f $jobName) -ForegroundColor Yellow
                        $allOk = $false
                    }
                }

                # ── Per-correl stacked snaps ──
                foreach ($folder in $modeCfg.PerCorrelFolders) {
                    foreach ($cid in $correlIds) {
                        $img = Get-SnapPath $folder $cid
                        if (Test-Path -LiteralPath $img) {
                            $pic = Insert-PictureSendToBack $ws $anchorRow 2 $img
                            Set-ShapeMetadata $pic $folder $cid
                            $prevRow = $anchorRow
                            $anchorRow = Get-NextAnchorRow $ws $pic $BlankRowsBetween
                            Write-Host ("  [OK]   B{0}  {1}\{2}.png" -f $prevRow, $folder, $cid) -ForegroundColor Green
                        } else {
                            Write-Host ("  [WARN] {0}\{1}.png missing" -f $folder, $cid) -ForegroundColor Yellow
                            $allOk = $false
                        }
                    }
                }

                # ── Tail per mode ──
                switch ($modeCfg.TailKind) {

                    'NoGfix' {
                        # Label once, then snaps (no repeated label)
                        if (-not [string]::IsNullOrWhiteSpace($modeCfg.TailLabel)) {
                            Write-PlainText $ws $anchorRow 2 $modeCfg.TailLabel
                            Write-Host ("  [OK]   B{0}  text: {1}" -f $anchorRow, $modeCfg.TailLabel) -ForegroundColor Green
                            $anchorRow = $anchorRow + 1
                        }
                        foreach ($cid in $correlIds) {
                            $img = Get-SnapPath $modeCfg.TailFolder $cid
                            if (Test-Path -LiteralPath $img) {
                                $pic = Insert-PictureSendToBack $ws $anchorRow 2 $img
                                Set-ShapeMetadata $pic $modeCfg.TailFolder $cid
                                $prevRow = $anchorRow
                                $anchorRow = Get-NextAnchorRow $ws $pic $BlankRowsBetween
                                Write-Host ("  [OK]   B{0}  {1}\{2}.png" -f $prevRow, $modeCfg.TailFolder, $cid) -ForegroundColor Green
                            } else {
                                # NoGfix snap is optional; some files do exist on GFIX side
                                Write-Host ("  [INFO] {0}\{1}.png absent (ok if file exists on GFIX)" -f $modeCfg.TailFolder, $cid) -ForegroundColor DarkGray
                            }
                        }
                    }

                    'Log' {
                        # Per-correl: label + log lines stub
                        foreach ($cid in $correlIds) {
                            if (-not [string]::IsNullOrWhiteSpace($modeCfg.TailLabel)) {
                                Write-PlainText $ws $anchorRow 2 ("{0}  ({1})" -f $modeCfg.TailLabel, $cid)
                                $anchorRow = $anchorRow + 1
                            }
                            $lines = @(Get-GfixLogLines $cid $WorkDir $allRows)
                            $anchorRow = Write-LogLines $ws $anchorRow 2 $lines
                            $anchorRow = $anchorRow + $BlankRowsBetween
                            Write-Host ("  [OK]   {0}: log {1} line(s)" -f $cid, $lines.Count) -ForegroundColor Green
                        }
                    }

                    'None' { }
                }
            }

            $wb.Save()
        } catch {
            Write-Host ("  [FAIL] processing: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $allOk = $false
        } finally {
            Close-Workbook $wb $false
        }

        if ($allOk) {
            # Set bit on all rows in this Excel_NAME group (in $allRows, not $workRows)
            $groupNames = @($g.Group | ForEach-Object { [string]$_.Correl_ID_M })
            foreach ($r in $allRows) {
                if ($groupNames -contains [string]$r.Correl_ID_M) {
                    Set-BitValue $r 'isReplaced' $modeCfg.Bit
                }
            }
            Write-Host ("  isReplaced |= {0} for {1} row(s)" -f $modeCfg.Bit, $g.Count) -ForegroundColor Green
            $cntDone++
        } else {
            Write-Host '  isReplaced NOT updated (allOk=false)' -ForegroundColor Yellow
            $cntFail++
        }
    }

    # Persist mapping if anything changed
    if ($cntDone -gt 0) {
        $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
        Write-Host ''
        Write-Host ("Mapping saved: {0}" -f $mappingPath) -ForegroundColor DarkGreen
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host ("===== ReplaceEvidence ({0}) Done =====" -f $Mode) -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
