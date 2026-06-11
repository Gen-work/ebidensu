# ============================================================
# SendVsGift.ps1 - compare-prep review for SEND data vs GIFT data
#
# Stage 1 MVP:
#   1. Gather exact metadata for every file under <WorkDir>\DATA\GIFT.
#   2. Save it to <WorkDir>\data\gift_metadata.csv.
#   3. Ensure mapping column SendVsGift exists.
#   4. For each pending workbook, print the matching GIFT metadata in the
#      console, open the evidence Excel for manual viewing, then Enter marks
#      SendVsGift=1, sets the cursor to A3, saves, and closes.
#
# Stage 2 handoff:
#   See docs/SendVsGift.md. OCR/image recognition is intentionally left as a
#   future enhancement because the MVP only needs exact file metadata and
#   operator confirmation.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string]$EvidenceDir = '',
    [string]$ExcelPrefix = '',
    [string]$CursorCell = 'A3',
    [string[]]$TargetIds = @(),
    [int]$SaveWaitMs = 1000,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$Maximize,
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path
        break
    }
}
if (-not $helpersPath) { throw 'ExcelHelpers.ps1 not found.' }
. $helpersPath
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Test-TargetRow($Row, [hashtable]$TargetSet) {
    if ($TargetSet.Count -eq 0) { return $true }
    return ($TargetSet.ContainsKey([string]$Row.Correl_ID_S) -or
            $TargetSet.ContainsKey([string]$Row.Correl_ID_M) -or
            $TargetSet.ContainsKey([string]$Row.JOB_NAME) -or
            $TargetSet.ContainsKey([string]$Row.Excel_NAME))
}

function Get-GiftDataDir([string]$Root) {
    $upperGift = Join-Path (Join-Path $Root 'DATA') 'GIFT'
    if (Test-Path -LiteralPath $upperGift) { return $upperGift }
    return (Join-Path (Join-Path $Root 'data') 'GIFT')
}

function ConvertTo-DisplaySize([long]$Bytes) {
    if ($Bytes -ge 1MB) { return ('{0:N2}MB ({1:N0} bytes)' -f ($Bytes / 1MB), $Bytes) }
    if ($Bytes -ge 1KB) { return ('{0:N2}KB ({1:N0} bytes)' -f ($Bytes / 1KB), $Bytes) }
    return ('{0:N0} bytes' -f $Bytes)
}

function Get-FirstToken([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $parts = @($Text.Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return '' }
    return $parts[0]
}

function Get-LastToken([string]$Text) {
    if ([string]::IsNullOrWhiteSpace($Text)) { return '' }
    $parts = @($Text.Trim() -split '\s+' | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
    if ($parts.Count -eq 0) { return '' }
    return $parts[$parts.Count - 1]
}

function Get-GiftFileMetadata([System.IO.FileInfo]$File) {
    $lineCount = 0
    $minLen = $null
    $maxLen = 0
    $firstRecord = ''
    $lastRecord = ''

    $reader = [System.IO.StreamReader]::new($File.FullName, [System.Text.Encoding]::Default, $true)
    try {
        while (($line = $reader.ReadLine()) -ne $null) {
            $lineCount++
            if ($lineCount -eq 1) { $firstRecord = $line }
            $lastRecord = $line
            $len = $line.Length
            if ($null -eq $minLen -or $len -lt $minLen) { $minLen = $len }
            if ($len -gt $maxLen) { $maxLen = $len }
        }
    } finally {
        $reader.Close()
    }

    if ($null -eq $minLen) { $minLen = 0 }
    [pscustomobject]@{
        FileName          = $File.Name
        FullName          = $File.FullName
        SizeBytes         = [long]$File.Length
        SizeDisplay       = ConvertTo-DisplaySize ([long]$File.Length)
        MaxRowNumber      = [int]$lineCount
        MinRecordLength   = [int]$minLen
        MaxRecordLength   = [int]$maxLen
        FirstRecordLength = [int]$firstRecord.Length
        LastRecordLength  = [int]$lastRecord.Length
        FirstRecordToken  = Get-FirstToken $firstRecord
        LastRecordToken   = Get-FirstToken $lastRecord
        FirstRecord       = $firstRecord
        LastRecord        = $lastRecord
        LastWriteTime     = $File.LastWriteTime.ToString('yyyy-MM-dd HH:mm:ss')
        MetadataVersion   = '1'
    }
}

function Write-GiftMetadata([string]$WorkRoot) {
    $giftDir = Get-GiftDataDir $WorkRoot
    $outDir = Join-Path $WorkRoot 'data'
    $outFile = Join-Path $outDir 'gift_metadata.csv'

    if (-not (Test-Path -LiteralPath $giftDir)) { throw "GIFT data folder not found: $giftDir" }
    if (-not (Test-Path -LiteralPath $outDir)) { New-Item -ItemType Directory -Path $outDir -Force | Out-Null }

    $files = @(Get-ChildItem -LiteralPath $giftDir -File -ErrorAction Stop | Sort-Object Name)
    $rows = @()
    foreach ($f in $files) { $rows += Get-GiftFileMetadata $f }
    $rows | Export-Csv -LiteralPath $outFile -Encoding UTF8 -NoTypeInformation -Force

    return @{ Path = $outFile; Rows = @($rows) }
}

function Find-GiftMetadataForRow($Row, [array]$MetadataRows) {
    $sid = [string]$Row.Correl_ID_S
    if ([string]::IsNullOrWhiteSpace($sid)) { return @() }
    $exact = @($MetadataRows | Where-Object { [string]$_.FileName -eq $sid })
    if ($exact.Count -gt 0) { return $exact }
    return @($MetadataRows | Where-Object { ([string]$_.FileName).StartsWith($sid) })
}

function Get-EvidencePath([string]$Dir, [string]$ExcelName) {
    $resolved = Find-WorkbookByExcelName -Dir $Dir -ExcelName $ExcelName
    if ($null -ne $resolved) { return $resolved }
    if ($ExcelName -match '\.xlsx$') { return (Join-Path $Dir $ExcelName) }
    return (Join-Path $Dir ($ExcelName + '.xlsx'))
}

function Set-WorkbookCursorAllSheets($Workbook, [string]$DefaultCell) {
    $arrow = [char]0x21D2
    $count = 0
    try { $count = [int]$Workbook.Worksheets.Count } catch { return }
    for ($i = $count; $i -ge 1; $i--) {
        $ws = $Workbook.Worksheets.Item($i)
        $cell = $DefaultCell
        try {
            if ([string]$ws.Name -like ('*' + $arrow + '*')) { $cell = 'A1' }
            $ws.Activate() | Out-Null
            $ws.Range($cell).Select() | Out-Null
        } catch {
            Write-Host ("  [WARN] failed to set cursor: sheet={0}, cell={1}" -f $ws.Name, $cell) -ForegroundColor Yellow
        }
    }
}

function Show-MetadataBlock($Row, [array]$Matches) {
    Write-Host ''
    Write-Host ("ID: {0}  Excel: {1}  Job: {2}" -f $Row.Correl_ID_S, $Row.Excel_NAME, $Row.JOB_NAME) -ForegroundColor Cyan
    if ($Matches.Count -eq 0) {
        Write-Host '  [WARN] no matching GIFT data file found.' -ForegroundColor Yellow
        return
    }
    foreach ($m in $Matches) {
        Write-Host ("  File        : {0}" -f $m.FullName)
        Write-Host ("  Size        : {0}" -f $m.SizeDisplay)
        Write-Host ("  Max row num : {0}" -f $m.MaxRowNumber)
        Write-Host ("  Length      : max={0}, min={1}, first={2}, last={3}" -f $m.MaxRecordLength, $m.MinRecordLength, $m.FirstRecordLength, $m.LastRecordLength)
        Write-Host ("  First row   : {0}" -f $m.FirstRecordToken)
        Write-Host ("  Last row    : {0}" -f $m.LastRecordToken)
    }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { throw "WorkDir not found: $WorkDir" }

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) { throw "mapping not found: $mappingPath" }

$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}

$metadataResult = Write-GiftMetadata $WorkDir
$metadataRows = @($metadataResult.Rows)
Write-Host ("[META] wrote {0} file metadata row(s): {1}" -f $metadataRows.Count, $metadataResult.Path) -ForegroundColor Green

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
if ($allRows.Count -eq 0) { throw "mapping has no rows: $mappingPath" }
Ensure-Column $allRows 'SendVsGift' '0'
$allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force

$selectedRows = @($allRows | Where-Object { Test-TargetRow $_ $targetSet })
$pendingRows = @($selectedRows | Where-Object { $Force.IsPresent -or [string]::IsNullOrWhiteSpace([string]$_.SendVsGift) -or [string]$_.SendVsGift -eq '0' })
if ($pendingRows.Count -eq 0) {
    Write-Host '[INFO] no pending SendVsGift rows.' -ForegroundColor Green
    return
}

Write-Host ("[INFO] pending SendVsGift rows: {0}" -f $pendingRows.Count) -ForegroundColor Cyan
if ($DryRun.IsPresent) { return }

$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    try { if ($Maximize.IsPresent) { $excel.WindowState = -4137 } } catch {}

    for ($idx = 0; $idx -lt $pendingRows.Count; $idx++) {
        $r = $pendingRows[$idx]
        $matches = @(Find-GiftMetadataForRow $r $metadataRows)
        Show-MetadataBlock $r $matches

        $prefix = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
        $fullStem = Get-ExcelFullStem -Prefix $prefix -Name ([string]$r.Excel_NAME)
        $file = Get-EvidencePath $EvidenceDir $fullStem
        if (-not (Test-Path -LiteralPath $file)) {
            Write-Host ("[{0}/{1}] MISSING workbook: {2}" -f ($idx + 1), $pendingRows.Count, $file) -ForegroundColor Yellow
            continue
        }

        $wb = $null
        try {
            Write-Host ("[{0}/{1}] OPEN: {2}" -f ($idx + 1), $pendingRows.Count, $file) -ForegroundColor Cyan
            $wb = $excel.Workbooks.Open($file, 0, $false)
            $ans = Read-Host 'Check the workbook and console metadata. Enter=mark SendVsGift=1/save/close, s=skip, q=quit'
            $ans = [string]$ans
            if ($ans.Trim().ToLower() -eq 'q') { throw '__SENDVSGIFT_QUIT__' }
            if ($ans.Trim().ToLower() -eq 's') { throw '__SENDVSGIFT_SKIP__' }

            Set-WorkbookCursorAllSheets $wb $CursorCell
            $wb.Save()
            Start-Sleep -Milliseconds $SaveWaitMs
            $wb.Close($false)
            $wb = $null

            $key = [string]$r.Correl_ID_S
            foreach ($row in $allRows) {
                if ([string]$row.Correl_ID_S -eq $key) { $row.SendVsGift = '1' }
            }
            $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
            Write-Host '  [DONE] SendVsGift=1' -ForegroundColor Green
        } catch {
            $msg = $_.Exception.Message
            if ($msg -eq '__SENDVSGIFT_QUIT__') {
                if ($wb) { try { $wb.Close($false) } catch {} }
                break
            }
            if ($msg -eq '__SENDVSGIFT_SKIP__') {
                if ($wb) { try { $wb.Close($false) } catch {} }
                Write-Host '  [SKIP]' -ForegroundColor Yellow
                continue
            }
            if ($wb) { try { $wb.Close($false) } catch {} }
            Write-Host ("  [ERROR] {0}" -f $msg) -ForegroundColor Red
        }
    }
} finally {
    if ($excel) {
        try { $excel.DisplayAlerts = $true } catch {}
        try { $excel.Quit() } catch {}
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) } catch {}
    }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}
