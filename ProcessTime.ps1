#Requires -Version 5.1
# ============================================================
#  ProcessTime.ps1   (Phase: ProcessTime)   -- UTF-8, NO BOM, ASCII source.
#
#  For each pending mapping row (ProcessTime_Inserted = 0), extracts the
#  HM batch processing start time / end time (and derives the duration)
#  for the GIFT and GFIX sides, then writes one summary row per correl
#  into a standalone ProcessTime evidence workbook.
#
#  Source, two tiers per side (cheapest/most-accurate first):
#    1. archived Ctrl+A page text HmSnap.ps1 saved at snap time
#       (WorkDir\snap\GIFT_HM\<correl>.txt / GFIX_HM\<correl>.txt, only
#       present when SnapVerify.SaveText was on) -- re-parsed with
#       SnapVerify.ps1's ConvertFrom-HmPageText (exact, TAB-anchored).
#    2. OCR of the HM screenshot ALREADY INSERTED into the evidence
#       workbook (GIFT/GFIX jushin-kekka sheet -- see ProjectLabels.ps1
#       SheetGiftRecv/SheetGfixRecv): Export-CorrelPicture exports just
#       that correl's HM picture, Invoke-WinOcrFile (OcrWindows.ps1) reads
#       it in both en-US and the configured secondary language (pooled,
#       same lesson as GIFT_MQ's OCR tier -- one recognizer often reads
#       the ASCII date/time cleanly while the other garbles it), and
#       ConvertFrom-ProcessTimeOcrLines (ProcessTimeParse.ps1) anchors on
#       two datetime tokens per line instead of trusting column position.
#  Each correl's HM screenshot sits immediately after its Correl_ID_S
#  label in column Replace.ColAnchor (default B) on the recv sheet --
#  the same layout EvidencePlan.ps1's Build-Gift/GfixEvidencePlan wrote,
#  so the first picture found in that correl's section (label row to the
#  next label row) is always the HM screenshot even on the GFIX sheet
#  (where a GFIX log block follows in the same section).
#
#  Newest-by-StartTime wins when a page shows more than one run for a
#  correl, matching this project's established convention (see
#  SnapVerify.ps1 Test-HmAbend / Mark.ps1's GIFT_MQ row-position tiers).
#
#  Mapping columns (MappingStore.ps1):
#    GIFT_ProcessTime / GFIX_ProcessTime : informational per-side result
#      ('0' not yet attempted, '1' start/end extracted, '2' not found).
#    ProcessTime_Inserted : plain 0/1 completion flag (this phase's
#      Get-PendingRows field) -- '1' once the row has been written into
#      the ProcessTime evidence workbook, regardless of whether either
#      side was actually detected (a "not detected" row is still listed
#      so the operator can see it was checked). Re-run with -Force to redo.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [string]$EvidenceDir = '',
    [string]$ExcelPrefix = '',

    # Column (1-indexed) the correl-id label sits in on the recv sheets.
    # Matches Replace.ColAnchor (default 2 = column B); EvidencePlan.ps1's
    # New-TextOp writes correl labels there.
    [int]$AnchorCol = 2,

    # Destination for the generated evidence workbook. Blank -> WorkDir\ProcessTime_<Owner>.xlsx.
    [string]$OutputPath = '',
    [string]$OutputSheetName = '',

    # Secondary OCR language pooled alongside 'en-US' for the OCR tier.
    # Empty (default) means en-US only; set e.g. 'ja' to also pool the
    # Japanese recognizer's reading of the same picture.
    [string]$OcrLanguage = '',
    # Picture export upscale (matches EvidenceImageExport.ps1's own default).
    [double]$ExportScale = 3.0,

    [switch]$Force,
    [switch]$DryRun,
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag  = [bool]$Force.IsPresent
$dryRunFlag = [bool]$DryRun.IsPresent

# -- Dot-source ExcelHelpers.ps1 + shared libs (none have param()) --------
$helpersPath = $null
foreach ($c in @($ExcelHelpersScript, (Join-Path $PSScriptRoot 'ExcelHelpers.ps1'))) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).ProviderPath; break
    }
}
if (-not $helpersPath) {
    Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1
}
. $helpersPath
. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')
. (Join-Path $PSScriptRoot 'EvidenceImageExport.ps1')
. (Join-Path $PSScriptRoot 'OcrWindows.ps1')
. (Join-Path $PSScriptRoot 'SnapVerify.ps1')
. (Join-Path $PSScriptRoot 'SendMetadata.ps1')
. (Join-Path $PSScriptRoot 'ProcessTimeParse.ps1')

# -- small local helpers ---------------------------------------------

function Format-ProcessTimeStamp {
    param($DateTime)
    if ($null -eq $DateTime) { return '' }
    return $DateTime.ToString('yyyy/MM/dd HH:mm:ss')
}

function Format-ProcessTimeResult {
    param($Result)
    if (-not $Result.Matched) { return 'not detected' }
    return ("{0} -> {1} ({2})" -f (Format-ProcessTimeStamp $Result.StartTime), (Format-ProcessTimeStamp $Result.EndTime), $Result.Duration)
}

# Finds the Correl_ID_S label cell in $AnchorCol on the recv sheet
# (whole-cell match first, then substring, mirroring SendVsGift.ps1's
# Find-SendCorrelCell -- same technique, different column).
function Find-ProcessTimeCorrelCell {
    param($Worksheet, [string]$CorrelId, [int]$Col)
    if ([string]::IsNullOrWhiteSpace($CorrelId)) { return $null }
    $missing = [System.Reflection.Missing]::Value
    $rng = $Worksheet.Columns.Item($Col)
    foreach ($lookAt in @(1, 2)) {   # xlWhole, then xlPart
        $cell = $null
        try { $cell = $rng.Find($CorrelId, $missing, -4163, $lookAt) } catch { $cell = $null }
        if ($null -ne $cell) { return $cell }
    }
    return $null
}

# Vertical bounds of one correl section: from its label cell down to the
# next non-empty cell in the same column (the next correl's label), or
# unbounded when the label is the last one. Mirrors SendVsGift.ps1's
# Get-SendSectionBounds.
function Get-ProcessTimeSectionBounds {
    param($Worksheet, $LabelCell, [int]$Col)
    $top = 0.0
    try { $top = [double]$LabelCell.Top } catch {}
    $bottom = -1.0
    try {
        $r = [int]$LabelCell.Row
        $below = $Worksheet.Cells.Item($r + 1, $Col)
        if (-not [string]::IsNullOrWhiteSpace([string]$below.Text)) {
            $bottom = [double]$below.Top
        } else {
            $next = $below.End(-4121)   # xlDown -> next non-empty cell
            if ([int]$next.Row -lt [int]$Worksheet.Rows.Count -and
                -not [string]::IsNullOrWhiteSpace([string]$next.Text)) {
                $bottom = [double]$next.Top
            }
        }
    } catch {}
    return @{ Top = $top; Bottom = $bottom }
}

# Exports just this correl's HM screenshot (the first picture in its
# section -- HM is always inserted immediately after the correl label,
# see the file header note) to a PNG. Returns the path, or $null when the
# sheet/label/picture is not found.
function Export-CorrelPicture {
    param($Workbook, [string]$SheetName, [string]$CorrelId, [string]$OutDir, [int]$AnchorCol, [double]$Scale,
          [string]$BaseName = '')
    $ws = Get-SheetByName $Workbook $SheetName
    if ($null -eq $ws) { return $null }
    $labelCell = Find-ProcessTimeCorrelCell $ws $CorrelId $AnchorCol
    if ($null -eq $labelCell) { return $null }

    # Deterministic per-side base name (GIFT_/GFIX_) so the two sides of one
    # correl don't collide on the same <correl>_NN.png / .ocr.txt in the
    # shared per-correl export dir. Clear stale artifacts from a previous run
    # so a MISS this run can't be masked by last run's leftover PNG/dump.
    $base = if ([string]::IsNullOrWhiteSpace($BaseName)) { $CorrelId } else { $BaseName }
    if (Test-Path -LiteralPath $OutDir) {
        foreach ($pat in @(('{0}_*.png' -f $base), ('{0}_*.txt' -f $base))) {
            Get-ChildItem -LiteralPath $OutDir -Filter $pat -File -ErrorAction SilentlyContinue |
                ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
        }
    }

    $bounds = Get-ProcessTimeSectionBounds $ws $labelCell $AnchorCol
    # Only the first picture in the section is the HM screenshot, so cap the
    # export at 1 -- keeps the retry below from chart-exporting every picture
    # from the label to the sheet end on a busy sheet.
    $pngs = @(Export-SheetPicturesToPng $Workbook $SheetName $OutDir $base $bounds.Top $bounds.Bottom $Scale 1 |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($pngs.Count -eq 0 -and $bounds.Bottom -ge 0) {
        Write-Host ("       [DIAG] no picture in bounded section for {0}; retrying from label to sheet end" -f $CorrelId) -ForegroundColor Yellow
        $pngs = @(Export-SheetPicturesToPng $Workbook $SheetName $OutDir $base $bounds.Top -1 $Scale 1 |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    }
    if ($pngs.Count -eq 0) {
        Write-Host ("       [MISS] no exportable HM picture for {0} on sheet '{1}'" -f $CorrelId, $SheetName) -ForegroundColor Yellow
        return $null
    }
    return $pngs[0]
}

# Tiered start/end/duration resolution for one side (GIFT or GFIX) of one
# correl. Returns @{ Matched; Source='archived'|'ocr'|'none'; StartTime; EndTime; Duration }.
function Resolve-ProcessTimeSide {
    param($Workbook, [string]$SheetName, [string]$CorrelId, [string]$SnapTextPath,
          [string]$OutDir, [int]$AnchorCol, [string]$SecondaryLanguage, [double]$Scale,
          [string]$ExportBaseName = '')

    $none = @{ Matched = $false; Source = 'none'; StartTime = $null; EndTime = $null; Duration = '' }

    # Tier 1: archived Ctrl+A snap text (fast, exact).
    if (-not [string]::IsNullOrWhiteSpace($SnapTextPath) -and (Test-Path -LiteralPath $SnapTextPath)) {
        try {
            $text = Get-Content -LiteralPath $SnapTextPath -Raw -Encoding UTF8
            $matched = @(ConvertFrom-HmPageText $text | Where-Object { $_.CorrelId -eq $CorrelId })
            $best = Get-NewestProcessTimeRow -Rows $matched
            if ($null -ne $best) {
                return @{
                    Matched = $true; Source = 'archived'
                    StartTime = $best.StartTime; EndTime = $best.EndTime
                    Duration = (Get-ProcessDurationText $best.StartTime $best.EndTime)
                }
            }
        } catch {
            Write-Host ("       [WARN] archived text parse failed ({0}): {1}" -f $SnapTextPath, $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    # Tier 2: OCR of the HM screenshot already inserted into the evidence workbook.
    $png = Export-CorrelPicture -Workbook $Workbook -SheetName $SheetName -CorrelId $CorrelId `
        -OutDir $OutDir -AnchorCol $AnchorCol -Scale $Scale -BaseName $ExportBaseName
    if ($null -eq $png) { return $none }

    $langs = @('en-US', $SecondaryLanguage) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $pooled = New-Object System.Collections.Generic.List[string]
    foreach ($lang in $langs) {
        try {
            Write-Host ("       [OCR] {0} lang={1}" -f (Split-Path $png -Leaf), $lang) -ForegroundColor DarkGray
            $ocr = Invoke-WinOcrFile -Path $png -LanguageTag $lang
            foreach ($ln in @(ConvertTo-SendRowLines $ocr.Lines)) { $pooled.Add($ln) }
        } catch {
            Write-Host ("       [WARN] OCR failed ({0}, {1}): {2}" -f (Split-Path $png -Leaf), $lang, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
    # Sidecar dump of the pooled OCR lines next to the PNG, so a run that
    # matched nothing can be diagnosed from the actual recognized text.
    try {
        $dumpPath = Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension($png)) + '.ocr.txt')
        [System.IO.File]::WriteAllText($dumpPath, (($pooled.ToArray()) -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
        Write-Host ("       [OCR] wrote {0} line(s) -> {1}" -f $pooled.Count, $dumpPath) -ForegroundColor DarkGray
    } catch {
        Write-Host ("       [WARN] OCR dump write failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
    $ocrRows = @(ConvertFrom-ProcessTimeOcrLines -Lines $pooled.ToArray())
    $best = Get-NewestProcessTimeRow -Rows $ocrRows
    if ($null -eq $best) { return $none }
    return @{
        Matched = $true; Source = 'ocr'
        StartTime = $best.StartTime; EndTime = $best.EndTime
        Duration = (Get-ProcessDurationText $best.StartTime $best.EndTime)
    }
}

# Tier-1-only preview for -DryRun (no Excel/OCR opened).
function Get-ArchivedProcessTimePreview {
    param([string]$Path, [string]$CorrelId)
    if (-not (Test-Path -LiteralPath $Path)) { return '(no archived text; would need OCR on a real run)' }
    try {
        $text = Get-Content -LiteralPath $Path -Raw -Encoding UTF8
        $matched = @(ConvertFrom-HmPageText $text | Where-Object { $_.CorrelId -eq $CorrelId })
        $best = Get-NewestProcessTimeRow -Rows $matched
        if ($null -eq $best) { return '(no matching row in archived text; would need OCR on a real run)' }
        return ("{0} -> {1} ({2})" -f (Format-ProcessTimeStamp $best.StartTime), (Format-ProcessTimeStamp $best.EndTime), `
            (Get-ProcessDurationText $best.StartTime $best.EndTime))
    } catch {
        return ("(archived text parse error: {0})" -f $_.Exception.Message)
    }
}

# Appends $Rows to $OutputPath (creating it, with a header row, if it does
# not exist yet), so an incremental (non -Force) run accumulates instead
# of clobbering earlier runs' extracted rows.
function Write-ProcessTimeWorkbook {
    param($Excel, [string]$OutputPath, [string]$SheetName, [object[]]$Rows)

    $headers = @(
        'Excel_NAME', 'JOB_NAME', 'Correl_ID_S',
        'GIFT Start', 'GIFT End', 'GIFT Duration', 'GIFT Source',
        'GFIX Start', 'GFIX End', 'GFIX Duration', 'GFIX Source'
    )

    $isNew = -not (Test-Path -LiteralPath $OutputPath)
    $wb = if ($isNew) { $Excel.Workbooks.Add() } else { $Excel.Workbooks.Open($OutputPath) }
    try {
        $ws = Get-SheetByName $wb $SheetName
        if ($null -eq $ws) {
            $ws = $wb.Worksheets.Item(1)
            $ws.Name = $SheetName
        }

        $xlUp = -4162
        $lastRow = 0
        try { $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row } catch { $lastRow = 0 }
        if ($lastRow -eq 1 -and [string]::IsNullOrWhiteSpace([string]$ws.Cells.Item(1, 1).Value2)) { $lastRow = 0 }

        if ($lastRow -eq 0) {
            for ($c = 0; $c -lt $headers.Count; $c++) {
                $cell = $ws.Cells.Item(1, $c + 1)
                $cell.Value2 = $headers[$c]
                try { $cell.Font.Bold = $true } catch {}
            }
            $lastRow = 1
        }

        $row = $lastRow + 1
        foreach ($r in @($Rows)) {
            $ws.Cells.Item($row, 1).Value2  = [string]$r.ExcelName
            $ws.Cells.Item($row, 2).Value2  = [string]$r.JobName
            $ws.Cells.Item($row, 3).Value2  = [string]$r.CorrelId
            $ws.Cells.Item($row, 4).Value2  = [string]$r.GiftStart
            $ws.Cells.Item($row, 5).Value2  = [string]$r.GiftEnd
            $ws.Cells.Item($row, 6).Value2  = [string]$r.GiftDuration
            $ws.Cells.Item($row, 7).Value2  = [string]$r.GiftSource
            $ws.Cells.Item($row, 8).Value2  = [string]$r.GfixStart
            $ws.Cells.Item($row, 9).Value2  = [string]$r.GfixEnd
            $ws.Cells.Item($row, 10).Value2 = [string]$r.GfixDuration
            $ws.Cells.Item($row, 11).Value2 = [string]$r.GfixSource
            $row++
        }
        try { $ws.Columns.AutoFit() | Out-Null } catch {}

        if ($isNew) { $wb.SaveAs($OutputPath, 51) } else { $wb.Save() }   # 51 = xlOpenXMLWorkbook (.xlsx)
    } finally {
        try { $wb.Close($false) } catch {}
    }
}

# -- validate + resolve paths ------------------------------------------

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

if ([string]::IsNullOrWhiteSpace($OutputPath)) { $OutputPath = Join-Path $WorkDir ("ProcessTime_{0}.xlsx" -f $Owner) }
elseif (-not [System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath = Join-Path $WorkDir $OutputPath }

$labels = Get-ProjectLabels
if ([string]::IsNullOrWhiteSpace($OutputSheetName)) { $OutputSheetName = $labels['SheetProcessTime'] }
$sheetGiftRecv = $labels['SheetGiftRecv']
$sheetGfixRecv = $labels['SheetGfixRecv']

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

$allRows = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red; exit 1
}
Ensure-MappingColumns -Rows $allRows | Out-Null

$targets = @(ConvertTo-TargetIdList $TargetIds)
$pending = @(Get-PendingRows -Rows $allRows -Field 'ProcessTime_Inserted' -Force $forceFlag -Targets $targets)

Write-Host ''
Write-Host '===== ProcessTime =====' -ForegroundColor Green
Write-Host ("  WorkDir      : {0}" -f $WorkDir)
Write-Host ("  EvidenceDir  : {0}" -f $EvidenceDir)
Write-Host ("  OutputPath   : {0}" -f $OutputPath)
Write-Host ("  AnchorCol    : {0}" -f $AnchorCol)
Write-Host ("  OcrLanguage  : {0}" -f $(if ([string]::IsNullOrWhiteSpace($OcrLanguage)) { 'en-US' } else { 'en-US + ' + $OcrLanguage }))
Write-Host ("  Pending      : {0}" -f $pending.Count)
Write-Host ("  Force        : {0}   DryRun : {1}" -f $forceFlag, $dryRunFlag)

Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'start' -Status 'info' `
    -Message ("pending={0}" -f $pending.Count)

if ($pending.Count -eq 0) {
    Write-Host '[ProcessTime] No pending rows.' -ForegroundColor Green
    exit 0
}

if ($dryRunFlag) {
    Write-Host ''
    Write-Host '  [DRY RUN] archived-text-only preview (no Excel/OCR opened):' -ForegroundColor Yellow
    foreach ($row in $pending) {
        $correlId = [string]$row.Correl_ID_S
        if ([string]::IsNullOrWhiteSpace($correlId)) { continue }
        $giftTxt = Join-Path (Join-Path $WorkDir 'snap\GIFT_HM') ("{0}.txt" -f $correlId)
        $gfixTxt = Join-Path (Join-Path $WorkDir 'snap\GFIX_HM') ("{0}.txt" -f $correlId)
        Write-Host ("    {0}" -f $correlId) -ForegroundColor Cyan
        Write-Host ("      GIFT: {0}" -f (Get-ArchivedProcessTimePreview $giftTxt $correlId)) -ForegroundColor DarkGray
        Write-Host ("      GFIX: {0}" -f (Get-ArchivedProcessTimePreview $gfixTxt $correlId)) -ForegroundColor DarkGray
    }
    Write-Host ("  would write -> {0}" -f $OutputPath) -ForegroundColor DarkGray
    exit 0
}

# -- group pending rows by Excel_NAME (mapping order; one workbook open each) --
$namesOrdered = New-Object System.Collections.Generic.List[string]
$rowsByName   = @{}
foreach ($row in $pending) {
    $name = [string]$row.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $rowsByName.ContainsKey($name)) {
        $rowsByName[$name] = New-Object System.Collections.Generic.List[object]
        $namesOrdered.Add($name)
    }
    $rowsByName[$name].Add($row)
}

$excel = $null
$results       = New-Object System.Collections.Generic.List[object]
$processedRows = New-Object System.Collections.Generic.List[object]
$exportRoot    = Join-Path $WorkDir 'snap\ProcessTime'

try {
    $excel = New-ExcelApp

    foreach ($name in $namesOrdered) {
        $groupRows = $rowsByName[$name]
        $prefix = Resolve-ExcelPrefixWithDisk -Row $groupRows[0] -DefaultPrefix $ExcelPrefix -ExcelName $name -EvidenceDir $EvidenceDir
        $fullStem = Get-ExcelFullStem -Prefix $prefix -Name $name
        $wbPath = Find-WorkbookByExcelName -Dir $EvidenceDir -ExcelName $fullStem -FullWidthFallback Reject

        Write-Host ''
        Write-Host ("----- {0} -----" -f $name) -ForegroundColor Cyan
        if ($null -eq $wbPath) {
            Write-Host ("  [MISS] no evidence workbook found: {0}" -f $fullStem) -ForegroundColor Yellow
            foreach ($row in $groupRows) {
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS ([string]$row.Correl_ID_S) `
                    -JobName ([string]$row.JOB_NAME) -Action 'find-workbook' -Status 'fail' `
                    -Message ("workbook not found: {0}" -f $fullStem)
            }
            continue
        }
        Write-Host ("  {0}" -f (Split-Path $wbPath -Leaf)) -ForegroundColor White

        $wb = $null
        try {
            $wb = $excel.Workbooks.Open($wbPath, 0, $true)   # read-only
            foreach ($row in $groupRows) {
                $correlId = [string]$row.Correl_ID_S
                if ([string]::IsNullOrWhiteSpace($correlId)) { continue }
                $jobName = [string]$row.JOB_NAME
                Write-Host ("    {0}" -f $correlId) -ForegroundColor White

                $giftTxt   = Join-Path (Join-Path $WorkDir 'snap\GIFT_HM') ("{0}.txt" -f $correlId)
                $gfixTxt   = Join-Path (Join-Path $WorkDir 'snap\GFIX_HM') ("{0}.txt" -f $correlId)
                $exportDir = Join-Path $exportRoot $correlId

                $giftResult = Resolve-ProcessTimeSide -Workbook $wb -SheetName $sheetGiftRecv -CorrelId $correlId `
                    -SnapTextPath $giftTxt -OutDir $exportDir -AnchorCol $AnchorCol -SecondaryLanguage $OcrLanguage -Scale $ExportScale `
                    -ExportBaseName ("GIFT_{0}" -f $correlId)
                $gfixResult = Resolve-ProcessTimeSide -Workbook $wb -SheetName $sheetGfixRecv -CorrelId $correlId `
                    -SnapTextPath $gfixTxt -OutDir $exportDir -AnchorCol $AnchorCol -SecondaryLanguage $OcrLanguage -Scale $ExportScale `
                    -ExportBaseName ("GFIX_{0}" -f $correlId)

                $row.GIFT_ProcessTime = if ($giftResult.Matched) { '1' } else { '2' }
                $row.GFIX_ProcessTime = if ($gfixResult.Matched) { '1' } else { '2' }

                Write-Host ("      GIFT: {0} [{1}]" -f (Format-ProcessTimeResult $giftResult), $giftResult.Source) -ForegroundColor DarkGray
                Write-Host ("      GFIX: {0} [{1}]" -f (Format-ProcessTimeResult $gfixResult), $gfixResult.Source) -ForegroundColor DarkGray

                Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS $correlId -JobName $jobName -Action 'extract' `
                    -Status $(if ($giftResult.Matched -or $gfixResult.Matched) { 'ok' } else { 'fail' }) `
                    -Message ("GIFT={0}[{1}] GFIX={2}[{3}]" -f $giftResult.Duration, $giftResult.Source, $gfixResult.Duration, $gfixResult.Source)

                $results.Add([pscustomobject]@{
                    ExcelName    = $name
                    JobName      = $jobName
                    CorrelId     = $correlId
                    GiftStart    = (Format-ProcessTimeStamp $giftResult.StartTime)
                    GiftEnd      = (Format-ProcessTimeStamp $giftResult.EndTime)
                    GiftDuration = $giftResult.Duration
                    GiftSource   = $giftResult.Source
                    GfixStart    = (Format-ProcessTimeStamp $gfixResult.StartTime)
                    GfixEnd      = (Format-ProcessTimeStamp $gfixResult.EndTime)
                    GfixDuration = $gfixResult.Duration
                    GfixSource   = $gfixResult.Source
                })
                $processedRows.Add($row)
            }
        } catch {
            Write-Host ("  [FAIL] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            foreach ($row in $groupRows) {
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS ([string]$row.Correl_ID_S) `
                    -JobName ([string]$row.JOB_NAME) -Action 'open-workbook' -Status 'fail' -Message $_.Exception.Message
            }
        } finally {
            if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
        }
    }

    if ($results.Count -eq 0) {
        Write-Host ''
        Write-Host '[ProcessTime] nothing extracted; no evidence workbook written.' -ForegroundColor Yellow
    } else {
        try {
            Write-ProcessTimeWorkbook -Excel $excel -OutputPath $OutputPath -SheetName $OutputSheetName -Rows $results.ToArray()
            foreach ($row in $processedRows) { $row.ProcessTime_Inserted = '1' }
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-Host ''
            Write-Host ("[OK] wrote {0} row(s) -> {1}" -f $results.Count, $OutputPath) -ForegroundColor Green
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'write-workbook' -Status 'ok' `
                -Message ("{0} row(s) -> {1}" -f $results.Count, $OutputPath)
        } catch {
            Write-Host ("[FAIL] could not write process-time workbook: {0}" -f $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'write-workbook' -Status 'fail' -Message $_.Exception.Message
            # Still persist the per-side detection flags even though ProcessTime_Inserted
            # stays 0 -- cheaper to keep this run's OCR verdicts than force a full re-OCR.
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
        }
    }
} finally {
    if ($excel) { Close-ExcelApp $excel }
}

Write-Host ''
Write-Host '===== ProcessTime Done =====' -ForegroundColor Green
