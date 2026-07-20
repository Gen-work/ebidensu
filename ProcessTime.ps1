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
#       SheetGiftRecv/SheetGfixRecv): candidate pictures are exported and
#       OCRed one at a time (Invoke-WinOcrFile in both en-US and the
#       configured secondary language, pooled -- the ja recognizer reads
#       the time-of-day the en-US one drops, and vice versa for other
#       fields), then VALIDATED BY CONTENT via ProcessTimeParse.ps1
#       (normalized datetime tokens + the correl id appearing in the
#       OCR text) rather than trusting picture position alone.
#  On tool-written workbooks each correl's HM screenshot sits immediately
#  after its Correl_ID_S label in column Replace.ColAnchor (default B) --
#  the layout EvidencePlan.ps1 wrote -- and the first section picture is
#  accepted directly. Hand-made workbooks (JDLW* office run, v2.12.2) can
#  put the picture ABOVE the label or crowd the label column so the
#  section collapses to one row; Resolve-ProcessTimeSide then widens the
#  search below and above the label, accepting a relaxed candidate only
#  when the correl id is actually seen in its OCR text.
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
    if ($Result.Matched) {
        return ("{0} -> {1} ({2})" -f (Format-ProcessTimeStamp $Result.StartTime), (Format-ProcessTimeStamp $Result.EndTime), $Result.Duration)
    }
    if ($null -ne $Result.StartTime) {
        # Partial OCR read: report WHICH time is there instead of a blanket miss.
        $durTxt = if (-not [string]::IsNullOrWhiteSpace($Result.Duration)) { (" (page duration {0})" -f $Result.Duration) } else { '' }
        return ("start {0}; end NOT read{1}" -f (Format-ProcessTimeStamp $Result.StartTime), $durTxt)
    }
    return 'not detected'
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

# OCRs one exported candidate PNG with the pooled recognizer languages and
# dumps the reconstructed rows to <png-stem>.ocr.txt next to it. Returns
# the pooled line array (plain array; callers wrap in @()).
function Read-ProcessTimeOcrLines {
    param([string]$Png, [string]$SecondaryLanguage, [string]$OutDir)
    $langs = @('en-US', $SecondaryLanguage) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $pooled = New-Object System.Collections.Generic.List[string]
    foreach ($lang in $langs) {
        try {
            Write-Host ("       [OCR] {0} lang={1}" -f (Split-Path $Png -Leaf), $lang) -ForegroundColor DarkGray
            $ocr = Invoke-WinOcrFile -Path $Png -LanguageTag $lang
            foreach ($ln in @(ConvertTo-SendRowLines $ocr.Lines)) { $pooled.Add($ln) }
        } catch {
            Write-Host ("       [WARN] OCR failed ({0}, {1}): {2}" -f (Split-Path $Png -Leaf), $lang, $_.Exception.Message) -ForegroundColor Yellow
        }
    }
    # Sidecar dump of the pooled OCR lines next to the PNG, so a run that
    # matched nothing can be diagnosed from the actual recognized text.
    try {
        $dumpPath = Join-Path $OutDir (([System.IO.Path]::GetFileNameWithoutExtension($Png)) + '.ocr.txt')
        [System.IO.File]::WriteAllText($dumpPath, (($pooled.ToArray()) -join [Environment]::NewLine), [System.Text.UTF8Encoding]::new($false))
        Write-Host ("       [OCR] wrote {0} line(s) -> {1}" -f $pooled.Count, $dumpPath) -ForegroundColor DarkGray
    } catch {
        Write-Host ("       [WARN] OCR dump write failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
    return $pooled.ToArray()
}

# Tiered start/end/duration resolution for one side (GIFT or GFIX) of one
# correl. Returns @{ Matched; Source; StartTime; EndTime; Duration; Note }.
#   Matched = $true only when BOTH start and end were extracted. A partial
#   OCR read (start only) leaves Matched=$false but fills StartTime and
#   explains the gap in Note, so the operator sees WHICH time is missing.
#   Source  = 'archived' | 'ocr' | 'ocr:<tier>' | 'ocr-partial[:<tier>]' | 'none'.
#
# OCR candidate pictures are tried in confidence order and validated by
# CONTENT (parsed time rows; the correl id appearing in the OCR text)
# instead of trusting position alone -- some hand-made workbooks put the
# HM picture ABOVE the correl label (JDLW* office-PC run, v2.12.2):
#   1. section      : first picture between the label and the next label
#                     (the layout Replace writes) -- a full time row here
#                     is accepted as-is, position is trusted.
#   2. below-label  : first 2 pictures below the section. Candidate 1 is
#                     the old v2.12.1 retry target (accepted with a note
#                     when its rows never show the correl id); candidate 2
#                     is only accepted when the correl id IS seen (it is
#                     usually the NEXT correl's picture).
#   3. above-label  : up to 3 pictures above the label, nearest first --
#                     accepted ONLY when the correl id is seen in the OCR
#                     text (position gives no evidence at all up there).
function Resolve-ProcessTimeSide {
    param($Workbook, [string]$SheetName, [string]$CorrelId, [string]$SnapTextPath,
          [string]$OutDir, [int]$AnchorCol, [string]$SecondaryLanguage, [double]$Scale,
          [string]$ExportBaseName = '')

    $result = @{ Matched = $false; Source = 'none'; StartTime = $null; EndTime = $null; Duration = ''; Note = '' }

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
                    Note = ''
                }
            }
        } catch {
            Write-Host ("       [WARN] archived text parse failed ({0}): {1}" -f $SnapTextPath, $_.Exception.Message) -ForegroundColor Yellow
        }
    }

    # Tier 2: OCR of the HM screenshot already inserted into the evidence workbook.
    $ws = Get-SheetByName $Workbook $SheetName
    if ($null -eq $ws) {
        $result.Note = ("sheet '{0}' not found in workbook" -f $SheetName)
        Write-Host ("       [MISS] {0}: {1}" -f $CorrelId, $result.Note) -ForegroundColor Yellow
        return $result
    }
    $labelCell = Find-ProcessTimeCorrelCell $ws $CorrelId $AnchorCol
    if ($null -eq $labelCell) {
        $result.Note = ("correl label not found in column {0} of sheet '{1}'" -f $AnchorCol, $SheetName)
        Write-Host ("       [MISS] {0}: {1}" -f $CorrelId, $result.Note) -ForegroundColor Yellow
        return $result
    }

    # Deterministic per-side base name (GIFT_/GFIX_) so the two sides of one
    # correl don't collide on the same PNG/dump names in the shared per-correl
    # export dir. Clear stale artifacts from a previous run so a MISS this run
    # can't be masked by last run's leftover PNG/dump.
    $base = if ([string]::IsNullOrWhiteSpace($ExportBaseName)) { $CorrelId } else { $ExportBaseName }
    if (Test-Path -LiteralPath $OutDir) {
        foreach ($pat in @(('{0}_*.png' -f $base), ('{0}_*.txt' -f $base))) {
            Get-ChildItem -LiteralPath $OutDir -Filter $pat -File -ErrorAction SilentlyContinue |
                ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
        }
    }

    $bounds = Get-ProcessTimeSectionBounds $ws $labelCell $AnchorCol
    $tiers = @()
    $belowMin = $bounds.Top
    if ($bounds.Bottom -ge 0) {
        $tiers += @{ Tag = 'section'; TopMin = $bounds.Top; TopMax = $bounds.Bottom; FromBottom = $false; Max = 1; RequireCorrel = $false }
        # below-label starts AFTER the section so a section picture that
        # already failed the content check is not re-exported.
        $belowMin = $bounds.Bottom
    }
    $tiers += @{ Tag = 'below-label'; TopMin = $belowMin; TopMax = -1;          FromBottom = $false; Max = 2; RequireCorrel = $false }
    $tiers += @{ Tag = 'above-label'; TopMin = -1;        TopMax = $bounds.Top; FromBottom = $true;  Max = 3; RequireCorrel = $true }

    $accepted = $null; $acceptedTag = ''
    $fallbackRow = $null; $fallbackRank = -1; $fallbackTag = ''
    $candTotal = 0
    $sectionHadPicture = $false
    $missNotes = New-Object System.Collections.Generic.List[string]

    foreach ($tier in $tiers) {
        if ($tier.Tag -ne 'section' -and $candTotal -gt 0) {
            Write-Host ("       [DIAG] no accepted picture yet for {0}; widening search: {1}" -f $CorrelId, $tier.Tag) -ForegroundColor Yellow
        }
        $tierBase = ('{0}_{1}' -f $base, ($tier.Tag -replace '-', ''))
        if ($tier.Tag -eq 'section') { $tierBase = $base }   # keep v2.12.1 file names for the trusted tier
        $pngs = @(Export-SheetPicturesToPng $Workbook $SheetName $OutDir $tierBase $tier.TopMin $tier.TopMax $Scale $tier.Max $tier.FromBottom |
            Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($tier.Tag -eq 'section' -and $pngs.Count -gt 0) { $sectionHadPicture = $true }

        $candIdx = 0
        foreach ($png in $pngs) {
            $candIdx++
            $candTotal++
            $lines = @(Read-ProcessTimeOcrLines -Png $png -SecondaryLanguage $SecondaryLanguage -OutDir $OutDir)
            $rows = @(ConvertFrom-ProcessTimeOcrLines -Lines $lines -CorrelId $CorrelId)
            # A below-label candidate is only position-trusted when it is the
            # FIRST picture below a section that had NO picture of its own
            # (the v2.12.1 retry target); once the section picture existed
            # (and failed validation), everything below the section belongs
            # to other correls and must show this correl's id to count.
            $strict = [bool]$tier.RequireCorrel -or
                ($tier.Tag -eq 'below-label' -and ($candIdx -ge 2 -or $sectionHadPicture))
            $sel = if ($strict) { Select-ProcessTimeRow -Rows $rows -RequireCorrelSeen } else { Select-ProcessTimeRow -Rows $rows }
            if ($null -eq $sel) {
                $note = if ($strict -and $rows.Count -gt 0) {
                    ("{0} time row(s) but correl id not in OCR text -- skipped (likely another correl's picture)" -f $rows.Count)
                } else {
                    Get-ProcessTimeOcrMissNote -Lines $lines
                }
                $missNotes.Add(("{0}: {1}" -f (Split-Path $png -Leaf), $note))
                Write-Host ("       [DIAG] {0}: {1}" -f (Split-Path $png -Leaf), $note) -ForegroundColor Yellow
                continue
            }
            $rank = Get-ProcessTimeRowRank $sel
            # Accept: full row + correl seen anywhere; a full row on position
            # alone only inside the trusted section tier.
            if ($rank -ge 3 -or ($tier.Tag -eq 'section' -and $rank -ge 2)) {
                $accepted = $sel; $acceptedTag = $tier.Tag; break
            }
            if ($rank -gt $fallbackRank) { $fallbackRow = $sel; $fallbackRank = $rank; $fallbackTag = $tier.Tag }
        }
        if ($null -ne $accepted) { break }
    }

    $row = $accepted; $tag = $acceptedTag
    $notes = New-Object System.Collections.Generic.List[string]
    if ($null -eq $row -and $null -ne $fallbackRow) {
        $row = $fallbackRow; $tag = $fallbackTag
        $seen = $row.PSObject.Properties['CorrelSeen'] -and [bool]$row.CorrelSeen
        if (-not $seen) { $notes.Add('correl id not seen in OCR text -- verify the picture') }
    }
    if ($null -eq $row) {
        if ($candTotal -eq 0) {
            $result.Note = ("no exportable picture found for the label on sheet '{0}'" -f $SheetName)
            Write-Host ("       [MISS] {0}: {1}" -f $CorrelId, $result.Note) -ForegroundColor Yellow
        } else {
            $result.Note = ("{0} candidate picture(s) OCRed, none yielded a usable time row -- {1}" -f $candTotal, ($missNotes -join '; '))
        }
        return $result
    }

    $result.StartTime = $row.StartTime
    $result.EndTime   = $row.EndTime
    $result.Matched   = ($null -ne $row.EndTime)

    # Duration: derived from start/end when both were read; cross-checked
    # against the page's own proc-time column, which also fills in when the
    # end time was unreadable (it is real on-page evidence, not invented).
    $pageDur = ''
    if ($row.PSObject.Properties['PageDuration']) { $pageDur = [string]$row.PageDuration }
    $derived = Get-ProcessDurationText $row.StartTime $row.EndTime
    if (-not [string]::IsNullOrWhiteSpace($derived)) {
        $result.Duration = $derived
        if (-not [string]::IsNullOrWhiteSpace($pageDur) -and $pageDur -ne $derived) {
            $notes.Add(("page duration {0} != derived {1} (kept derived)" -f $pageDur, $derived))
        }
    } elseif (-not [string]::IsNullOrWhiteSpace($pageDur)) {
        $result.Duration = $pageDur
        $notes.Add('duration read from the page column (end time not read)')
    }
    if (-not $result.Matched) { $notes.Add('end time not read from OCR') }

    $srcTag = if ($result.Matched) { 'ocr' } else { 'ocr-partial' }
    if ($tag -ne 'section' -and -not [string]::IsNullOrWhiteSpace($tag)) { $srcTag = ('{0}:{1}' -f $srcTag, $tag) }
    $result.Source = $srcTag
    $result.Note = ($notes -join '; ')
    return $result
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
# of clobbering earlier runs' extracted rows. A row whose
# (Excel_NAME, Correl_ID_S) pair is being rewritten this run REPLACES the
# old one (deleted first), so a -Force redo keeps ONE row per correl
# instead of stacking a fresh 43 rows on top of the previous run's 43.
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

        # Replace-in-place: delete existing rows for the correls being
        # rewritten (bottom-up so row indices stay valid). Keys are plain
        # string compares -- Excel_NAME/Correl_ID_S are always alphanumeric.
        if ($lastRow -gt 1) {
            $keys = @{}
            foreach ($r in @($Rows)) {
                $keys[('{0}|{1}' -f [string]$r.ExcelName, [string]$r.CorrelId)] = $true
            }
            $deleted = 0
            for ($rr = $lastRow; $rr -ge 2; $rr--) {
                $k = ('{0}|{1}' -f [string]$ws.Cells.Item($rr, 1).Value2, [string]$ws.Cells.Item($rr, 3).Value2)
                if ($keys.ContainsKey($k)) {
                    $ws.Rows.Item($rr).Delete() | Out-Null
                    $deleted++
                }
            }
            if ($deleted -gt 0) {
                Write-Host ("  [DIAG] replaced {0} existing row(s) in the ProcessTime workbook" -f $deleted) -ForegroundColor DarkGray
                try { $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 1).End($xlUp).Row } catch { $lastRow = 1 }
                if ($lastRow -lt 1) { $lastRow = 1 }
            }
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
$seenThisRun   = @{}   # 'Excel_NAME|Correl_ID_S' -> per-side flags of the first occurrence
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

                # The mapping can carry the same correl twice (observed:
                # JIDSCS4S / JIDSQS4S in a real run). Extract once per
                # (Excel_NAME, correl); later duplicates mirror the first
                # occurrence's flags instead of redoing export+OCR and
                # writing a second identical workbook row.
                $dupKey = ('{0}|{1}' -f $name, $correlId)
                if ($seenThisRun.ContainsKey($dupKey)) {
                    Write-Host '      [DIAG] duplicate mapping row for this correl; reusing this run''s result (check the mapping for unintended duplicates)' -ForegroundColor Yellow
                    $flags = $seenThisRun[$dupKey]
                    $row.GIFT_ProcessTime = $flags.Gift
                    $row.GFIX_ProcessTime = $flags.Gfix
                    $processedRows.Add($row)
                    continue
                }

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
                $seenThisRun[$dupKey] = @{ Gift = $row.GIFT_ProcessTime; Gfix = $row.GFIX_ProcessTime }

                Write-Host ("      GIFT: {0} [{1}]" -f (Format-ProcessTimeResult $giftResult), $giftResult.Source) -ForegroundColor DarkGray
                if (-not [string]::IsNullOrWhiteSpace([string]$giftResult.Note)) {
                    Write-Host ("        note: {0}" -f $giftResult.Note) -ForegroundColor Yellow
                }
                Write-Host ("      GFIX: {0} [{1}]" -f (Format-ProcessTimeResult $gfixResult), $gfixResult.Source) -ForegroundColor DarkGray
                if (-not [string]::IsNullOrWhiteSpace([string]$gfixResult.Note)) {
                    Write-Host ("        note: {0}" -f $gfixResult.Note) -ForegroundColor Yellow
                }

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
