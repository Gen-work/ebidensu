#Requires -Version 5.1
# ============================================================
#  ProcessTime.ps1   (Phase: ProcessTime)   -- UTF-8, NO BOM, ASCII source.
#
#  For each pending mapping row (ProcessTime_Inserted bit 2 not yet set),
#  extracts the HM batch processing start time / end time (and derives the
#  duration) for the GIFT and GFIX sides, then writes one row per side per
#  correl into one or more ProcessTime evidence workbooks, classified by
#  configurable Excel_NAME tag (see "Output classification" below).
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
#    ProcessTime_Inserted : BITMASK (v2.15.0, matches the project's
#      isReplaced/isMarked/isReviewed convention) -- bit 1 (1) = this
#      correl's OCR result has been extracted and cached; bit 2 (2) = the
#      row has been written into an output workbook; 3 = both done. A
#      pre-v2.15.0 mapping's plain '1' (the old "written" flag) is migrated
#      to '3' once, on load (Get-ProcessTimeMigratedInsertedValue) -- see
#      Resolve-ProcessTimeRowPlan (ProcessTimeParse.ps1).
#
#  Output classification (v2.15.0): -OutputMode 'Split' (default) buckets
#  result rows by the first entry of -OutputTags found as a substring of
#  the row's Excel_NAME (e.g. 'JDL'/'JRV'/'JDS', fully configurable via
#  ProcessTime.OutputTags -- not hardcoded to just JDL/JRV), writing one
#  workbook per tag; a row matching no configured tag is routed to the
#  -UnclassifiedTag bucket (default 'Other') with a console WARN instead of
#  aborting the whole write. -OutputMode 'Single' ignores tags and writes
#  every result row into one workbook. -OutputDirectoryByTag lets a tag
#  route to its own destination directory instead of every tag sharing
#  -OutputDirectory.
#
#  End of run, every correl this run touched is checked against its cached
#  OCR result and any side that did not match prints in a
#  "needs manual check" summary, so the operator does not have to scroll
#  back through the OCR log to find which ids to verify by hand.
#
#  -Stage Ocr | Write | Both (default Both): lets the extraction pass and
#  the output-workbook write be run and re-run independently.
#    Ocr   -- resolve GIFT/GFIX for pending correls and cache the result to
#             a per-correl sidecar (WorkDir\snap\ProcessTime\<correl>\
#             result.json); never opens/writes the output workbook.
#    Write -- write the output workbook from already-cached sidecars only;
#             never opens an evidence workbook or runs OCR. A correl with
#             no cached sidecar is reported as a MISS and skipped (its
#             ProcessTime_Inserted flag is left untouched).
#    Both  -- OCR whatever is still needed, then write whatever is still
#             needed (reusing any sidecar already on disk instead of
#             redoing OCR for it).
#  Per-row re-run detection (non -Force) is now two INDEPENDENT signals,
#  not one shared one: the OCR stage is considered done for a correl when
#  its sidecar file exists OR ProcessTime_Inserted bit 1 is set (a
#  per-correl filesystem check plus the row's own bit, NOT the shared
#  output .xlsx many rows write into -- that file can't tell "already
#  extracted, just needs writing" apart from "never touched"); the write
#  stage is gated on ProcessTime_Inserted bit 2. -Force ignores both
#  signals for whichever stage(s) -Stage selects. A legacy row fully done
#  before this sidecar cache existed is migrated (plain '1' -> bitmask '3')
#  on load, so it is NOT treated as needing a fresh OCR redo -- see
#  Resolve-ProcessTimeRowPlan / Get-ProcessTimeMigratedInsertedValue
#  (ProcessTimeParse.ps1) for the exact rule and their unit tests.
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

    # Legacy destination option. Its directory is used when OutputDirectory
    # is blank; the file name itself is ignored because output is generated
    # per-tag (Split mode) or as one file (Single mode) from the ProcessTime
    # label (ProjectLabels.ps1 SheetProcessTime) as the stem.
    [string]$OutputPath = '',
    # Default destination directory. Blank -> WorkDir. Used for any tag with
    # no OutputDirectoryByTag override (Split mode), or for the one output
    # file (Single mode).
    [string]$OutputDirectory = '',
    [string]$OutputSheetName = '',
    # 'Split' (default): bucket result rows by -OutputTags into one workbook
    # per tag (falls back to -UnclassifiedTag when no tag matches). 'Single':
    # ignore tags and write every result row into one workbook.
    [string]$OutputMode = 'Split',
    # Ordered list of Excel_NAME substrings used to classify each result row
    # into its own output workbook (Split mode only). Not hardcoded to
    # JDL/JRV -- add more (e.g. 'JDS') via ProcessTime.OutputTags.
    [string[]]$OutputTags = @('JDL', 'JRV'),
    # Bucket name for a result row matching none of -OutputTags (Split mode).
    [string]$UnclassifiedTag = 'Other',
    # Optional Tag -> destination directory overrides (Split mode). A tag
    # absent here uses -OutputDirectory.
    [hashtable]$OutputDirectoryByTag = @{},

    # Which stage(s) to run: 'Ocr' (extract + cache sidecars only),
    # 'Write' (write the output workbook from cached sidecars only), or
    # 'Both' (default -- OCR whatever is still needed, then write whatever
    # is still needed). Case-insensitive; validated below.
    [string]$Stage = 'Both',

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

# Per-correl OCR result cache: WorkDir\snap\ProcessTime\<correl>\result.json,
# in the SAME per-correl folder Resolve-ProcessTimeSide already exports its
# diagnostic PNG/.ocr.txt candidates into. This is the per-row, filesystem-
# based "was this correl's OCR already extracted" signal
# (Resolve-ProcessTimeRowPlan, ProcessTimeParse.ps1) -- independent of
# whether the result has been WRITTEN into the shared output workbook yet
# (ProcessTime_Inserted), so a -Stage Write (or a Both) rerun can tell
# "already extracted, just needs writing" apart from "never touched".
function Get-ProcessTimeSidecarPath {
    param([string]$ExportRoot, [string]$CorrelId)
    return Join-Path (Join-Path $ExportRoot $CorrelId) 'result.json'
}

function Save-ProcessTimeSidecar {
    param([string]$Path, [pscustomobject]$Payload)
    $dir = Split-Path -Parent $Path
    if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
    ($Payload | ConvertTo-Json -Depth 6) | Set-Content -LiteralPath $Path -Encoding UTF8
}

# Returns $null when the sidecar is missing, unreadable, or not valid JSON --
# callers treat that exactly like "never OCR'd" (never throws; a half-written
# or corrupt cache file is never trusted over doing the real extraction).
function Read-ProcessTimeSidecar {
    param([string]$Path)
    if (-not (Test-Path -LiteralPath $Path)) { return $null }
    try {
        return (Get-Content -LiteralPath $Path -Raw -Encoding UTF8 | ConvertFrom-Json)
    } catch {
        Write-Host ("       [WARN] ProcessTime sidecar unreadable ({0}): {1}" -f $Path, $_.Exception.Message) -ForegroundColor Yellow
        return $null
    }
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
    $pooled  = New-Object System.Collections.Generic.List[string]
    $enLines = New-Object System.Collections.Generic.List[string]
    foreach ($lang in $langs) {
        try {
            Write-Host ("       [OCR] {0} lang={1}" -f (Split-Path $Png -Leaf), $lang) -ForegroundColor DarkGray
            $ocr = Invoke-WinOcrFile -Path $Png -LanguageTag $lang
            $rowLines = @(ConvertTo-SendRowLines $ocr.Lines)
            foreach ($ln in $rowLines) { $pooled.Add($ln) }
            # Keep the en-US rows separately: the Latin-digit recognizer reads
            # the 14-digit data-creation datestamp cleanly, so its dates are the
            # trusted source for correcting a ja date-digit misread.
            if ($lang -eq 'en-US') { foreach ($ln in $rowLines) { $enLines.Add($ln) } }
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
    return @{ Lines = $pooled.ToArray(); DateHints = (Get-ProcessTimeDateHints -Lines $enLines.ToArray()) }
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

    $result = @{ Matched = $false; Source = 'none'; StartTime = $null; EndTime = $null; Duration = ''; RecordCount = ''; Note = '' }

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
                    RecordCount = $(if ($best.PSObject.Properties['RecordCount']) { [string]$best.RecordCount } else { '' })
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

    # Deterministic per-side base name (GIFT_/GFIX_) so the two sides of one
    # correl don't collide on the same PNG/dump names in the shared per-correl
    # export dir. Clear stale artifacts from a previous run so a MISS this run
    # can't be masked by last run's leftover PNG/dump. Done regardless of the
    # label so the no-label fallback below also starts clean.
    $base = if ([string]::IsNullOrWhiteSpace($ExportBaseName)) { $CorrelId } else { $ExportBaseName }
    if (Test-Path -LiteralPath $OutDir) {
        foreach ($pat in @(('{0}_*.png' -f $base), ('{0}_*.txt' -f $base))) {
            Get-ChildItem -LiteralPath $OutDir -Filter $pat -File -ErrorAction SilentlyContinue |
                ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
        }
    }

    $tiers = @()
    if ($null -eq $labelCell) {
        # No per-correl text label on this side (some hand-made workbooks omit
        # the Correl_ID_S label in column B). Without a label there is no
        # section to anchor, so scan EVERY picture on the sheet and accept the
        # one that OCRs as a full HM row for THIS correl (fuzzy id). The HM
        # row's two-datetime structure is itself the classifier: the Excel
        # send-metadata strip, the MQ transfer table, and the Jenkins file
        # list never yield a full start+end HM row, so only the HM screenshot
        # qualifies. RequireCorrel keeps a multi-correl no-label sheet from
        # returning a neighbor's HM picture.
        Write-Host ("       [DIAG] {0}: no correl label in column {1} of sheet '{2}'; scanning every picture (no-label fallback)" -f $CorrelId, $AnchorCol, $SheetName) -ForegroundColor Yellow
        $tiers += @{ Tag = 'wholesheet'; TopMin = -1; TopMax = -1; FromBottom = $false; Max = 12; RequireCorrel = $true }
    } else {
        $bounds = Get-ProcessTimeSectionBounds $ws $labelCell $AnchorCol
        $belowMin = $bounds.Top
        if ($bounds.Bottom -ge 0) {
            $tiers += @{ Tag = 'section'; TopMin = $bounds.Top; TopMax = $bounds.Bottom; FromBottom = $false; Max = 1; RequireCorrel = $false }
            # below-label starts AFTER the section so a section picture that
            # already failed the content check is not re-exported.
            $belowMin = $bounds.Bottom
        }
        $tiers += @{ Tag = 'below-label'; TopMin = $belowMin; TopMax = -1;          FromBottom = $false; Max = 2; RequireCorrel = $false }
        $tiers += @{ Tag = 'above-label'; TopMin = -1;        TopMax = $bounds.Top; FromBottom = $true;  Max = 3; RequireCorrel = $true }
    }

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
            $read = Read-ProcessTimeOcrLines -Png $png -SecondaryLanguage $SecondaryLanguage -OutDir $OutDir
            $lines = @($read.Lines)
            $rows = @(ConvertFrom-ProcessTimeOcrLines -Lines $lines -CorrelId $CorrelId -StartDateHints $read.DateHints)
            # A below-label candidate is only position-trusted when it is the
            # FIRST picture below a section that had NO picture of its own
            # (the v2.12.1 retry target); once the section picture existed
            # (and failed validation), everything below the section belongs
            # to other correls and must show this correl's id to count.
            $strict = [bool]$tier.RequireCorrel -or
                ($tier.Tag -eq 'below-label' -and ($candIdx -ge 2 -or $sectionHadPicture))
            $sel = if ($strict) { Select-ProcessTimeRow -Rows $rows -RequireCorrelSeen } else { Select-ProcessTimeRow -Rows $rows }
            if ($null -eq $sel) {
                $dayRows = @($rows | Where-Object { $null -ne $_.StartTime -and $_.StartTime.TimeOfDay -ge ([timespan]'09:00:00') })
                $note = if ($rows.Count -gt 0 -and $dayRows.Count -eq 0) {
                    ("{0} time row(s), all before 09:00 -- skipped as history" -f $rows.Count)
                } elseif ($strict -and $rows.Count -gt 0) {
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
    if ($row.PSObject.Properties['RecordCount']) { $result.RecordCount = [string]$row.RecordCount }
    if ($row.PSObject.Properties['DateCorrected'] -and [bool]$row.DateCorrected) {
        $notes.Add('start date taken from en-US datestamp (ja date OCR-corrected)')
    }

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

# Writes the operator-facing, vertical layout requested for delivery:
# No. / GIFT-GFIX / correl / start / end / duration / count / job. Each
# correl in -Rows (the combined GIFT+GFIX shape ProcessTime.ps1 resolves,
# cached one-per-correl in the sidecar) becomes TWO output rows here, one
# per side. Incremental (non -Force) runs retain older records; a repeated
# (side, correl) pair REPLACES its prior row (deleted first), so a redo
# keeps one row per (side, correl) instead of stacking duplicates.
#
# Cell writes go through ExcelHelpers.ps1's Set-RangeValue2, not a bare
# Value2 assignment: column 1 (No.) writes an [int] while every other
# column writes a [string], and PS 5.1's COM property-set binder can cache
# a conversion rule from one Value2 call and misapply it to a later Value2
# call with a different value type on the same worksheet -- observed on a
# real run as "Unable to cast object of type 'System.Int32' to type
# 'System.String'" (same root cause as FillCheckSheet.ps1's v2.10.8 date
# write bug). Set-RangeValue2 retries via reflection on that failure.
function Write-ProcessTimeWorkbook {
    param($Excel, [string]$OutputPath, [string]$SheetName, [object[]]$Rows)

    $headers = @(
        'No.', 'GIFT/GFIX',
        ([string][char]0x76F8 + [char]0x95A2 + 'ID'),
        ([string][char]0x958B + [char]0x59CB + [char]0x65E5 + [char]0x6642),
        ([string][char]0x7D42 + [char]0x4E86 + [char]0x65E5 + [char]0x6642),
        ([string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593),
        ([string][char]0x51E6 + [char]0x7406 + [char]0x4EF6 + [char]0x6570),
        ([string][char]0x30B8 + [char]0x30E7 + [char]0x30D6)
    )

    $flatRows = New-Object System.Collections.Generic.List[object]
    $jobOrder = New-Object System.Collections.Generic.List[string]
    foreach ($r in @($Rows)) {
        $job = [string]$r.JobName
        if (-not $jobOrder.Contains($job)) { $jobOrder.Add($job) }
    }
    # Within each job, list all GIFT correls and then all GFIX correls
    # (rather than interleaving the two sides).
    foreach ($job in $jobOrder) {
        foreach ($side in @('GIFT', 'GFIX')) {
            foreach ($r in @($Rows | Where-Object { [string]$_.JobName -eq $job })) {
                $isGift = $side -eq 'GIFT'
                $flatRows.Add([pscustomobject]@{
                    Side = $side; CorrelId = [string]$r.CorrelId; JobName = [string]$r.JobName
                    Start = [string]$(if ($isGift) { $r.GiftStart } else { $r.GfixStart })
                    End = [string]$(if ($isGift) { $r.GiftEnd } else { $r.GfixEnd })
                    Duration = [string]$(if ($isGift) { $r.GiftDuration } else { $r.GfixDuration })
                    Count = [string]$(if ($isGift) { $r.GiftCount } else { $r.GfixCount })
                })
            }
        }
    }

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
                Set-RangeValue2 $cell $headers[$c] | Out-Null
                try { $cell.Font.Bold = $true } catch {}
            }
            $lastRow = 1
        }

        # Replace-in-place by (GIFT/GFIX, correl).
        if ($lastRow -gt 1) {
            $keys = @{}
            foreach ($r in $flatRows) {
                $keys[('{0}|{1}' -f $r.Side, $r.CorrelId)] = $true
            }
            $deleted = 0
            for ($rr = $lastRow; $rr -ge 2; $rr--) {
                $k = ('{0}|{1}' -f [string]$ws.Cells.Item($rr, 2).Value2, [string]$ws.Cells.Item($rr, 3).Value2)
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
        foreach ($r in $flatRows) {
            Set-RangeValue2 $ws.Cells.Item($row, 1) ($row - 1) | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 2) $r.Side | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 3) $r.CorrelId | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 4) $r.Start | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 5) $r.End | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 6) $r.Duration | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 7) $r.Count | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 8) $r.JobName | Out-Null
            $row++
        }
        # Renumber retained + appended records and make the range filterable.
        try {
            $finalRow = [int]$ws.Cells.Item($ws.Rows.Count, 2).End($xlUp).Row
            for ($rr = 2; $rr -le $finalRow; $rr++) { Set-RangeValue2 $ws.Cells.Item($rr, 1) ($rr - 1) | Out-Null }
            # Use A1 addresses rather than passing two Range COM proxies to
            # Worksheet.Range.  Some office-PC Excel/PowerShell combinations
            # reject the proxy overload with DISP_E_TYPEMISMATCH ("argument
            # type mismatch"), which must never prevent the workbook save.
            $tableRange = $ws.Range(('A1:H{0}' -f $finalRow))
            $tableRange.AutoFilter() | Out-Null
            $headerRange = $ws.Range('A1:H1')
            $headerRange.Interior.Color = 10053120
            $headerRange.Font.Color = 16777215
            $headerRange.Font.Bold = $true
            $tableRange.Borders.LineStyle = 1

            # The reference sheet uses Excel's Japanese default typeface at
            # 11 pt throughout, with compact 18-point rows.  Apply this to
            # the complete table after incremental rows have been merged so
            # retained and newly appended records are visually identical.
            $tableRange.Font.Name = 'Yu Gothic'
            $tableRange.Font.Size = 11
            $tableRange.Rows.RowHeight = 18
            $headerRange.HorizontalAlignment = -4108 # xlCenter
            $headerRange.VerticalAlignment = -4108   # xlCenter
            $ws.Range(('A2:A{0}' -f $finalRow)).HorizontalAlignment = -4108
            $tableRange.VerticalAlignment = -4108
            for ($rr = 2; $rr -le $finalRow; $rr++) {
                $side = [string]$ws.Cells.Item($rr, 2).Value2
                # Color is explicitly Double for the Excel COM Variant
                # binder; using the branch's boxed Int32 can also produce a
                # type mismatch on older Windows PowerShell/Excel versions.
                [double]$fill = if ($side -eq 'GIFT') { 15398626 } else { 16777215 }
                $ws.Range(('A{0}:H{0}' -f $rr)).Interior.Color = $fill
            }
        } catch {}
        try {
            # Fixed widths reproduce the template proportions and keep the
            # date/time fields from changing width with different data.
            $widths = @(9, 10, 14, 24, 24, 20, 14, 17)
            for ($c = 1; $c -le $widths.Count; $c++) {
                $ws.Columns.Item($c).ColumnWidth = $widths[$c - 1]
            }
        } catch {}

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

if (-not [string]::IsNullOrWhiteSpace($OutputDirectory)) {
    if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) { $OutputDirectory = Join-Path $WorkDir $OutputDirectory }
} elseif (-not [string]::IsNullOrWhiteSpace($OutputPath)) {
    if (-not [System.IO.Path]::IsPathRooted($OutputPath)) { $OutputPath = Join-Path $WorkDir $OutputPath }
    $OutputDirectory = Split-Path $OutputPath -Parent
} else {
    $OutputDirectory = $WorkDir
}
if (-not (Test-Path -LiteralPath $OutputDirectory)) {
    [void](New-Item -ItemType Directory -Path $OutputDirectory -Force)
}

$stageMap = @{ 'ocr' = 'Ocr'; 'write' = 'Write'; 'both' = 'Both' }
$stageKey = ([string]$Stage).Trim().ToLowerInvariant()
if ([string]::IsNullOrWhiteSpace($stageKey)) { $stageKey = 'both' }
if (-not $stageMap.ContainsKey($stageKey)) {
    Write-Host ("[ERROR] -Stage must be Ocr, Write, or Both (got '{0}')" -f $Stage) -ForegroundColor Red
    exit 1
}
$Stage = $stageMap[$stageKey]

$outputModeMap = @{ 'split' = 'Split'; 'single' = 'Single' }
$outputModeKey = ([string]$OutputMode).Trim().ToLowerInvariant()
if ([string]::IsNullOrWhiteSpace($outputModeKey)) { $outputModeKey = 'split' }
if (-not $outputModeMap.ContainsKey($outputModeKey)) {
    Write-Host ("[ERROR] -OutputMode must be Split or Single (got '{0}')" -f $OutputMode) -ForegroundColor Red
    exit 1
}
$OutputMode = $outputModeMap[$outputModeKey]
$OutputTags = @($OutputTags | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })
if ($null -eq $OutputDirectoryByTag) { $OutputDirectoryByTag = @{} }

$labels = Get-ProjectLabels
$outputLabel = $labels['SheetProcessTime']
if ([string]::IsNullOrWhiteSpace($OutputSheetName)) { $OutputSheetName = $outputLabel }
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

# One-way migration of a pre-v2.15.0 plain 0/1 ProcessTime_Inserted value
# ('1' = written) to the new bitmask shape (bit 1 = OCR'd, bit 2 = written;
# '3' = both, since a legacy write could only happen after OCR succeeded).
# See Get-ProcessTimeMigratedInsertedValue (ProcessTimeParse.ps1).
$migratedCount = 0
foreach ($row in $allRows) {
    $cur = Get-RowProp $row 'ProcessTime_Inserted'
    $migrated = Get-ProcessTimeMigratedInsertedValue $cur
    if ($migrated -ne $cur) { $row.ProcessTime_Inserted = $migrated; $migratedCount++ }
}
if ($migratedCount -gt 0) {
    Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
    Write-Host ("  [INFO] migrated {0} legacy ProcessTime_Inserted flag(s) (1 -> 3; bitmask: 1=OCR'd, 2=written)" -f $migratedCount) -ForegroundColor DarkGray
}

$exportRoot = Join-Path $WorkDir 'snap\ProcessTime'
$targets = @(ConvertTo-TargetIdList $TargetIds)

# Per-row plan: which stage(s) THIS run still needs, resolved from -Stage
# plus completion signals (Resolve-ProcessTimeRowPlan, ProcessTimeParse.ps1)
# -- a filesystem sidecar cache and the ProcessTime_Inserted bitmask's OCR
# (bit 1) and write (bit 2) bits. This replaces gating BOTH stages off the
# single shared output workbook, which cannot tell "OCR already extracted,
# just needs writing" apart from "never touched at all".
$plans = New-Object System.Collections.Generic.List[object]
foreach ($row in $allRows) {
    if (-not (Test-TargetRow $row $targets)) { continue }
    $correlId = [string]$row.Correl_ID_S
    if ([string]::IsNullOrWhiteSpace($correlId)) { continue }
    $sidecarPath = Get-ProcessTimeSidecarPath -ExportRoot $exportRoot -CorrelId $correlId
    $sidecarExists = Test-Path -LiteralPath $sidecarPath
    $insertedVal = Get-RowProp $row 'ProcessTime_Inserted'
    $ocrBitSet   = Test-BitDone $insertedVal 1
    $writeBitSet = Test-BitDone $insertedVal 2
    $plan = Resolve-ProcessTimeRowPlan -SidecarExists $sidecarExists -OcrDone $ocrBitSet -WriteDone $writeBitSet -Stage $Stage -Force $forceFlag
    if (-not $plan.Touch) { continue }
    $plans.Add([pscustomobject]@{
        Row = $row; CorrelId = $correlId; SidecarPath = $sidecarPath
        NeedsOcr = $plan.NeedsOcr; NeedsWrite = $plan.NeedsWrite
    })
}
$ocrPlans   = @($plans | Where-Object { $_.NeedsOcr })
$writePlans = @($plans | Where-Object { $_.NeedsWrite })

Write-Host ''
Write-Host '===== ProcessTime =====' -ForegroundColor Green
Write-Host ("  WorkDir      : {0}" -f $WorkDir)
Write-Host ("  EvidenceDir  : {0}" -f $EvidenceDir)
Write-Host ("  OutputMode   : {0}" -f $OutputMode)
if ($OutputMode -eq 'Single') {
    Write-Host ("  Output       : {0}" -f (Join-Path $OutputDirectory (Get-ProcessTimeOutputFileName -Label $outputLabel -Tag '' -Single $true)))
} else {
    Write-Host ("  OutputTags   : {0} (+ '{1}' fallback bucket)" -f ($OutputTags -join ', '), $UnclassifiedTag)
    Write-Host ("  OutputDir    : {0}" -f $OutputDirectory)
}
Write-Host ("  AnchorCol    : {0}" -f $AnchorCol)
Write-Host ("  OcrLanguage  : {0}" -f $(if ([string]::IsNullOrWhiteSpace($OcrLanguage)) { 'en-US' } else { 'en-US + ' + $OcrLanguage }))
Write-Host ("  Stage        : {0}" -f $Stage)
Write-Host ("  Pending      : {0} row(s) touched -- OCR {1}, Write {2}" -f $plans.Count, $ocrPlans.Count, $writePlans.Count)
Write-Host ("  Force        : {0}   DryRun : {1}" -f $forceFlag, $dryRunFlag)

Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'start' -Status 'info' `
    -Message ("stage={0} pending={1} ocr={2} write={3}" -f $Stage, $plans.Count, $ocrPlans.Count, $writePlans.Count)

if ($plans.Count -eq 0) {
    Write-Host ("[ProcessTime] No pending rows for -Stage {0}." -f $Stage) -ForegroundColor Green
    exit 0
}

if ($dryRunFlag) {
    Write-Host ''
    Write-Host '  [DRY RUN] archived-text-only preview (no Excel/OCR opened):' -ForegroundColor Yellow
    foreach ($p in $plans) {
        $correlId = $p.CorrelId
        $ocrTxt   = if ($p.NeedsOcr) { 'yes' } else { 'no' }
        $writeTxt = if ($p.NeedsWrite) { 'yes' } else { 'no' }
        Write-Host ("    {0}  (OCR:{1} Write:{2})" -f $correlId, $ocrTxt, $writeTxt) -ForegroundColor Cyan
        if ($p.NeedsOcr) {
            $giftTxt = Join-Path (Join-Path $WorkDir 'snap\GIFT_HM') ("{0}.txt" -f $correlId)
            $gfixTxt = Join-Path (Join-Path $WorkDir 'snap\GFIX_HM') ("{0}.txt" -f $correlId)
            Write-Host ("      GIFT: {0}" -f (Get-ArchivedProcessTimePreview $giftTxt $correlId)) -ForegroundColor DarkGray
            Write-Host ("      GFIX: {0}" -f (Get-ArchivedProcessTimePreview $gfixTxt $correlId)) -ForegroundColor DarkGray
        } else {
            Write-Host ("      (cached OCR result from a previous run -- {0})" -f $p.SidecarPath) -ForegroundColor DarkGray
        }
    }
    if ($OutputMode -eq 'Single') {
        Write-Host ("  would write -> {0}" -f (Join-Path $OutputDirectory (Get-ProcessTimeOutputFileName -Label $outputLabel -Tag '' -Single $true))) -ForegroundColor DarkGray
    } else {
        Write-Host ("  would write -> one workbook per tag ({0}) under {1} (+ '{2}' fallback)" -f ($OutputTags -join ', '), $OutputDirectory, $UnclassifiedTag) -ForegroundColor DarkGray
    }
    exit 0
}

# -- OCR-pending rows grouped by Excel_NAME (mapping order; one workbook
#    open each). A -Stage Write run never reaches here with anything in
#    $ocrPlans, so it never opens an evidence workbook at all. --
$namesOrdered = New-Object System.Collections.Generic.List[string]
$rowsByName   = @{}
foreach ($p in $ocrPlans) {
    $name = [string]$p.Row.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $rowsByName.ContainsKey($name)) {
        $rowsByName[$name] = New-Object System.Collections.Generic.List[object]
        $namesOrdered.Add($name)
    }
    $rowsByName[$name].Add($p)
}

$excel = $null
$sidecarCache   = @{}   # CorrelId -> sidecar payload produced/read this run
$ocrDoneThisRun = @{}   # CorrelId -> $true iff its sidecar was saved this run (dedupes duplicate mapping rows; key presence alone means "already OCR'd this run")
$mappingDirty   = $false

try {
    $excel = New-ExcelApp

    if ($ocrPlans.Count -gt 0) {
        Write-Host ''
        Write-Host ("----- OCR stage: {0} row(s) -----" -f $ocrPlans.Count) -ForegroundColor Green
        foreach ($name in $namesOrdered) {
            $groupPlans = $rowsByName[$name]
            $prefix = Resolve-ExcelPrefixWithDisk -Row $groupPlans[0].Row -DefaultPrefix $ExcelPrefix -ExcelName $name -EvidenceDir $EvidenceDir
            $fullStem = Get-ExcelFullStem -Prefix $prefix -Name $name
            $wbPath = Find-WorkbookByExcelName -Dir $EvidenceDir -ExcelName $fullStem -FullWidthFallback Reject

            Write-Host ''
            Write-Host ("----- {0} -----" -f $name) -ForegroundColor Cyan
            if ($null -eq $wbPath) {
                Write-Host ("  [MISS] no evidence workbook found: {0}" -f $fullStem) -ForegroundColor Yellow
                foreach ($p in $groupPlans) {
                    Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS $p.CorrelId `
                        -JobName ([string]$p.Row.JOB_NAME) -Action 'find-workbook' -Status 'fail' `
                        -Message ("workbook not found: {0}" -f $fullStem)
                }
                continue
            }
            Write-Host ("  {0}" -f (Split-Path $wbPath -Leaf)) -ForegroundColor White

            $wb = $null
            try {
                $wb = $excel.Workbooks.Open($wbPath, 0, $true)   # read-only
                foreach ($p in $groupPlans) {
                    $correlId = $p.CorrelId
                    $row = $p.Row
                    $jobName = [string]$row.JOB_NAME
                    Write-Host ("    {0}" -f $correlId) -ForegroundColor White

                    # The mapping can carry the same correl twice (duplicate
                    # rows observed in a real run). Extract once per correl;
                    # later duplicates mirror the first occurrence's flags
                    # instead of redoing export+OCR and re-saving the sidecar.
                    if ($ocrDoneThisRun.ContainsKey($correlId)) {
                        Write-Host '      [DIAG] duplicate mapping row for this correl; reusing this run''s OCR result (check the mapping for unintended duplicates)' -ForegroundColor Yellow
                        $cached = $sidecarCache[$correlId]
                        $row.GIFT_ProcessTime = if ([bool]$cached.GiftMatched) { '1' } else { '2' }
                        $row.GFIX_ProcessTime = if ([bool]$cached.GfixMatched) { '1' } else { '2' }
                        if ($ocrDoneThisRun[$correlId]) { Set-MappingBit -Row $row -Field 'ProcessTime_Inserted' -Bit 1 }
                        $mappingDirty = $true
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
                    $mappingDirty = $true

                    $giftCountTxt = if (-not [string]::IsNullOrWhiteSpace([string]$giftResult.RecordCount)) { (" count={0}" -f $giftResult.RecordCount) } else { '' }
                    $gfixCountTxt = if (-not [string]::IsNullOrWhiteSpace([string]$gfixResult.RecordCount)) { (" count={0}" -f $gfixResult.RecordCount) } else { '' }
                    Write-Host ("      GIFT: {0} [{1}]{2}" -f (Format-ProcessTimeResult $giftResult), $giftResult.Source, $giftCountTxt) -ForegroundColor DarkGray
                    if (-not [string]::IsNullOrWhiteSpace([string]$giftResult.Note)) {
                        Write-Host ("        note: {0}" -f $giftResult.Note) -ForegroundColor Yellow
                    }
                    Write-Host ("      GFIX: {0} [{1}]{2}" -f (Format-ProcessTimeResult $gfixResult), $gfixResult.Source, $gfixCountTxt) -ForegroundColor DarkGray
                    if (-not [string]::IsNullOrWhiteSpace([string]$gfixResult.Note)) {
                        Write-Host ("        note: {0}" -f $gfixResult.Note) -ForegroundColor Yellow
                    }

                    Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS $correlId -JobName $jobName -Action 'extract' `
                        -Status $(if ($giftResult.Matched -or $gfixResult.Matched) { 'ok' } else { 'fail' }) `
                        -Message ("GIFT={0}[{1}] GFIX={2}[{3}]" -f $giftResult.Duration, $giftResult.Source, $gfixResult.Duration, $gfixResult.Source)

                    $resultRow = [pscustomobject]@{
                        ExcelName    = $name
                        JobName      = $jobName
                        CorrelId     = $correlId
                        GiftStart    = (Format-ProcessTimeStamp $giftResult.StartTime)
                        GiftEnd      = (Format-ProcessTimeStamp $giftResult.EndTime)
                        GiftDuration = $giftResult.Duration
                        GiftCount    = $giftResult.RecordCount
                        GiftSource   = $giftResult.Source
                        GfixStart    = (Format-ProcessTimeStamp $gfixResult.StartTime)
                        GfixEnd      = (Format-ProcessTimeStamp $gfixResult.EndTime)
                        GfixDuration = $gfixResult.Duration
                        GfixCount    = $gfixResult.RecordCount
                        GfixSource   = $gfixResult.Source
                    }
                    $payload = [pscustomobject]@{
                        SchemaVersion = 1
                        SavedAt       = (Get-Date).ToString('yyyy-MM-dd HH:mm:ss')
                        GiftMatched   = [bool]$giftResult.Matched
                        GfixMatched   = [bool]$gfixResult.Matched
                        GiftNote      = [string]$giftResult.Note
                        GfixNote      = [string]$gfixResult.Note
                        Row           = $resultRow
                    }
                    $sidecarSaved = $false
                    try {
                        Save-ProcessTimeSidecar -Path $p.SidecarPath -Payload $payload
                        $sidecarSaved = $true
                    } catch {
                        Write-Host ("      [WARN] could not write ProcessTime sidecar ({0}): {1}" -f $p.SidecarPath, $_.Exception.Message) -ForegroundColor Yellow
                    }
                    if ($sidecarSaved) { Set-MappingBit -Row $row -Field 'ProcessTime_Inserted' -Bit 1 }
                    $sidecarCache[$correlId] = $payload
                    $ocrDoneThisRun[$correlId] = $sidecarSaved
                }
            } catch {
                Write-Host ("  [FAIL] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
                foreach ($p in $groupPlans) {
                    Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS $p.CorrelId `
                        -JobName ([string]$p.Row.JOB_NAME) -Action 'open-workbook' -Status 'fail' -Message $_.Exception.Message
                }
            } finally {
                if ($null -ne $wb) { try { $wb.Close($false) } catch {} }
            }
        }

        if ($mappingDirty) {
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            $mappingDirty = $false
        }
    }

    if ($Stage -eq 'Ocr') {
        Write-Host ''
        Write-Host ("[OK] OCR stage complete -- {0} correl(s) cached under {1}" -f $sidecarCache.Count, $exportRoot) -ForegroundColor Green
    } else {
        # Write stage: build the output rows purely from cached sidecars --
        # this run's own OCR pass first, else whatever is already on disk
        # from a previous -Stage Ocr run. No evidence workbook is opened here.
        $results = New-Object System.Collections.Generic.List[object]
        $writtenRowsByCorrel = @{}   # CorrelId -> list of mapping row(s) contributed this run (duplicate-correl-safe)
        foreach ($p in $writePlans) {
            $correlId = $p.CorrelId
            $payload = $sidecarCache[$correlId]
            if ($null -eq $payload) { $payload = Read-ProcessTimeSidecar -Path $p.SidecarPath }
            if ($null -eq $payload -or $null -eq $payload.Row) {
                Write-Host ("  [MISS] {0}: no cached ProcessTime OCR result -- run -Stage Ocr (or Both) first" -f $correlId) -ForegroundColor Yellow
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -CorrelIdS $correlId -JobName ([string]$p.Row.JOB_NAME) `
                    -Action 'write-row' -Status 'fail' -Message 'no cached OCR result; run -Stage Ocr (or Both) first'
                continue
            }
            $results.Add($payload.Row)
            if (-not $writtenRowsByCorrel.ContainsKey($correlId)) { $writtenRowsByCorrel[$correlId] = New-Object System.Collections.Generic.List[object] }
            $writtenRowsByCorrel[$correlId].Add($p.Row)
        }

        if ($results.Count -eq 0) {
            Write-Host ''
            Write-Host '[ProcessTime] nothing to write; no evidence workbook touched.' -ForegroundColor Yellow
        } else {
            # Group by output tag (Split mode) or write everything to one
            # workbook (Single mode). Unlike the old JDL/JRV-only classifier,
            # a row matching none of -OutputTags is routed to the
            # -UnclassifiedTag bucket instead of throwing -- one unexpected
            # Excel_NAME (e.g. a project tag not yet added to OutputTags)
            # must never abort every OTHER tag's already-resolved write.
            # Each tag is also written in its OWN try/catch, so one tag
            # failing to save (e.g. its output file is open elsewhere) does
            # not prevent the other tags from being written this run.
            $buckets = @{}
            $bucketOrder = New-Object System.Collections.Generic.List[string]
            if ($OutputMode -eq 'Single') {
                $bucketOrder.Add('')
                # Keep every bucket as the same concrete collection type.  In
                # Windows PowerShell 5.1, wrapping a generic List[object]
                # obtained through a hashtable index in @() can fail with
                # "Argument types do not match" instead of enumerating it.
                $buckets[''] = $results
            } else {
                foreach ($r in $results) {
                    $tag = Get-ProcessTimeOutputTag -ExcelName ([string]$r.ExcelName) -Tags $OutputTags -UnclassifiedTag $UnclassifiedTag
                    if (-not $buckets.ContainsKey($tag)) { $buckets[$tag] = New-Object System.Collections.Generic.List[object]; $bucketOrder.Add($tag) }
                    $buckets[$tag].Add($r)
                }
            }

            $written = New-Object System.Collections.Generic.List[string]
            $failedTags = New-Object System.Collections.Generic.List[string]
            $writtenCorrelCount = 0
            foreach ($tag in $bucketOrder) {
                $rowsForTag = ConvertTo-ProcessTimeBucketArray -Bucket $buckets[$tag]
                if ($rowsForTag.Count -eq 0) { continue }
                if ($OutputMode -ne 'Single' -and $tag -eq $UnclassifiedTag) {
                    $names = ($rowsForTag | ForEach-Object { [string]$_.ExcelName } | Select-Object -Unique) -join ', '
                    Write-Host ("  [WARN] {0} result row(s) matched no configured OutputTag ({1}) -- routed to the '{2}' bucket: {3}" -f `
                        $rowsForTag.Count, ($OutputTags -join ', '), $UnclassifiedTag, $names) -ForegroundColor Yellow
                }
                $dir = if ($OutputMode -eq 'Single') { $OutputDirectory } else { Resolve-ProcessTimeOutputDir -Tag $tag -DirByTag $OutputDirectoryByTag -DefaultDir $OutputDirectory }
                if (-not [System.IO.Path]::IsPathRooted($dir)) { $dir = Join-Path $WorkDir $dir }
                if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
                $fileName = Get-ProcessTimeOutputFileName -Label $outputLabel -Tag $tag -Single ($OutputMode -eq 'Single')
                $path = Join-Path $dir $fileName
                try {
                    Write-ProcessTimeWorkbook -Excel $excel -OutputPath $path -SheetName $OutputSheetName -Rows $rowsForTag
                    $written.Add($path)
                    $writtenCorrelCount += $rowsForTag.Count
                    # Only the rows that made it into a SUCCESSFULLY written
                    # workbook get their write bit set; a row whose tag's
                    # workbook failed to save stays pending for the next run.
                    # Every mapping row sharing a correl (duplicates included)
                    # is marked, not just the first.
                    foreach ($r in $rowsForTag) {
                        foreach ($mrow in @($writtenRowsByCorrel[$r.CorrelId])) {
                            Set-MappingBit -Row $mrow -Field 'ProcessTime_Inserted' -Bit 2
                        }
                    }
                } catch {
                    $failedTags.Add($tag)
                    Write-Host ("  [FAIL] could not write '{0}' output workbook ({1}): {2}" -f $tag, $path, $_.Exception.Message) -ForegroundColor Red
                    Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'write-workbook' -Status 'fail' `
                        -Message ("tag={0} path={1} {2}" -f $tag, $path, $_.Exception.Message)
                }
            }

            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-Host ''
            if ($written.Count -gt 0) {
                Write-Host ("[OK] wrote {0} correl(s), {1} output row(s) -> {2}" -f $writtenCorrelCount, ($writtenCorrelCount * 2), ($written -join ', ')) -ForegroundColor Green
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'write-workbook' -Status 'ok' `
                    -Message ("{0} correl(s), {1} output row(s) -> {2}" -f $writtenCorrelCount, ($writtenCorrelCount * 2), ($written -join ', '))
            }
            if ($failedTags.Count -gt 0) {
                Write-Host ("[ProcessTime] {0} tag(s) failed to write and stay pending for the next run: {1}" -f $failedTags.Count, ($failedTags -join ', ')) -ForegroundColor Yellow
            }
        }
    }
} finally {
    if ($excel) { Close-ExcelApp $excel }
}

# -- end-of-run manual-check summary ------------------------------------
# Every correl this run touched, checked against its cached OCR result
# (this run's own pass, else whatever is already cached on disk), so the
# operator gets a plain list of which ids to open and verify by hand
# instead of scrolling back through the whole OCR log.
$needsCheck = New-Object System.Collections.Generic.List[string]
foreach ($p in $plans) {
    $correlId = $p.CorrelId
    $payload = $sidecarCache[$correlId]
    if ($null -eq $payload) { $payload = Read-ProcessTimeSidecar -Path $p.SidecarPath }
    if ($null -eq $payload) {
        $needsCheck.Add(("{0}  -- no OCR result cached yet (run -Stage Ocr or Both)" -f $correlId))
        continue
    }
    $line = Get-ProcessTimeCheckSummaryLine -CorrelId $correlId -GiftMatched ([bool]$payload.GiftMatched) -GfixMatched ([bool]$payload.GfixMatched) `
        -GiftNote ([string]$payload.GiftNote) -GfixNote ([string]$payload.GfixNote)
    if (-not [string]::IsNullOrWhiteSpace($line)) { $needsCheck.Add($line) }
}
Write-Host ''
if ($needsCheck.Count -gt 0) {
    Write-Host ("===== ProcessTime: {0} correl(s) need manual check =====" -f $needsCheck.Count) -ForegroundColor Yellow
    foreach ($ln in $needsCheck) { Write-Host ("  {0}" -f $ln) -ForegroundColor Yellow }
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'ProcessTime' -Action 'manual-check-summary' -Status 'info' `
        -Message ("{0} correl(s): {1}" -f $needsCheck.Count, ($needsCheck -join ' | '))
} else {
    Write-Host '[ProcessTime] no correls need manual check this run.' -ForegroundColor Green
}

Write-Host ''
Write-Host '===== ProcessTime Done =====' -ForegroundColor Green
