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
#  Source, three tiers per side (cheapest/most-accurate first):
#    1. archived Ctrl+A page text HmSnap.ps1 saved at snap time
#       (WorkDir\snap\GIFT_HM\<correl>.txt / GFIX_HM\<correl>.txt, only
#       present when SnapVerify.SaveText was on) -- re-parsed with
#       SnapVerify.ps1's ConvertFrom-HmPageText (exact, TAB-anchored).
#    1.5 (v2.15.2) OCR of the snap-time HM screenshot itself
#       (WorkDir\snap\GIFT_HM\<correl>.png / GFIX_HM\<correl>.png) -- the
#       per-correl-named original capture, a cleaner OCR target than the
#       evidence-workbook copy and trusted like the section tier.
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
#  workbook per tag; when no configured tag matches, -AutoDeriveTag
#  (default $true, v2.15.2) derives the tag from the Excel_NAME's own
#  '?XXX????' shape (chars 2-4: CJODWDEJ -> JOD) so an unlisted project
#  family still routes to its own workbook, and only a non-conforming name
#  is routed to the -UnclassifiedTag bucket (default 'Other') with a
#  console WARN instead of aborting the whole write. -OutputMode 'Single' ignores tags and writes
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
    # When no configured -OutputTags entry matches, derive the tag from the
    # Excel_NAME's own '?XXX????' shape (chars 2-4, e.g. CJODWDEJ -> JOD)
    # instead of routing the row to -UnclassifiedTag. On by default so a
    # project family missing from OutputTags still gets its own workbook.
    [bool]$AutoDeriveTag = $true,
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
    # OCR image preprocessing (System.Drawing): upscale + grayscale +
    # contrast stretch applied to every picture BEFORE it is handed to the
    # Windows OCR engine, so thin digit strokes ('9' vs '3') get crisp
    # pixel boundaries instead of the recognizer eating the raw
    # low-resolution pixels. Falls back to the original image on any
    # preprocessing failure -- never blocks OCR.
    [bool]$OcrPreprocess = $true,
    # Upscale factor for preprocessing (capped so the result stays under
    # the WinRT OCR engine's MaxImageDimension; an image already at/over
    # the cap is only grayscaled/contrast-stretched, never downscaled).
    [double]$OcrPreprocessScale = 2.0,
    # Linear contrast factor applied around mid-gray (1.0 = off).
    [double]$OcrPreprocessContrast = 1.3,
    # RESERVED (v2.16.0): OCR preprocessing binarization. Not wired into any
    # logic yet -- accepted, threaded from config, and documented so a later
    # stage can turn a global/adaptive threshold on. Toggling it today has no
    # effect on the image pipeline.
    [bool]$OcrPreprocessBinarize = $false,
    [int]$OcrPreprocessThreshold = 128,
    # Picture export upscale (matches EvidenceImageExport.ps1's own default).
    [double]$ExportScale = 3.0,
    # Emit the on-sheet audit ("check") columns (I/J/K) after the A..H data
    # columns in each output workbook -- the worksheet-side duration
    # re-derivation (=E-D), its T/F compare against the written duration, and
    # the record-count check (ProcessTimeCheck.ps1's Get-ProcessTimeCheckColumnSpec).
    # $false writes A..H data only and skips the whole check-column pass.
    [bool]$EmitCheckColumns = $true,

    # -- Old-snap 9->3 hand-verification (docs/ProcessTime-OldSnap-Verify-Plan.md) --
    # D1 + deterministic triage of the finite backlog of OLD snaps that have
    # only a low-res PNG (no immune Ctrl+A .txt) and so fall back to OCR.
    # Master gate; when $false none of the below has any effect.
    [bool]$OldSnapVerifyEnabled = $true,
    # D1: make each output row's correl-id cell a clickable hyperlink to that
    # correl's snap image, so any flagged row is one click from human review.
    [bool]$OldSnapEmitHyperlink = $true,
    # The kenshou (Verify) verdict column appended after the audit columns:
    # 'txt' (trusted) / 'OCR-OK' (auto-confirmed) / needs-check / no-image.
    [bool]$OldSnapEmitVerifyColumn = $true,
    # D2 per-digit 3/9 image discrimination. Default OFF until the Phase-0
    # separability gate passes (mock-page/pixeldiff prototype); the COM/GDI
    # wiring is static-checked only and confirmed on an office PC.
    [bool]$OldSnapPixelDiff = $false,
    [double]$OldSnapPixelThreshold = 0.15,
    [string]$OldSnapRenderFont = 'MS Gothic',
    # {0} = side stage (GIFT/GFIX); matches the snap layout HmSnap.ps1 writes.
    [string]$OldSnapDirPattern = 'snap\{0}_HM',
    # Optional cross-engine (en-US vs ja) digit disagreement flag. Reserved;
    # default OFF (en drops many fields -- weak but free when it does read).
    [bool]$OldSnapCrossEngine = $false,

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
. (Join-Path $PSScriptRoot 'ProcessTimeCheck.ps1')
. (Join-Path $PSScriptRoot 'OldSnapVerify.ps1')
. (Join-Path $PSScriptRoot 'PixelDigitMatch.ps1')
. (Join-Path $PSScriptRoot 'OldSnapPixelVerify.ps1')

# System.Drawing backs the OCR image preprocessing (upscale + grayscale +
# contrast). Warn-only: a load failure just disables preprocessing (the
# original image is OCR'd as before), it never blocks the phase.
try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch {
    Write-Host ("[WARN] System.Drawing unavailable -- OCR image preprocessing disabled: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    $OcrPreprocess = $false
}

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

# Preprocesses one PNG for OCR via System.Drawing: high-quality-bicubic
# upscale (capped below the WinRT OCR engine's MaxImageDimension, ~2600 px
# -- an oversized bitmap makes RecognizeAsync fail outright), then a single
# ColorMatrix pass combining grayscale + linear contrast stretch around
# mid-gray. Root-cause fix for the recurring ja-recognizer digit misreads
# ('9' read as '3' in time-of-day / datestamp / count fields, JIGPC06S):
# the misread yields a FORMAT-VALID timestamp, so no post-hoc regex check
# can catch it -- the pixels themselves have to be unambiguous before the
# engine sees them (the same reason Snipping Tool's extraction reads the
# page cleanly). Writes '<stem>_pre.png' next to the dump files (cleaned
# up by the per-run stale-artifact pass like every other candidate
# artifact) and returns its path; returns the ORIGINAL path unchanged on
# any failure or when preprocessing is off -- never blocks OCR.
function ConvertTo-ProcessTimeOcrImage {
    param([string]$Png, [string]$OutDir, [string]$Stem,
          [double]$Scale = 2.0, [double]$Contrast = 1.3, [int]$MaxDimension = 2500)
    $src = $null; $dst = $null; $g = $null; $ms = $null
    try {
        $bytes = [System.IO.File]::ReadAllBytes($Png)
        $ms = New-Object System.IO.MemoryStream (,$bytes)
        $src = [System.Drawing.Bitmap]::FromStream($ms)

        $longest = [Math]::Max($src.Width, $src.Height)
        $eff = $Scale
        if ($longest * $eff -gt $MaxDimension) { $eff = [double]$MaxDimension / [double]$longest }
        if ($eff -lt 1.0) { $eff = 1.0 }   # never downscale an already-large image

        $w = [int][Math]::Round($src.Width * $eff)
        $h = [int][Math]::Round($src.Height * $eff)
        $dst = New-Object System.Drawing.Bitmap ($w, $h, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $dst.SetResolution(96, 96)
        $g = [System.Drawing.Graphics]::FromImage($dst)
        $g.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::HighQualityBicubic
        $g.PixelOffsetMode   = [System.Drawing.Drawing2D.PixelOffsetMode]::HighQuality

        # One ColorMatrix = grayscale (Rec.601 luma weights) x contrast
        # stretch around 0.5. t re-centers so mid-gray stays put.
        $c = $Contrast
        if ($c -le 0) { $c = 1.0 }
        $lr = 0.299 * $c; $lg = 0.587 * $c; $lb = 0.114 * $c
        $t = 0.5 * (1.0 - $c)
        # ColorMatrix row i = input channel i's weight into each output
        # column: every output RGB column gets the same luma mix, so row 0
        # (input R) is lr across columns 0-2, row 1 (input G) lg, row 2
        # (input B) lb; row 4 is the additive re-center term.
        $rows = New-Object 'single[][]' 5
        for ($i = 0; $i -lt 5; $i++) { $rows[$i] = New-Object 'single[]' 5 }
        $rows[0][0] = $lr; $rows[0][1] = $lr; $rows[0][2] = $lr
        $rows[1][0] = $lg; $rows[1][1] = $lg; $rows[1][2] = $lg
        $rows[2][0] = $lb; $rows[2][1] = $lb; $rows[2][2] = $lb
        $rows[3][3] = 1.0
        $rows[4][0] = $t; $rows[4][1] = $t; $rows[4][2] = $t; $rows[4][4] = 1.0
        $matrix = New-Object System.Drawing.Imaging.ColorMatrix (,$rows)
        $attrs = New-Object System.Drawing.Imaging.ImageAttributes
        $attrs.SetColorMatrix($matrix)

        $destRect = New-Object System.Drawing.Rectangle (0, 0, $w, $h)
        $g.DrawImage($src, $destRect, 0, 0, $src.Width, $src.Height, [System.Drawing.GraphicsUnit]::Pixel, $attrs)

        $outPath = Join-Path $OutDir ($Stem + '_pre.png')
        $dst.Save($outPath, [System.Drawing.Imaging.ImageFormat]::Png)
        Write-Host ("       [OCR] preprocessed {0} -> {1} ({2}x{3} -> {4}x{5}, contrast {6})" -f `
            (Split-Path $Png -Leaf), (Split-Path $outPath -Leaf), $src.Width, $src.Height, $w, $h, $Contrast) -ForegroundColor DarkGray
        return $outPath
    } catch {
        Write-Host ("       [WARN] OCR image preprocessing failed ({0}); using the original image: {1}" -f (Split-Path $Png -Leaf), $_.Exception.Message) -ForegroundColor Yellow
        return $Png
    } finally {
        if ($null -ne $g)   { try { $g.Dispose() } catch {} }
        if ($null -ne $dst) { try { $dst.Dispose() } catch {} }
        if ($null -ne $src) { try { $src.Dispose() } catch {} }
        if ($null -ne $ms)  { try { $ms.Dispose() } catch {} }
    }
}

# OCRs one exported candidate PNG with the pooled recognizer languages and
# dumps the reconstructed rows to <png-stem>.ocr.txt next to it. Returns
# the pooled line array (plain array; callers wrap in @()). When
# $OcrPreprocess (script param) is on, the picture is first run through
# ConvertTo-ProcessTimeOcrImage and BOTH recognizers read the preprocessed
# copy; the dump keeps its original stem so downstream tooling is unchanged.
function Read-ProcessTimeOcrLines {
    param([string]$Png, [string]$SecondaryLanguage, [string]$OutDir, [string]$DumpBaseName = '')
    $langs = @('en-US', $SecondaryLanguage) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) } | Select-Object -Unique
    $pooled  = New-Object System.Collections.Generic.List[string]
    $enLines = New-Object System.Collections.Generic.List[string]

    # Preprocess once; both recognizers read the same preprocessed copy.
    # ($OcrPreprocess/-Scale/-Contrast are this script's params, visible
    # here through PowerShell's dynamic scoping.)
    $ocrPng = $Png
    if ($OcrPreprocess) {
        $preStem = if ([string]::IsNullOrWhiteSpace($DumpBaseName)) { [System.IO.Path]::GetFileNameWithoutExtension($Png) } else { $DumpBaseName }
        $ocrPng = ConvertTo-ProcessTimeOcrImage -Png $Png -OutDir $OutDir -Stem $preStem -Scale $OcrPreprocessScale -Contrast $OcrPreprocessContrast
    }

    foreach ($lang in $langs) {
        try {
            Write-Host ("       [OCR] {0} lang={1}{2}" -f (Split-Path $Png -Leaf), $lang, $(if ($ocrPng -ne $Png) { ' (preprocessed)' } else { '' })) -ForegroundColor DarkGray
            $ocr = Invoke-WinOcrFile -Path $ocrPng -LanguageTag $lang
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
        $dumpStem = if ([string]::IsNullOrWhiteSpace($DumpBaseName)) { [System.IO.Path]::GetFileNameWithoutExtension($Png) } else { $DumpBaseName }
        $dumpPath = Join-Path $OutDir ($dumpStem + '.ocr.txt')
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
#
# Between the archived text and the evidence-workbook tiers sits the
# snap-PNG tier (-SnapPngPath, snap\<Stage>_HM\<correl>.png): the ORIGINAL
# screenshot HmSnap saved for exactly this correl, OCR'd directly. It is a
# better OCR target than the evidence-workbook copy (never rescaled by the
# paste, and its per-correl file name makes ownership certain, so no
# correl-seen gate is needed), and it avoids exporting workbook pictures
# entirely when it reads cleanly. Full priority order per side:
#   snap .txt -> snap .png OCR -> evidence-workbook picture OCR.
function Resolve-ProcessTimeSide {
    param($Workbook, [string]$SheetName, [string]$CorrelId, [string]$SnapTextPath,
          [string]$OutDir, [int]$AnchorCol, [string]$SecondaryLanguage, [double]$Scale,
          [string]$ExportBaseName = '', [string]$SnapPngPath = '')

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

    # Deterministic per-side base name (GIFT_/GFIX_) so the two sides of one
    # correl don't collide on the same PNG/dump names in the shared per-correl
    # export dir. Clear stale artifacts from a previous run so a MISS this run
    # can't be masked by last run's leftover PNG/dump. Done before the
    # snap-PNG tier so its own dump also starts clean (and is not deleted by
    # this pass afterwards).
    $base = if ([string]::IsNullOrWhiteSpace($ExportBaseName)) { $CorrelId } else { $ExportBaseName }
    if (-not (Test-Path -LiteralPath $OutDir)) {
        try { New-Item -ItemType Directory -Path $OutDir -Force | Out-Null } catch {}
    }
    if (Test-Path -LiteralPath $OutDir) {
        foreach ($pat in @(('{0}_*.png' -f $base), ('{0}_*.txt' -f $base))) {
            Get-ChildItem -LiteralPath $OutDir -Filter $pat -File -ErrorAction SilentlyContinue |
                ForEach-Object { try { Remove-Item -LiteralPath $_.FullName -Force -ErrorAction Stop } catch {} }
        }
    }

    $accepted = $null; $acceptedTag = ''
    $fallbackRow = $null; $fallbackRank = -1; $fallbackTag = ''
    $candTotal = 0
    $sectionHadPicture = $false
    $missNotes = New-Object System.Collections.Generic.List[string]

    # Tier 1.5: OCR the snap-time HM screenshot itself (snap\<Stage>_HM\
    # <correl>.png). It was captured FOR this correl (per-correl file name),
    # so ownership is certain and a full time row is accepted on that alone,
    # same trust level as the evidence workbook's section tier -- and it is
    # the cleaner OCR target (never rescaled by the evidence paste). Only
    # when it is absent or unreadable does the evidence-workbook picture
    # search below run at all.
    if (-not [string]::IsNullOrWhiteSpace($SnapPngPath) -and (Test-Path -LiteralPath $SnapPngPath)) {
        $candTotal++
        $read = Read-ProcessTimeOcrLines -Png $SnapPngPath -SecondaryLanguage $SecondaryLanguage -OutDir $OutDir -DumpBaseName ($base + '_snapocr')
        $lines = @($read.Lines)
        $rows = @(ConvertFrom-ProcessTimeOcrLines -Lines $lines -CorrelId $CorrelId -StartDateHints $read.DateHints)
        $sel = Select-ProcessTimeRow -Rows $rows
        if ($null -ne $sel) {
            $rank = Get-ProcessTimeRowRank $sel
            if ($rank -ge 2) {
                $accepted = $sel; $acceptedTag = 'snap-png'
            } elseif ($rank -gt $fallbackRank) {
                $fallbackRow = $sel; $fallbackRank = $rank; $fallbackTag = 'snap-png'
            }
        }
        if ($null -eq $accepted) {
            $note = if ($null -ne $sel) { 'partial time row only (kept as fallback)' }
                    elseif ($rows.Count -gt 0) { ("{0} time row(s), all before 09:00 -- skipped as history" -f $rows.Count) }
                    else { Get-ProcessTimeOcrMissNote -Lines $lines }
            $missNotes.Add(("{0}: {1}" -f (Split-Path $SnapPngPath -Leaf), $note))
            Write-Host ("       [DIAG] snap png {0}: {1}" -f (Split-Path $SnapPngPath -Leaf), $note) -ForegroundColor Yellow
        }
    }

    # Tier 2: OCR of the HM screenshot already inserted into the evidence
    # workbook -- only reached when neither snap artifact resolved the side.
    $tiers = @()
    if ($null -eq $accepted) {
        $ws = Get-SheetByName $Workbook $SheetName
        if ($null -eq $ws) {
            $missNotes.Add(("sheet '{0}' not found in workbook" -f $SheetName))
            Write-Host ("       [MISS] {0}: sheet '{1}' not found in workbook" -f $CorrelId, $SheetName) -ForegroundColor Yellow
        } else {
            $labelCell = Find-ProcessTimeCorrelCell $ws $CorrelId $AnchorCol
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
        }
    }

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
            $result.Note = if ($missNotes.Count -gt 0) { ($missNotes -join '; ') }
                           else { ("no exportable picture found for the label on sheet '{0}'" -f $SheetName) }
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

# Set-ProcessTimeCheckColumns (COM)
#   Writes the audit ("check") columns from ProcessTimeCheck.ps1's spec into
#   an already-populated ProcessTime worksheet: per spec entry, the header
#   cell (HeaderRow), a per-row formula (FirstDataRow..LastDataRow, via the
#   pure New-ProcessTimeCheckFormula) and the column NumberFormat. The spec
#   formulas are SELF-GUARDING (blank/text source cells -> "" in the cell),
#   so the "a partial row stays blank" rule holds with no per-row inspection
#   here. Value writes go through Set-RangeValue2 (same PS 5.1 COM binder
#   hazard as the data write loop). Headers are the spec's [char]-built
#   strings, so this stays ASCII source.
function Set-ProcessTimeCheckColumns {
    param($Worksheet, [int]$HeaderRow, [int]$FirstDataRow, [int]$LastDataRow, [object[]]$Spec)
    foreach ($col in @($Spec)) {
        $ci = [int]$col.ColIndex
        $hc = $Worksheet.Cells.Item($HeaderRow, $ci)
        Set-RangeValue2 $hc ([string]$col.Header) | Out-Null
        try { $hc.Font.Bold = $true } catch {}
        $nf       = [string]$col.NumberFormat
        $template = [string]$col.Formula
        for ($r = $FirstDataRow; $r -le $LastDataRow; $r++) {
            $cell = $Worksheet.Cells.Item($r, $ci)
            if (-not [string]::IsNullOrEmpty($nf)) {
                try { $cell.NumberFormat = $nf } catch {}
            }
            $cell.Formula = (New-ProcessTimeCheckFormula -Template $template -Row $r)
        }
    }
}

# Writes the operator-facing, vertical layout requested for delivery:
# No. / GIFT-GFIX / correl / start / end / duration / count / job (cols
# A..H). Each -Rows entry (the combined GIFT+GFIX shape ProcessTime.ps1
# resolves, cached per-correl in the sidecar) becomes TWO output rows, one
# per side. Incremental (non -Force) runs retain older records; a repeated
# (side, correl) pair REPLACES its prior row (deleted first), so a redo
# keeps one row per (side, correl) instead of stacking duplicates.
#
# The on-sheet audit ("check") columns after the data (I/J/K -- duration
# re-derivation, T/F compare, count check) are NOT written here: when
# -EmitCheckColumns, Set-ProcessTimeCheckColumns applies them uniformly
# from ProcessTimeCheck.ps1's data-driven spec after all data rows exist.
#
# Cell writes go through ExcelHelpers.ps1's Set-RangeValue2 (not a bare
# Value2 assign): col 1 (No.) is an [int] while others are [string], and
# PS 5.1's COM property-set binder can cache a conversion rule from one
# Value2 call and misapply it to a later one of a different type on the
# same sheet ("Unable to cast ... Int32 to ... String" -- same root cause
# as FillCheckSheet.ps1's v2.10.8 date write bug); Set-RangeValue2 retries
# via reflection on that failure.
# A..Z worksheet column letter for a 1-based index (the ProcessTime table
# never exceeds a couple of dozen columns, so single-letter is enough).
function Get-ProcessTimeColLetter {
    param([int]$Index)
    return [string][char]([int][char]'A' + $Index - 1)
}

function Write-ProcessTimeWorkbook {
    param($Excel, [string]$OutputPath, [string]$SheetName, [object[]]$Rows,
          [bool]$EmitCheckColumns = $true,
          # Old-snap 9->3 hand-verification (D1 hyperlink + kenshou verify column).
          [bool]$EmitVerifyColumn = $false, [bool]$EmitHyperlink = $false,
          [string]$WorkDir = '', [string]$SnapDirPattern = 'snap\{0}_HM',
          [bool]$PixelEnabled = $false, [string]$PixelFont = 'MS Gothic',
          [double]$PixelMinMargin = 0.04, [hashtable]$PixelGeometry = $null)

    # A..H data headers only; the I/J/K check headers come from the spec.
    $headers = @(
        'No.', 'GIFT/GFIX',
        ([string][char]0x76F8 + [char]0x95A2 + 'ID'),
        ([string][char]0x958B + [char]0x59CB + [char]0x65E5 + [char]0x6642),
        ([string][char]0x7D42 + [char]0x4E86 + [char]0x65E5 + [char]0x6642),
        ([string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593),
        ([string][char]0x51E6 + [char]0x7406 + [char]0x4EF6 + [char]0x6570),
        ([string][char]0x30B8 + [char]0x30E7 + [char]0x30D6)
    )
    $checkSpec = if ($EmitCheckColumns) { @(Get-ProcessTimeCheckColumnSpec) } else { @() }
    # Column layout: 8 data columns (A..H), then the check columns (I/J/K when
    # emitted), then the single kenshou verify column (old-snap triage) when
    # emitted. The verify column's index is positional, so it is computed here
    # rather than baked into a spec.
    $dataColCount   = 8
    $verifySpec     = if ($EmitVerifyColumn) { Get-OldSnapVerifyColumnSpec } else { $null }
    $verifyColIndex = if ($EmitVerifyColumn) { $dataColCount + $checkSpec.Count + 1 } else { 0 }
    $lastColIndex   = $dataColCount + $checkSpec.Count + $(if ($EmitVerifyColumn) { 1 } else { 0 })
    # Drives every A1-range the formatting uses.
    $lastColLetter  = Get-ProcessTimeColLetter $lastColIndex

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
                    Source = [string]$(if ($isGift) { $r.GiftSource } else { $r.GfixSource })
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

        # kenshou (Verify) column header -- written even on an existing workbook
        # whose earlier run predated this feature, so the column is always
        # labelled. Its index sits after the audit columns (positional).
        if ($EmitVerifyColumn -and $verifyColIndex -gt 0) {
            try {
                $vh = $ws.Cells.Item(1, $verifyColIndex)
                Set-RangeValue2 $vh ([string]$verifySpec.Header) | Out-Null
                $vh.Font.Bold = $true
                if (-not [string]::IsNullOrEmpty([string]$verifySpec.NumberFormat)) {
                    try { $ws.Columns.Item($verifyColIndex).NumberFormat = [string]$verifySpec.NumberFormat } catch {}
                }
            } catch {
                Write-Warning ("ProcessTime workbook formatting: verify header failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
            }
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

        $verifyNeedsCheck = 0   # count of rows this write flagged you-kaku-nin (needs check)
        $row = $lastRow + 1
        foreach ($r in $flatRows) {
            Set-RangeValue2 $ws.Cells.Item($row, 1) ($row - 1) | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 2) $r.Side | Out-Null
            $correlCell = $ws.Cells.Item($row, 3)
            Set-RangeValue2 $correlCell $r.CorrelId | Out-Null

            # -- Old-snap 9->3 hand-verification (D1 hyperlink + kenshou verdict) --
            # Both are per NEW row; retained rows keep whatever a prior run
            # wrote. The snap-path existence check (I/O) lives here on the COM
            # side; the path build + verdict are pure (OldSnapVerify.ps1).
            if ($EmitVerifyColumn -or $EmitHyperlink) {
                $snapPath = Resolve-OldSnapImagePath -WorkDir $WorkDir -Side $r.Side -CorrelId $r.CorrelId -DirPattern $SnapDirPattern
                $snapExists = (-not [string]::IsNullOrWhiteSpace($snapPath)) -and (Test-Path -LiteralPath $snapPath)
                if ($EmitHyperlink -and $snapExists) {
                    try { $ws.Hyperlinks.Add($correlCell, $snapPath) | Out-Null } catch {}
                }
                if ($EmitVerifyColumn -and $verifyColIndex -gt 0) {
                    $arith = Test-OldSnapDurationArithmetic -Start $r.Start -End $r.End -Duration $r.Duration
                    # D2 per-digit 3/9 image check (only when enabled AND a snap
                    # exists). The crop geometry per field is office-PC-calibrated
                    # (PixelGeometry); without it Get-OldSnapRowPixelVerdict
                    # returns '' -> the verdict falls back to a conservative flag.
                    $pixelResult = ''
                    if ($PixelEnabled -and $snapExists) {
                        $fields = @(
                            @{ Text = $r.Start;    Geometry = $(if ($null -ne $PixelGeometry) { $PixelGeometry['Start'] } else { $null }) },
                            @{ Text = $r.End;      Geometry = $(if ($null -ne $PixelGeometry) { $PixelGeometry['End'] } else { $null }) },
                            @{ Text = $r.Duration; Geometry = $(if ($null -ne $PixelGeometry) { $PixelGeometry['Duration'] } else { $null }) }
                        )
                        $pixelResult = Get-OldSnapRowPixelVerdict -SnapPath $snapPath -Fields $fields -FontName $PixelFont -MinMargin $PixelMinMargin
                    }
                    $verdict = Get-OldSnapVerifyVerdict -Source $r.Source -SnapExists $snapExists `
                        -ArithmeticOk $arith -PixelResult $pixelResult -PixelEnabled $PixelEnabled
                    if ($verdict -eq 'NeedsCheck') { $verifyNeedsCheck++ }
                    Set-RangeValue2 $ws.Cells.Item($row, $verifyColIndex) (Get-OldSnapVerifyLabel $verdict) | Out-Null
                }
            }

            # Start/End/Duration are written as REAL Excel date/time values
            # (OADate serial + NumberFormat, format set BEFORE the value per
            # the v2.10.7 FillCheckSheet lesson) so the I/J check-formula
            # columns can compute; an unparseable value falls back to the old
            # plain-text write so nothing is ever lost.
            $startDt = ConvertTo-ProcessTimeDateTimeValue $r.Start
            $endDt   = ConvertTo-ProcessTimeDateTimeValue $r.End
            $durVal  = ConvertTo-ProcessTimeDurationValue $r.Duration
            foreach ($spec in @(
                @{ Col = 4; Dt = $startDt; Text = $r.Start },
                @{ Col = 5; Dt = $endDt;   Text = $r.End })) {
                $cell = $ws.Cells.Item($row, $spec.Col)
                if ($null -ne $spec.Dt) {
                    try { $cell.NumberFormat = 'yyyy/mm/dd hh:mm:ss' } catch {}
                    Set-RangeValue2 $cell ([double]$spec.Dt.ToOADate()) | Out-Null
                } else {
                    Set-RangeValue2 $cell $spec.Text | Out-Null
                }
            }
            $durCell = $ws.Cells.Item($row, 6)
            if ($null -ne $durVal) {
                try { $durCell.NumberFormat = '[h]:mm:ss' } catch {}
                Set-RangeValue2 $durCell ([double]$durVal) | Out-Null
            } else {
                Set-RangeValue2 $durCell $r.Duration | Out-Null
            }

            Set-RangeValue2 $ws.Cells.Item($row, 7) $r.Count | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 8) $r.JobName | Out-Null
            $row++
        }
        # Renumber retained + appended records and make the range filterable.
        #
        # Template formatting reference (TODO: extract to a formatting module):
        #   header A1:<last>1 fill 10053120 (BGR teal) + white bold centered;
        #   body font 'Yu Gothic' 11pt, 18pt rows, thin borders, v-centered
        #   (col A also h-centered); per-row fill by side (GIFT 15398626 /
        #   GFIX 16777215 white, both BGR Longs); fixed column widths. Cols
        #   D/E are real date/time values (OADate + 'yyyy/mm/dd hh:mm:ss') and
        #   F a real time serial ('[h]:mm:ss'); the I/J/K audit columns and
        #   their headers/formats are written by Set-ProcessTimeCheckColumns
        #   from ProcessTimeCheck.ps1's spec (skipped when -EmitCheckColumns
        #   is off, in which case the table is A..H only).
        try {
            $finalRow = [int]$ws.Cells.Item($ws.Rows.Count, 2).End($xlUp).Row
            for ($rr = 2; $rr -le $finalRow; $rr++) { Set-RangeValue2 $ws.Cells.Item($rr, 1) ($rr - 1) | Out-Null }
            # Use A1 addresses rather than passing two Range COM proxies to
            # Worksheet.Range.  Some office-PC Excel/PowerShell combinations
            # reject the proxy overload with DISP_E_TYPEMISMATCH ("argument
            # type mismatch"), which must never prevent the workbook save.
            $tableRange = $ws.Range(('A1:{0}{1}' -f $lastColLetter, $finalRow))
            $headerRange = $ws.Range(('A1:{0}1' -f $lastColLetter))
        } catch {
            Write-Warning ("ProcessTime workbook formatting: could not resolve table/header range for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        # Audit (check) columns I/J/K: uniform header + per-row formula +
        # number-format pass over every data row, from ProcessTimeCheck.ps1's
        # data-driven spec. Self-guarding formulas leave a partial row (e.g.
        # end time not read) blank on their own, so no per-row inspection is
        # needed here. Skipped entirely when -EmitCheckColumns is off.
        if ($checkSpec.Count -gt 0 -and $finalRow -ge 2) {
            try {
                Set-ProcessTimeCheckColumns -Worksheet $ws -HeaderRow 1 -FirstDataRow 2 -LastDataRow $finalRow -Spec $checkSpec
            } catch {
                Write-Warning ("ProcessTime workbook formatting: check columns failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
            }
        }

        try {
            $tableRange.AutoFilter() | Out-Null
            $tableRange.Borders.LineStyle = 1
        } catch {
            Write-Warning ("ProcessTime workbook formatting: AutoFilter/borders failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        try {
            $headerRange.Interior.Color = 10053120
            $headerRange.Font.Color = 16777215
            $headerRange.Font.Bold = $true
        } catch {
            Write-Warning ("ProcessTime workbook formatting: header fill/font failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        try {
            # The reference sheet uses Excel's Japanese default typeface at
            # 11 pt throughout, with compact 18-point rows.  Apply this to
            # the complete table after incremental rows have been merged so
            # retained and newly appended records are visually identical.
            $tableRange.Font.Name = 'Yu Gothic'
            $tableRange.Font.Size = 11
            $tableRange.Rows.RowHeight = 18
        } catch {
            Write-Warning ("ProcessTime workbook formatting: font/row-height failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        try {
            $headerRange.HorizontalAlignment = -4108 # xlCenter
            $headerRange.VerticalAlignment = -4108   # xlCenter
            $ws.Range(('A2:A{0}' -f $finalRow)).HorizontalAlignment = -4108
            $tableRange.VerticalAlignment = -4108
        } catch {
            Write-Warning ("ProcessTime workbook formatting: alignment failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        try {
            for ($rr = 2; $rr -le $finalRow; $rr++) {
                $side = [string]$ws.Cells.Item($rr, 2).Value2
                # Color is explicitly Double for the Excel COM Variant
                # binder; using the branch's boxed Int32 can also produce a
                # type mismatch on older Windows PowerShell/Excel versions.
                [double]$fill = if ($side -eq 'GIFT') { 15398626 } else { 16777215 }
                $ws.Range(('A{0}:{1}{0}' -f $rr, $lastColLetter)).Interior.Color = $fill
            }
        } catch {
            Write-Warning ("ProcessTime workbook formatting: GIFT/GFIX row fill failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        try {
            # Fixed widths reproduce the template proportions and keep the
            # date/time fields from changing width with different data. The
            # A..H data widths are fixed here; each check column's width comes
            # from its spec entry so the widths track the emitted columns.
            $widths = New-Object System.Collections.Generic.List[double]
            foreach ($w in @(9, 10, 14, 24, 24, 20, 14, 17)) { $widths.Add([double]$w) }
            foreach ($col in $checkSpec) { $widths.Add([double]$col.Width) }
            if ($EmitVerifyColumn -and $null -ne $verifySpec) { $widths.Add([double]$verifySpec.Width) }
            for ($c = 1; $c -le $widths.Count; $c++) {
                $ws.Columns.Item($c).ColumnWidth = $widths[$c - 1]
            }
        } catch {
            Write-Warning ("ProcessTime workbook formatting: column widths failed for '{0}': {1}" -f $OutputPath, $_.Exception.Message)
        }

        if ($EmitVerifyColumn -and $verifyNeedsCheck -gt 0) {
            $lblNeeds = Get-OldSnapVerifyLabel 'NeedsCheck'
            Write-Host ("  [OldSnapVerify] {0} row(s) flagged '{1}' in {2}" -f $verifyNeedsCheck, $lblNeeds, (Split-Path $OutputPath -Leaf)) -ForegroundColor Yellow
        }

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
                    $giftPng   = Join-Path (Join-Path $WorkDir 'snap\GIFT_HM') ("{0}.png" -f $correlId)
                    $gfixPng   = Join-Path (Join-Path $WorkDir 'snap\GFIX_HM') ("{0}.png" -f $correlId)
                    $exportDir = Join-Path $exportRoot $correlId

                    $giftResult = Resolve-ProcessTimeSide -Workbook $wb -SheetName $sheetGiftRecv -CorrelId $correlId `
                        -SnapTextPath $giftTxt -SnapPngPath $giftPng -OutDir $exportDir -AnchorCol $AnchorCol -SecondaryLanguage $OcrLanguage -Scale $ExportScale `
                        -ExportBaseName ("GIFT_{0}" -f $correlId)
                    $gfixResult = Resolve-ProcessTimeSide -Workbook $wb -SheetName $sheetGfixRecv -CorrelId $correlId `
                        -SnapTextPath $gfixTxt -SnapPngPath $gfixPng -OutDir $exportDir -AnchorCol $AnchorCol -SecondaryLanguage $OcrLanguage -Scale $ExportScale `
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
                $derivedTags = @{}
                foreach ($r in $results) {
                    $tag = Get-ProcessTimeOutputTag -ExcelName ([string]$r.ExcelName) -Tags $OutputTags -UnclassifiedTag $UnclassifiedTag -DeriveFromName $AutoDeriveTag
                    if ($tag -ne $UnclassifiedTag -and (@($OutputTags) -notcontains $tag)) { $derivedTags[$tag] = $true }
                    if (-not $buckets.ContainsKey($tag)) { $buckets[$tag] = New-Object System.Collections.Generic.List[object]; $bucketOrder.Add($tag) }
                    $buckets[$tag].Add($r)
                }
                if ($derivedTags.Count -gt 0) {
                    Write-Host ("  [INFO] output tag(s) derived from the Excel_NAME '?XXX????' pattern (not in OutputTags): {0}" -f (($derivedTags.Keys | Sort-Object) -join ', ')) -ForegroundColor DarkGray
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
                    Write-ProcessTimeWorkbook -Excel $excel -OutputPath $path -SheetName $OutputSheetName -Rows $rowsForTag -EmitCheckColumns $EmitCheckColumns `
                        -EmitVerifyColumn ($OldSnapVerifyEnabled -and $OldSnapEmitVerifyColumn) `
                        -EmitHyperlink ($OldSnapVerifyEnabled -and $OldSnapEmitHyperlink) `
                        -WorkDir $WorkDir -SnapDirPattern $OldSnapDirPattern `
                        -PixelEnabled ($OldSnapVerifyEnabled -and $OldSnapPixelDiff) `
                        -PixelFont $OldSnapRenderFont -PixelMinMargin $OldSnapPixelThreshold
                    $written.Add($path)
                    $writtenCorrelCount += $rowsForTag.Count
                    # Only the rows that made it into a SUCCESSFULLY written
                    # workbook get their write bit set; a row whose tag's
                    # workbook failed to save stays pending for the next run.
                    # Every mapping row sharing a correl (duplicates included)
                    # is marked, not just the first. The bucket-array helper
                    # is REQUIRED here: @() over a List[object] pulled out of
                    # a hashtable can throw "Argument types do not match" on
                    # PS 5.1 -- a real office-PC run hit exactly that AFTER
                    # the workbook saved, so the tag was reported both FAILED
                    # and written and no write bit was ever set.
                    foreach ($r in $rowsForTag) {
                        foreach ($mrow in (ConvertTo-ProcessTimeBucketArray -Bucket $writtenRowsByCorrel[$r.CorrelId])) {
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
