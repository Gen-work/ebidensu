# ============================================================
#  Mark.ps1
#
#  Phase: MarkGift / MarkGfix / MarkDf
#
#  Walks each evidence workbook and, for every Picture shape stamped
#  with a metadata payload (set by ReplaceEvidence.ps1), draws red
#  rectangles relative to the picture's top-left corner.
#
#  Box geometry comes from the -BoxesConfig hashtable (filled from
#  VerifyConfig.psd1's Mark.Boxes). Empty list for a folder = no marks.
#
#  Idempotent: existing rectangles whose Name starts with the configured
#  prefix (default 'verifyMark_') are deleted first.
#
#  Sets isMarked |= bit on all rows in the group (1=Gift, 2=Gfix, 4=Df).
#
#  Usage:
#    .\Mark.ps1 -Mode Gift
#    .\Mark.ps1 -Mode Gfix -TargetIds JIGPL48S
#    .\Mark.ps1 -Mode Df -Force
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [ValidateSet('Gift','Gfix','Df')]
    [string]$Mode,

    [string]$WorkDir,
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [string]$ExcelPrefix = '',
    [switch]$Force,

    [string]$CommonScript = '',
    [string]$ExcelHelpersScript = '',

    [hashtable]$BoxesConfig = @{},
    [string]$NamePrefix = 'verifyMark_',
    [double]$LineWeight = 1.5,
    [string]$NoGfixNoteColumn = 'AZ',

    # Image-recognition box placement (opt-in per Mark.Boxes entry via a
    # 'Template' key). Folder searched for a bare Template filename;
    # <repo>\mark_templates is always tried too. Tolerance is the default
    # LockBits color tolerance (a box's own 'Tolerance' key overrides it).
    [string]$TemplateDir = '',
    [int]$ImageMatchTolerance = 15,

    # Optional stamp image inserted next to a 'verifyNote' annotation (see
    # EvidenceExecutor.ps1 -> Set-ShapeMetadata 'verifyNote'), keyed by the
    # note's Folder value (currently only 'GIFT_noGfixfile', the F4/M6
    # past-data annotation). Each entry: @{ Image; Column; RowOffset }.
    # Reuses the pixel rect already carried in the verifyNote payload (from
    # the snap-time <correl>.loc.json / .note.json sidecars) instead of
    # re-scanning the source PNG at Mark time. Empty/missing key = no stamp.
    [hashtable]$NoteStampConfig = @{},

    # GFIX log yellow-highlight settings. Folded in from the old standalone
    # MarkGfixLog phase: in -Mode Gfix the log "Command:" row is highlighted in
    # the same pass that draws the red rectangles (one workbook open, one bit).
    [string]$GfixLogAnchor = '',
    [string]$GfixLogCommandPattern = "Command:\s*'/appl/[A-Za-z0-9]+/shell/",
    [long]$GfixLogHighlightColor = 65535,
    [int]$GfixLogColStart = 2,
    [int]$GfixLogColEnd   = 51,
    [bool]$GfixLogAutoWidth = $true,
    [int]$GfixLogPadCols = 1,
    # Font the GFIX log was PASTED in (Replace.GfixLogFontName/GfixLogFontSize).
    # Used by the AutoWidth measurement so the computed highlight width always
    # matches the rendered text. Blank/0 -> measure with the cell's own font.
    [string]$GfixLogFontName = '',
    [double]$GfixLogFontSize = 0
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) { Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1 }

$forceFlag = [bool]$Force.IsPresent

# Default GFIX log anchor: ▼GFIXログ (kept ASCII via [char] code points).
if ([string]::IsNullOrWhiteSpace($GfixLogAnchor)) {
    $GfixLogAnchor = [char]0x25BC + 'GFIX' + [char]0x30ED + [char]0x30B0
}

# -- Dot-source ExcelHelpers.ps1 -----------------------------
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).ProviderPath; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
if (-not (Get-Command -Name 'Add-RedRectangle' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed (Add-RedRectangle not loaded).' -ForegroundColor Red; exit 1
}

# -- Image-recognition box placement (optional, per-box 'Template') --------
# Needs System.Drawing (source PNG pixel size) and Locate-ByImage.ps1 (LockBits
# template match, its own Add-Type). Both are loaded unconditionally here but
# only ever exercised when a Mark.Boxes entry actually carries a 'Template'
# key -- with no Template configured, behavior is byte-for-byte the old
# fixed-offset path.
try { Add-Type -AssemblyName System.Drawing -ErrorAction Stop } catch {
    Write-Host ("[WARN] System.Drawing unavailable ({0}); image-match boxes will fall back to fixed offsets." -f $_.Exception.Message) -ForegroundColor Yellow
}
$locateByImagePath = Join-Path $PSScriptRoot 'Locate-ByImage.ps1'
if (-not (Test-Path -LiteralPath $locateByImagePath)) { $locateByImagePath = '' }

# -- Row-position fallback chain for boxes with a 'RowHeight' key (GIFT_MQ) --
# SnapVerify.ps1 (pure MQ page-text parser + Get-MatchedRowIndex) and
# OcrWindows.ps1 (Windows built-in OCR) are both no-param() dot-source-safe
# per CLAUDE.md. Both are optional here -- a missing file only disables that
# one fallback tier (Get-Command guards in Get-MarkMqRowInfo* below), it
# never blocks Mark.
$snapVerifyPath = Join-Path $PSScriptRoot 'SnapVerify.ps1'
if (Test-Path -LiteralPath $snapVerifyPath) { . $snapVerifyPath }
$ocrWindowsPath = Join-Path $PSScriptRoot 'OcrWindows.ps1'
if (Test-Path -LiteralPath $ocrWindowsPath) { . $ocrWindowsPath }

function Resolve-MarkTemplatePath {
    <#
    Resolves a Mark.Boxes 'Template' value to a real file path. Tries the
    value as-is (absolute or relative to cwd), then <TemplateDir>\<value>,
    then <repo>\mark_templates\<value>. Returns $null when nothing matches.
    #>
    param([string]$Template, [string]$TemplateDir)
    if ([string]::IsNullOrWhiteSpace($Template)) { return $null }
    if (Test-Path -LiteralPath $Template) { return (Resolve-Path -LiteralPath $Template).ProviderPath }
    $candidates = @()
    if (-not [string]::IsNullOrWhiteSpace($TemplateDir)) { $candidates += (Join-Path $TemplateDir $Template) }
    $candidates += (Join-Path (Join-Path $PSScriptRoot 'mark_templates') $Template)
    foreach ($c in $candidates) {
        if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).ProviderPath }
    }
    return $null
}

function Get-ImagePixelSize {
    param([string]$Path)
    $img = $null
    try {
        $img = [System.Drawing.Image]::FromFile($Path)
        return @{ Width = [double]$img.Width; Height = [double]$img.Height }
    } finally {
        if ($null -ne $img) { $img.Dispose() }
    }
}

function Get-MarkTemplateHitFromSidecar {
    <#
    Reads <WorkDir>\snap\<Folder>\<Cid>.tplhit.json (written at snap time by
    JenkinsSnap.ps1 + SnapLocalize.ps1's Write-MarkTemplateHits, opt-in via
    the same 'Template' key Mark.Boxes already uses) and returns the entry
    for ($BoxIndex, $Template) as @{ X; Y; Width; Height } in source-PNG
    pixel units -- the same shape Locate-ByImage.ps1 returns, so callers can
    treat it identically. Returns $null on anything short of an exact match
    (file missing/unreadable, no entry for this box index, Template name
    differs from what is configured now, or the PNG the sidecar was recorded
    against is a different size than the current one) so a stale or absent
    sidecar always falls back to a live re-match rather than risk a wrong box.
    #>
    param(
        [string]$WorkDir,
        [string]$Folder,
        [string]$Cid,
        [int]$BoxIndex,
        [string]$Template,
        [int]$ExpectedSourceWidth,
        [int]$ExpectedSourceHeight
    )
    $sidecarPath = Join-Path $WorkDir ("snap\{0}\{1}.tplhit.json" -f $Folder, $Cid)
    if (-not (Test-Path -LiteralPath $sidecarPath)) { return $null }
    try {
        $data = Get-Content -LiteralPath $sidecarPath -Raw -Encoding UTF8 | ConvertFrom-Json
        if ([int]$data.SourceWidth -ne $ExpectedSourceWidth -or [int]$data.SourceHeight -ne $ExpectedSourceHeight) { return $null }
        $entry = @($data.Boxes) | Where-Object { [int]$_.Index -eq $BoxIndex -and [string]$_.Template -eq $Template } | Select-Object -First 1
        if ($null -eq $entry) { return $null }
        return [PSCustomObject]@{
            X      = [double]$entry.X
            Y      = [double]$entry.Y
            Width  = [double]$entry.Width
            Height = [double]$entry.Height
        }
    } catch {
        return $null
    }
}

function Find-MarkBoxByImage {
    <#
    Attempts template-match placement for one Mark.Boxes entry against the
    original snap PNG for ($Folder, $Cid) (<WorkDir>\snap\<Folder>\<Cid>.png
    -- the same file ReplaceEvidence pasted into the picture, per
    EvidencePlan.ps1's Get-SnapPath). Returns @{ Left; Top; Width; Height }
    in points, or $null when the box has no Template / files are missing / no
    match is found / anything errors -- caller falls back to the fixed
    OffsetX/OffsetY/Width/Height box in that case, so this never blocks Mark.

    Sizing: when the box ALSO configures its own 'Width'/'Height' (the same
    fields the fixed-offset fallback box uses), those are used as-is and the
    match only supplies the anchor (top-left corner) -- Template finds WHERE
    to draw, config says HOW BIG, so a small/unique anchor crop (e.g. a
    stable icon or label near the real target) does not force the drawn box
    to that crop's own dimensions. Without Width/Height, the box keeps the
    legacy behavior: sized to the matched crop's own pixel size, scaled to
    sheet points and padded by PadX/PadY on each side.

    Anchor source: a snap-time <correl>.tplhit.json sidecar (written by
    JenkinsSnap.ps1 when it was given this same box's Template, right after
    the screenshot was captured) is preferred over a fresh Locate-ByImage
    scan when present and not stale -- see Get-MarkTemplateHitFromSidecar for
    what counts as stale. Falls back to a live match (the original behavior)
    whenever the sidecar is missing.
    #>
    param(
        [hashtable]$Box,
        [string]$Folder,
        [string]$Cid,
        [int]$BoxIndex,
        [string]$WorkDir,
        [string]$TemplateDir,
        [int]$DefaultTolerance,
        [string]$LocateScript,
        $Shape
    )
    if (-not $Box.ContainsKey('Template')) { return $null }
    $tplName = [string]$Box.Template
    if ([string]::IsNullOrWhiteSpace($tplName)) { return $null }

    $srcPath = Join-Path $WorkDir ("snap\{0}\{1}.png" -f $Folder, $Cid)
    if (-not (Test-Path -LiteralPath $srcPath)) { return $null }

    try {
        $dim = Get-ImagePixelSize -Path $srcPath
        if ($dim.Width -le 0 -or $dim.Height -le 0) { return $null }
        $scaleX = ([double]$Shape.Width)  / $dim.Width
        $scaleY = ([double]$Shape.Height) / $dim.Height

        $hit = Get-MarkTemplateHitFromSidecar -WorkDir $WorkDir -Folder $Folder -Cid $Cid `
            -BoxIndex $BoxIndex -Template $tplName `
            -ExpectedSourceWidth ([int]$dim.Width) -ExpectedSourceHeight ([int]$dim.Height)
        $hitSource = 'sidecar'

        if ($null -eq $hit) {
            $hitSource = 'live'
            if ([string]::IsNullOrWhiteSpace($LocateScript)) { return $null }
            $tplPath = Resolve-MarkTemplatePath -Template $tplName -TemplateDir $TemplateDir
            if ($null -eq $tplPath) {
                Write-Host ("  [WARN] image-match template not found: {0}" -f $tplName) -ForegroundColor Yellow
                return $null
            }
            $tol = $DefaultTolerance
            if ($Box.ContainsKey('Tolerance')) { try { $tol = [int]$Box.Tolerance } catch {} }
            $hit = & $LocateScript -SourcePath $srcPath -TemplatePath $tplPath -Tolerance $tol -Quiet
            if ($null -eq $hit) { return $null }
        }

        $padX = 0.0; $padY = 0.0
        if ($Box.ContainsKey('PadX')) { try { $padX = [double]$Box.PadX } catch {} }
        if ($Box.ContainsKey('PadY')) { try { $padY = [double]$Box.PadY } catch {} }

        $px = [double]$hit.X - $padX
        $py = [double]$hit.Y - $padY

        if ($Box.ContainsKey('Width') -or $Box.ContainsKey('Height')) {
            $fw = 100.0; $fh = 20.0
            try { $fw = [double]$Box.Width }  catch {}
            try { $fh = [double]$Box.Height } catch {}
            return @{
                Left   = ([double]$Shape.Left) + ($px * $scaleX)
                Top    = ([double]$Shape.Top)  + ($py * $scaleY)
                Width  = $fw
                Height = $fh
                Source = $hitSource
            }
        }

        $pw = [double]$hit.Width  + (2 * $padX)
        $ph = [double]$hit.Height + (2 * $padY)

        return @{
            Left   = ([double]$Shape.Left) + ($px * $scaleX)
            Top    = ([double]$Shape.Top)  + ($py * $scaleY)
            Width  = $pw * $scaleX
            Height = $ph * $scaleY
            Source = $hitSource
        }
    } catch {
        Write-Host ("  [WARN] image-match failed for {0}\{1}: {2}" -f $Folder, $Cid, $_.Exception.Message) -ForegroundColor Yellow
        return $null
    }
}

# -- Row-position fallback chain (opt-in via a box's 'RowHeight' key) -------
#
# Some snap pages (GIFT_MQ) repeat a variable number of records for one
# correl (usually 2, sometimes 1 or 3+), and the box must always land on the
# LAST/newest one -- not always the 'BaseRow' the fixed OffsetY was
# calibrated against. Three tiers, each returning @{ RowIndex; NumRecords;
# Source } or $null (never throws):
#   1. sidecar   <correl>.mqrow.json written at snap time by MqSnap.ps1
#      (fastest + most accurate: computed while the live page was open).
#   2. txt       re-parse the archived Ctrl+A page capture <correl>.txt.
#   3. ocr       English-OCR the source PNG and parse that text the same way
#      (MQ records are ASCII/numeric -- a reasonable OCR candidate).
function Get-MarkMqRowInfoFromSidecar {
    param([string]$WorkDir, [string]$Folder, [string]$Cid)
    $sidecarPath = Join-Path $WorkDir ("snap\{0}\{1}.mqrow.json" -f $Folder, $Cid)
    if (-not (Test-Path -LiteralPath $sidecarPath)) { return $null }
    try {
        $data = Get-Content -LiteralPath $sidecarPath -Raw -Encoding UTF8 | ConvertFrom-Json
        $rowIndex = [int]$data.rowIndex
        if ($rowIndex -lt 1) { return $null }
        return [PSCustomObject]@{ RowIndex = $rowIndex; NumRecords = [int]$data.numRecords; Source = 'sidecar' }
    } catch {
        return $null
    }
}

# Shared tail for the txt / ocr tiers: parse page text with SnapVerify.ps1's
# pure MQ helpers and pick the target row with NO time window (Expected =
# $null -> newest overall / first matched). SnapVerify.TimeCheck is off by
# default, so this reproduces the snap-time pick in the common case; under a
# time-windowed run this fallback-of-a-fallback may pick a slightly different
# row than the sidecar would have -- acceptable since it only ever fires when
# the sidecar is already missing.
function ConvertTo-MarkMqRowInfo {
    param([string]$Text, [string]$Cid, [string]$Source)
    if ([string]::IsNullOrWhiteSpace($Text)) { return $null }
    if (-not (Get-Command -Name 'ConvertFrom-MqPageText' -ErrorAction SilentlyContinue)) { return $null }
    if (-not (Get-Command -Name 'Get-MatchedRowIndex' -ErrorAction SilentlyContinue)) { return $null }
    $parsed = ConvertFrom-MqPageText $Text
    $rowIndex = Get-MatchedRowIndex -Rows $parsed.Rows -CorrelId $Cid -DateProperty 'RecvDate' -Expected $null
    if ($rowIndex -lt 1) { return $null }
    return [PSCustomObject]@{ RowIndex = $rowIndex; NumRecords = @($parsed.Rows).Count; Source = $Source }
}

function Get-MarkMqRowInfoFromArchivedText {
    param([string]$WorkDir, [string]$Folder, [string]$Cid)
    $txtPath = Join-Path $WorkDir ("snap\{0}\{1}.txt" -f $Folder, $Cid)
    if (-not (Test-Path -LiteralPath $txtPath)) { return $null }
    try {
        $text = Get-Content -LiteralPath $txtPath -Raw -Encoding UTF8
        return ConvertTo-MarkMqRowInfo -Text $text -Cid $Cid -Source 'txt'
    } catch {
        return $null
    }
}

function Get-MarkMqRowInfoFromOcr {
    param([string]$WorkDir, [string]$Folder, [string]$Cid)
    if (-not (Get-Command -Name 'Invoke-WinOcrFile' -ErrorAction SilentlyContinue)) { return $null }
    $pngPath = Join-Path $WorkDir ("snap\{0}\{1}.png" -f $Folder, $Cid)
    if (-not (Test-Path -LiteralPath $pngPath)) { return $null }
    try {
        $ocr = Invoke-WinOcrFile -Path $pngPath -LanguageTag 'en'
        return ConvertTo-MarkMqRowInfo -Text ([string]$ocr.Text) -Cid $Cid -Source 'ocr'
    } catch {
        return $null
    }
}

function Get-MarkMqRowInfo {
    param([string]$WorkDir, [string]$Folder, [string]$Cid)
    $info = Get-MarkMqRowInfoFromSidecar -WorkDir $WorkDir -Folder $Folder -Cid $Cid
    if ($null -ne $info) { return $info }
    $info = Get-MarkMqRowInfoFromArchivedText -WorkDir $WorkDir -Folder $Folder -Cid $Cid
    if ($null -ne $info) { return $info }
    return Get-MarkMqRowInfoFromOcr -WorkDir $WorkDir -Folder $Folder -Cid $Cid
}

# -- Target filter -------------------------------------------
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

# -- Mode config (sheet names + which folders carry marks) ---
$sheetGiftRecv = "GIFT" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GIFT受信結果
$sheetGfixRecv = "GFIX" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GFIX受信結果
$sheetDfDiff   = "GIFT" + [char]0x30C7 + [char]0x30FC + [char]0x30BF +
                 "vs" + "GFIX" + [char]0x30C7 + [char]0x30FC + [char]0x30BF          # GIFTデータvsGFIXデータ

$modeCfg = switch ($Mode) {
    'Gift' { @{
        Sheet = $sheetGiftRecv
        Bit   = 1
        Folders = @('excel','GIFT_HM','GIFT_MQ','GIFT_Jenkins','GIFT_noGfixfile')
    } }
    'Gfix' { @{
        Sheet = $sheetGfixRecv
        Bit   = 2
        Folders = @('excel','GFIX_HM','GFIX_Jenkins')
    } }
    'Df'   { @{
        Sheet = $sheetDfDiff
        Bit   = 4
        Folders = @('DF')
    } }
}

# -- Header --------------------------------------------------
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

# ProjectLabels supplies the Japanese NoGfix past-data note from [char] code
# points so this source stays ASCII (no raw Japanese / mojibake on CP932).
$projectLabels = @{}
$labelsPath = Join-Path $PSScriptRoot 'ProjectLabels.ps1'
if (Test-Path -LiteralPath $labelsPath) {
    . $labelsPath
    if (Get-Command -Name 'Get-ProjectLabels' -ErrorAction SilentlyContinue) {
        $projectLabels = Get-ProjectLabels
    }
}

Write-Host ''
Write-Host ("===== Mark ({0}) =====" -f $Mode) -ForegroundColor Green
Write-Host ("  WorkDir   : {0}" -f $WorkDir)
Write-Host ("  Mapping   : {0}" -f $mappingPath)
Write-Host ("  Sheet     : {0}" -f $modeCfg.Sheet)
Write-Host ("  Bit       : {0}" -f $modeCfg.Bit)
Write-Host ("  Folders   : {0}" -f ($modeCfg.Folders -join ', '))
Write-Host ("  NamePrefix: {0}" -f $NamePrefix)
Write-Host ("  Force     : {0}" -f $forceFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }
Write-Host ''

# Box config summary
$configuredFolders = @()
foreach ($f in $modeCfg.Folders) {
    $boxes = $BoxesConfig[$f]
    if ($boxes -and @($boxes).Count -gt 0) {
        $configuredFolders += ("{0}({1})" -f $f, @($boxes).Count)
    }
}
if ($configuredFolders.Count -eq 0 -and $Mode -ne 'Gift') {
    Write-Host '[WARN] No Boxes configured for any folder in this mode.' -ForegroundColor Yellow
    Write-Host '       Edit VerifyConfig.psd1 -> Mark.Boxes after probing with ProbeShapes.' -ForegroundColor DarkGray
    Write-Host '       Nothing to do; exiting.' -ForegroundColor Yellow
    return
}
Write-Host ("  Boxes     : {0}" -f ($configuredFolders -join ', ')) -ForegroundColor DarkGray
if ($Mode -eq 'Gift') { Write-Host ("  NoGfixNoteColumn: {0}" -f $NoGfixNoteColumn) -ForegroundColor DarkGray }

# Image-match summary: how many boxes carry a 'Template' key (opt-in).
$templatedBoxCount = 0
foreach ($f in $modeCfg.Folders) {
    foreach ($b in @($BoxesConfig[$f])) {
        if ($b -is [hashtable] -and $b.ContainsKey('Template') -and -not [string]::IsNullOrWhiteSpace([string]$b.Template)) { $templatedBoxCount++ }
    }
}
if ($templatedBoxCount -gt 0) {
    $tplDirShown = if ([string]::IsNullOrWhiteSpace($TemplateDir)) { (Join-Path $PSScriptRoot 'mark_templates') } else { $TemplateDir }
    Write-Host ("  ImageMatch: {0} box(es) with Template configured (TemplateDir={1}, Tolerance={2})" -f $templatedBoxCount, $tplDirShown, $ImageMatchTolerance) -ForegroundColor DarkGray
    if ([string]::IsNullOrWhiteSpace($locateByImagePath)) {
        Write-Host '  [WARN] Locate-ByImage.ps1 not found -- image-match boxes will fall back to fixed offsets.' -ForegroundColor Yellow
    }
}

# StampImage summary: boxes that insert an image at a Template match instead
# of a rectangle (no match = no stamp; see the box loop below).
$stampBoxCount = 0
foreach ($f in $modeCfg.Folders) {
    foreach ($b in @($BoxesConfig[$f])) {
        if ($b -is [hashtable] -and $b.ContainsKey('StampImage') -and -not [string]::IsNullOrWhiteSpace([string]$b.StampImage)) { $stampBoxCount++ }
    }
}
if ($stampBoxCount -gt 0) {
    Write-Host ("  StampImage: {0} box(es) configured (image-recognition only, no fixed-offset fallback)" -f $stampBoxCount) -ForegroundColor DarkGray
}
if ($Mode -eq 'Gift' -and $NoteStampConfig -and $NoteStampConfig.Count -gt 0) {
    Write-Host ("  NoteStamps: {0}" -f (($NoteStampConfig.Keys | Sort-Object) -join ', ')) -ForegroundColor DarkGray
}
Write-Host ''

if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}
if (-not (Test-Path -LiteralPath $evDir)) {
    Write-Host "[ERROR] evidence dir missing: $evDir" -ForegroundColor Red; exit 1
}

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
Ensure-Column $allRows 'isMarked' '0'

$workRows = @($allRows | Where-Object { Test-TargetRow $_ })
$groups = $workRows | Group-Object Excel_NAME | Sort-Object Name
if ($groups.Count -eq 0) {
    Write-Host '[INFO] No rows after filter.' -ForegroundColor Yellow
    return
}

# -- Main loop -----------------------------------------------
$excel = New-ExcelApp
$cntDone = 0
$cntSkip = 0
$cntFail = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName   = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $excelPrefix = Resolve-ExcelPrefix -Row $first -DefaultPrefix $ExcelPrefix
        $fullStem    = Get-ExcelFullStem -Prefix $excelPrefix -Name $excelName

        $wbPath = Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem
        if ($null -eq $wbPath) {
            Write-Host ("[SKIP] {0}: workbook missing" -f $excelName) -ForegroundColor Yellow
            $cntSkip++; continue
        }

        $curBits = Get-BitValue $first 'isMarked'
        if (-not $forceFlag -and (($curBits -band $modeCfg.Bit) -eq $modeCfg.Bit)) {
            Write-Host ("[SKIP] {0}: bit {1} already set" -f $excelName, $modeCfg.Bit) -ForegroundColor DarkGray
            $cntSkip++; continue
        }

        Write-Host ''
        Write-Host ("----- {0} -----" -f $excelName) -ForegroundColor Cyan

        $wb = $null
        try { $wb = Open-Workbook $excel $wbPath } catch {
            Write-Host ("  [FAIL] open: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $cntFail++; continue
        }

        $allOk = $true
        $marksDrawn = 0
        try {
            $ws = Get-SheetByName $wb $modeCfg.Sheet
            if ($null -eq $ws) {
                Write-Host ("  [FAIL] sheet not found: {0}" -f $modeCfg.Sheet) -ForegroundColor Red
                $allOk = $false
            } else {
                # 1) Wipe previous marks
                $removed = Remove-MarkShapes $ws $NamePrefix
                if ($removed -gt 0) {
                    Write-Host ("  [CLR ] removed {0} existing mark(s)" -f $removed) -ForegroundColor DarkGray
                }

                # 2) Walk shapes, draw marks per metadata
                $shapesToProcess = @()
                foreach ($s in $ws.Shapes) { $shapesToProcess += $s }

                foreach ($s in $shapesToProcess) {
                    $meta = Get-ShapeMetadata $s
                    if ($null -eq $meta) { continue }

                    if ([string]$meta.Key -eq 'verifyNote') {
                        $parts = ([string]$meta.Value) -split '\|'
                        if ($Mode -eq 'Gift' -and $parts.Count -ge 4 -and $parts[0] -eq 'GIFT_noGfixfile') {
                            try {
                                $cid = [string]$parts[1]
                                $xywh = @($parts[2] -split ',' | ForEach-Object { [double]$_ })
                                $imageWidth = [double]$parts[3]
                                if ($imageWidth -le 0) { $imageWidth = [double]$s.Width }
                                $scale = ([double]$s.Width) / $imageWidth
                                $left = ([double]$s.Left) + ($xywh[0] * $scale)
                                $top  = ([double]$s.Top)  + ($xywh[1] * $scale)
                                $bw   = $xywh[2] * $scale
                                $bh   = $xywh[3] * $scale
                                $name = ("{0}verifyNote_{1}_0" -f $NamePrefix, $cid)
                                Add-RedRectangle $ws $left $top $bw $bh $name $LineWeight | Out-Null
                                $marksDrawn++

                                $noteCol = [int]$ws.Range(("{0}1" -f $NoGfixNoteColumn)).Column
                                $noteRow = [int]$s.TopLeftCell.Row
                                $pastData = [string]$projectLabels['NoGfixPastData']
                                if ([string]::IsNullOrEmpty($pastData)) {
                                    # Fallback if ProjectLabels was unavailable.
                                    $pastData = [char]0x904E + [char]0x53BB + [char]0x5206 + [char]0x30C7 + [char]0x30FC + [char]0x30BF + [char]0x30FC
                                }
                                $ws.Cells.Item($noteRow, $noteCol).Value2 = $pastData
                                Write-Host ("  [NOTE] GIFT_noGfixfile {0,-12} L={1,6:0.0} T={2,6:0.0} W={3,5:0.0} H={4,5:0.0} {5}{6}" -f $cid, $left, $top, $bw, $bh, $NoGfixNoteColumn, $noteRow) -ForegroundColor Green

                                # Optional stamp image (e.g. already_exists.png), keyed by the
                                # note's Folder ('GIFT_noGfixfile') in -NoteStampConfig. Reuses
                                # the same scaled $top this block already computed for the red
                                # rectangle -- no re-scan of the source PNG needed. Best-effort:
                                # a stamp failure only warns, it never fails the verifyNote mark.
                                if ($NoteStampConfig -and $NoteStampConfig.ContainsKey($parts[0])) {
                                    try {
                                        $stampCfg = $NoteStampConfig[$parts[0]]
                                        $stampImage = [string]$stampCfg.Image
                                        $stampTplPath = Resolve-MarkTemplatePath -Template $stampImage -TemplateDir $TemplateDir
                                        if ($null -eq $stampTplPath) {
                                            Write-Host ("  [WARN] noteStamp image not found: {0}" -f $stampImage) -ForegroundColor Yellow
                                        } else {
                                            $stampCol = 'AF'
                                            if (-not [string]::IsNullOrWhiteSpace([string]$stampCfg.Column)) { $stampCol = [string]$stampCfg.Column }
                                            $stampRowOffset = 0
                                            try { $stampRowOffset = [int]$stampCfg.RowOffset } catch {}

                                            # $top is the row's pixel-scaled sheet Top; Get-RowAtOrBelow
                                            # returns the row AFTER it (Top >= target), same off-by-one
                                            # Get-PictureBottomRow corrects for -- so -1 lands ON the
                                            # highlighted row itself (RowOffset=0 means "same row").
                                            $stampStartScan = 1
                                            try { $stampStartScan = [Math]::Max(1, [int]([Math]::Floor($top / 15.0))) } catch {}
                                            $rowAfterTop = Get-RowAtOrBelow $ws $top $stampStartScan 0
                                            $highlightRow = [Math]::Max(1, $rowAfterTop - 1)
                                            $stampRow = [Math]::Max(1, $highlightRow + $stampRowOffset)
                                            $stampColIdx = [int]$ws.Range(("{0}1" -f $stampCol)).Column

                                            $stampPic = Insert-PictureBringToFront $ws $stampRow $stampColIdx $stampTplPath
                                            try { $stampPic.Name = ("{0}verifyNoteStamp_{1}_0" -f $NamePrefix, $cid) } catch {}
                                            $marksDrawn++
                                            Write-Host ("  [STAMP] GIFT_noGfixfile {0,-12} {1}{2}" -f $cid, $stampCol, $stampRow) -ForegroundColor Green
                                        }
                                    } catch {
                                        Write-Host ("  [WARN] noteStamp failed for {0}: {1}" -f $cid, $_.Exception.Message) -ForegroundColor Yellow
                                    }
                                }
                            } catch {
                                Write-Host ("  [FAIL] verifyNote: {0}" -f $_.Exception.Message) -ForegroundColor Red
                                $allOk = $false
                            }
                        }
                        continue
                    }

                    $folder = [string]$meta.Key
                    $cid    = [string]$meta.Value
                    if ($modeCfg.Folders -notcontains $folder) { continue }

                    $boxes = @($BoxesConfig[$folder])
                    if ($boxes.Count -eq 0) { continue }

                    $picLeft = [double]$s.Left
                    $picTop  = [double]$s.Top

                    $idx = 0
                    foreach ($b in $boxes) {
                        $lw = $LineWeight
                        if ($b.ContainsKey('LineWeight')) {
                            try { $lw = [double]$b.LineWeight } catch {}
                        }

                        # Stamp-image boxes (opt-in via 'StampImage'): pure image-
                        # recognition, no fixed-offset fallback. Requires 'Template'
                        # on the same box; when the template is found on the source
                        # snap PNG, StampImage is inserted (native size) at the
                        # matched+scaled location instead of a red rectangle. No
                        # match = nothing to stamp -- this box's whole point is
                        # "only appear when the target pattern is actually found"
                        # (e.g. GIFT_noGfixfile: a visible past-data hit vs. the
                        # normal no-file-found case), so there is deliberately no
                        # OffsetX/OffsetY fallback here.
                        if ($b.ContainsKey('StampImage')) {
                            $stampImageName = [string]$b.StampImage
                            $imgHit = Find-MarkBoxByImage -Box $b -Folder $folder -Cid $cid -BoxIndex $idx -WorkDir $WorkDir `
                                -TemplateDir $TemplateDir -DefaultTolerance $ImageMatchTolerance `
                                -LocateScript $locateByImagePath -Shape $s
                            if ($null -eq $imgHit) {
                                Write-Host ("  [SKIP-STAMP] {0,-16} {1,-12} [{2,3}] no Template match -- nothing to stamp" -f $folder, $cid, $idx) -ForegroundColor DarkGray
                                $idx++
                                continue
                            }
                            $stampPath = Resolve-MarkTemplatePath -Template $stampImageName -TemplateDir $TemplateDir
                            if ($null -eq $stampPath) {
                                Write-Host ("  [WARN] StampImage not found: {0}" -f $stampImageName) -ForegroundColor Yellow
                                $idx++
                                continue
                            }
                            $name = ("{0}{1}_{2}_{3}" -f $NamePrefix, $folder, $cid, $idx)
                            try {
                                $stampPic = Insert-PictureAtPointBringToFront $ws $imgHit.Left $imgHit.Top $stampPath
                                try { $stampPic.Name = $name } catch {}
                                $marksDrawn++
                                Write-Host ("  [STAMP-IMG] {0,-16} {1,-12} [{2,3}] L={3,6:0.0} T={4,6:0.0} ({5})" -f $folder, $cid, $idx, $imgHit.Left, $imgHit.Top, $imgHit.Source) -ForegroundColor Green
                            } catch {
                                Write-Host ("  [FAIL] StampImage {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
                                $allOk = $false
                            }
                            $idx++
                            continue
                        }

                        $left = 0.0; $top = 0.0; $bw = 100.0; $bh = 20.0
                        $imgHit = $null
                        if ($b.ContainsKey('CellCols')) {
                            # Cell-range positioning: place rect relative to sheet
                            # columns/rows rather than the picture's pixel corner.
                            $rowsFromBot = 2
                            if ($b.ContainsKey('RowsFromBottom')) {
                                try { $rowsFromBot = [int]$b.RowsFromBottom } catch {}
                            }
                            $bottomRow = Get-PictureBottomRow $ws $s
                            $topRow    = [Math]::Max(1, $bottomRow - $rowsFromBot + 1)
                            $rect = Get-CellRangeRect $ws ([string]$b.CellCols) $topRow $bottomRow
                            $left = $rect.Left
                            $top  = $rect.Top
                            $bw   = $rect.Width
                            $bh   = $rect.Height
                        } else {
                            # Image-recognition placement (opt-in via the box's
                            # 'Template' key): locate the actual mark target on
                            # the source snap PNG instead of trusting a fixed
                            # offset. Falls back to OffsetX/OffsetY/Width/Height
                            # below when no Template is configured or no match
                            # is found, so this degrades gracefully.
                            $imgHit = Find-MarkBoxByImage -Box $b -Folder $folder -Cid $cid -BoxIndex $idx -WorkDir $WorkDir `
                                -TemplateDir $TemplateDir -DefaultTolerance $ImageMatchTolerance `
                                -LocateScript $locateByImagePath -Shape $s
                            if ($null -ne $imgHit) {
                                $left = $imgHit.Left; $top = $imgHit.Top; $bw = $imgHit.Width; $bh = $imgHit.Height
                            } else {
                                $ox = 0.0; $oy = 0.0
                                try { $ox = [double]$b.OffsetX } catch {}
                                try { $oy = [double]$b.OffsetY } catch {}
                                try { $bw = [double]$b.Width } catch {}
                                try { $bh = [double]$b.Height } catch {}

                                # Row-position adjustment (opt-in via 'RowHeight' > 0):
                                # some pages (GIFT_MQ) repeat a variable number of
                                # records and the target is always the LAST/newest
                                # one for this correl, which is not always the box's
                                # calibrated 'BaseRow' (default 2, the common
                                # 2-record case). Shift OffsetY by the row delta;
                                # RowHeight = 0 (default) keeps the legacy fixed
                                # offset untouched.
                                $rowHeight = 0.0
                                try { $rowHeight = [double]$b.RowHeight } catch {}
                                if ($rowHeight -gt 0) {
                                    $baseRow = 2
                                    if ($b.ContainsKey('BaseRow')) { try { $baseRow = [int]$b.BaseRow } catch {} }
                                    $rowInfo = Get-MarkMqRowInfo -WorkDir $WorkDir -Folder $folder -Cid $cid
                                    if ($null -ne $rowInfo) {
                                        $dY = ($rowInfo.RowIndex - $baseRow) * $rowHeight
                                        $oy += $dY
                                        Write-Host ("  [ROW ] {0,-16} {1,-12} [{2,3}] row {3}/{4} ({5}) dY={6:0.0}" -f $folder, $cid, $idx, $rowInfo.RowIndex, $rowInfo.NumRecords, $rowInfo.Source, $dY) -ForegroundColor DarkCyan
                                    } else {
                                        Write-Host ("  [WARN] {0,-16} {1,-12} [{2,3}] row info unavailable; using BaseRow {3} offset as-is" -f $folder, $cid, $idx, $baseRow) -ForegroundColor Yellow
                                    }
                                }

                                $left = $picLeft + $ox
                                $top  = $picTop  + $oy
                            }
                        }

                        $name = ("{0}{1}_{2}_{3}" -f $NamePrefix, $folder, $cid, $idx)
                        $tag  = if ($null -ne $imgHit) { 'MARK-IMG' } else { 'MARK' }
                        $tagSuffix = if ($null -ne $imgHit) { " ({0})" -f $imgHit.Source } else { '' }

                        try {
                            Add-RedRectangle $ws $left $top $bw $bh $name $lw | Out-Null
                            $marksDrawn++
                            Write-Host ("  [{0}] {1,-16} {2,-12} [{3,3}] L={4,6:0.0} T={5,6:0.0} W={6,5:0.0} H={7,5:0.0}{8}" -f $tag, $folder, $cid, $idx, $left, $top, $bw, $bh, $tagSuffix) -ForegroundColor Green
                        } catch {
                            Write-Host ("  [FAIL] AddShape {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
                            $allOk = $false
                        }
                        $idx++
                    }
                }

                # GFIX log yellow highlight, folded in from the old MarkGfixLog
                # phase. Best-effort: missing anchors warn but never block the
                # isMarked bit -- the red rectangles are the gating evidence.
                if ($Mode -eq 'Gfix') {
                    $hl = Invoke-GfixLogHighlight -ws $ws -LogAnchor $GfixLogAnchor `
                        -CommandPattern $GfixLogCommandPattern -HighlightColor $GfixLogHighlightColor `
                        -ColStart $GfixLogColStart -ColEnd $GfixLogColEnd `
                        -AutoWidth $GfixLogAutoWidth -PadCols $GfixLogPadCols `
                        -FontName $GfixLogFontName -FontSize $GfixLogFontSize
                    foreach ($w in @($hl.Warnings)) { Write-Host ("  [GfixLog WARN] {0}" -f $w) -ForegroundColor Yellow }
                    Write-Host ("  [GfixLog] highlights applied: {0} (anchors: {1}, AutoWidth={2})" -f $hl.Applied, $hl.Anchors, $GfixLogAutoWidth) -ForegroundColor DarkGray
                }
            }

            $wb.Save()
        } catch {
            Write-Host ("  [FAIL] processing: {0}" -f $_.Exception.Message) -ForegroundColor Red
            $allOk = $false
        } finally {
            Close-Workbook $wb $false
        }

        Write-Host ("  marks drawn: {0}" -f $marksDrawn) -ForegroundColor DarkGray

        if ($allOk -and $marksDrawn -gt 0) {
            $groupNames = @($g.Group | ForEach-Object { [string]$_.Correl_ID_M })
            foreach ($r in $allRows) {
                if ($groupNames -contains [string]$r.Correl_ID_M) {
                    Set-BitValue $r 'isMarked' $modeCfg.Bit
                }
            }
            Write-Host ("  isMarked |= {0} for {1} row(s)" -f $modeCfg.Bit, $g.Count) -ForegroundColor Green
            $cntDone++
        } elseif ($marksDrawn -eq 0) {
            Write-Host '  [WARN] no marks drawn (no matching shapes with metadata?)' -ForegroundColor Yellow
            $cntFail++
        } else {
            Write-Host '  isMarked NOT updated (allOk=false)' -ForegroundColor Yellow
            $cntFail++
        }
    }

    if ($cntDone -gt 0) {
        $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
        Write-Host ''
        Write-Host ("Mapping saved: {0}" -f $mappingPath) -ForegroundColor DarkGreen
    }
} finally {
    Close-ExcelApp $excel
}

Write-Host ''
Write-Host ("===== Mark ({0}) Done =====" -f $Mode) -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
