# ============================================================
# SendVsGift.ps1 - compare-prep review for SEND data vs GIFT data
#
# Stage 1 (manual review):
#   1. Gather exact metadata for every file under <WorkDir>\DATA\GIFT.
#   2. Save it to <WorkDir>\data\gift_metadata.csv.
#   3. Ensure mapping column SendVsGift exists.
#   4. Pending rows are grouped per evidence workbook: each workbook is
#      opened ONCE, and for every correl row the cursor is moved to the
#      Correl_ID_S label cell in column A of the send-data sheet (the
#      pictures sit right below it). After each console answer Excel is
#      brought back to the foreground so the next correl can be checked
#      without manual Alt+Tab into Excel.
#   5. Enter marks SendVsGift=1; n marks SendVsGift=2 (NG, needs follow-up);
#      s skips; q quits. The workbook is saved/closed when its last correl
#      row is done.
#
# Stage 2 (-Ocr):
#   See docs/SendVsGift.md. For each correl row only the pictures between
#   its column-A label and the next label are exported
#   (EvidenceImageExport.ps1 - Ctrl+G groups are flattened to child
#   pictures), OCR'd with the built-in Windows engine (OcrWindows.ps1),
#   and judged by the pure rules in SendMetadata.ps1:
#     gift 0 bytes  -> 'used CYLINDERS : 0' screen, or begin+end-of-data
#                      markers on one image with no 000001 line
#     gift has data -> zero-padded max row number must appear, and the
#                      first/last records must match by first token
#                      (exact) or >=80% prefix similarity (OCR noise)
#   Verdict ok -> auto-mark SendVsGift=1; ng -> auto-mark SendVsGift=2 and
#   report at the end; unknown -> fall back to the manual prompt.
#   -NoAutoMark keeps the verdict advisory-only (always prompt).
#   send_metadata.csv keeps the parsed OCR record per correl for audit.
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
    [switch]$Ocr,
    [switch]$NoAutoMark,
    [string]$OcrLanguage = 'ja',
    [string]$SendSheetName = '',
    [string]$ZeroBytePattern = '',
    [string]$ZeroTemplate = '',
    [string]$ExcelHelpersScript = ''
)

$ErrorActionPreference = 'Stop'

# capture switches BEFORE dot-sourcing (see CLAUDE.md dot-source safety rule)
$ocrFlag      = [bool]$Ocr.IsPresent
$forceFlag    = [bool]$Force.IsPresent
$dryRunFlag   = [bool]$DryRun.IsPresent
$maximizeFlag = [bool]$Maximize.IsPresent
$autoMark     = -not [bool]$NoAutoMark.IsPresent

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
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')
. (Join-Path $PSScriptRoot 'SendMetadata.ps1')
. (Join-Path $PSScriptRoot 'OcrWindows.ps1')
. (Join-Path $PSScriptRoot 'EvidenceImageExport.ps1')

# Foreground helper: after the operator answers in the console, hand the
# focus back to Excel so the next correl is immediately visible. Direct
# SetForegroundWindow on the Excel hwnd is more reliable than a blind
# Alt+Tab; SendKeys %{TAB} stays as the fallback.
try {
    Add-Type -Namespace VerifySvg -Name Win32 -MemberDefinition @'
[DllImport("user32.dll")]
public static extern bool SetForegroundWindow(IntPtr hWnd);
'@ -ErrorAction Stop
} catch {}

function Set-ExcelForeground($Excel) {
    try {
        if (([type]'VerifySvg.Win32') -and [VerifySvg.Win32]::SetForegroundWindow([intptr]$Excel.Hwnd)) { return }
    } catch {}
    try {
        $sh = New-Object -ComObject WScript.Shell
        $sh.SendKeys('%{TAB}')
    } catch {}
}

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

# Finds the worksheet carrying the send-data evidence pictures.
function Find-WorksheetByName($Workbook, [string]$SheetName) {
    foreach ($s in $Workbook.Worksheets) {
        if ([string]$s.Name -eq $SheetName) { return $s }
    }
    return $null
}

# Finds the Correl_ID_S label cell in column A of the send sheet
# (whole-cell match first, then substring for decorated labels).
function Find-SendCorrelCell($Worksheet, [string]$Sid) {
    if ([string]::IsNullOrWhiteSpace($Sid)) { return $null }
    $missing = [System.Reflection.Missing]::Value
    $rng = $Worksheet.Range('A:A')
    foreach ($lookAt in @(1, 2)) {   # xlWhole, then xlPart
        $cell = $null
        try { $cell = $rng.Find($Sid, $missing, -4163, $lookAt) } catch { $cell = $null }
        if ($null -ne $cell) { return $cell }
    }
    return $null
}

# Vertical bounds of one correl section on the send sheet: from the label
# cell down to the next non-empty cell in column A (the next correl's
# label), or unbounded when the label is the last one.
function Get-SendSectionBounds($Worksheet, $LabelCell) {
    $top = 0.0
    try { $top = [double]$LabelCell.Top } catch {}
    $bottom = -1.0
    try {
        $r = [int]$LabelCell.Row
        $below = $Worksheet.Cells.Item($r + 1, 1)
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

function Show-MetadataBlock($Row, [array]$GiftRows) {
    Write-Host ''
    Write-Host ("ID: {0}  Excel: {1}  Job: {2}" -f $Row.Correl_ID_S, $Row.Excel_NAME, $Row.JOB_NAME) -ForegroundColor Cyan
    if ($GiftRows.Count -eq 0) {
        Write-Host '  [WARN] no matching GIFT data file found.' -ForegroundColor Yellow
        return
    }
    foreach ($m in $GiftRows) {
        Write-Host ("  File        : {0}" -f $m.FullName)
        Write-Host ("  Size        : {0}" -f $m.SizeDisplay)
        Write-Host ("  Max row num : {0}" -f $m.MaxRowNumber)
        Write-Host ("  Length      : max={0}, min={1}, first={2}, last={3}" -f $m.MaxRecordLength, $m.MinRecordLength, $m.FirstRecordLength, $m.LastRecordLength)
        Write-Host ("  First row   : {0}" -f $m.FirstRecordToken)
        Write-Host ("  Last row    : {0}" -f $m.LastRecordToken)
    }
}

# Stage 2: export only THIS correl's pictures (between its column-A label
# and the next), OCR them and judge with the pure SendMetadata rules.
# Returns @{ Meta = <send_metadata record or $null>; Verdict = ok|ng|unknown|none }.
function Invoke-SendOcrReview {
    param($Workbook, $Worksheet, $LabelCell, $Row, [array]$GiftRows,
          [string]$SheetName, [string]$ImagesRoot, [string]$LanguageTag, [string]$ZeroPattern,
          [string]$ZeroTemplatePath = '')
    $sid = [string]$Row.Correl_ID_S
    if ([string]::IsNullOrWhiteSpace($sid)) { return @{ Meta = $null; Verdict = 'none' } }
    if ($GiftRows.Count -eq 0) {
        Write-Host '  [OCR] no matching gift metadata row; manual check only.' -ForegroundColor Yellow
        return @{ Meta = $null; Verdict = 'none' }
    }
    if ($null -eq $Worksheet -or $null -eq $LabelCell) {
        Write-Host '  [OCR] correl label not located on the send sheet; manual check only.' -ForegroundColor Yellow
        return @{ Meta = $null; Verdict = 'none' }
    }

    $bounds = Get-SendSectionBounds $Worksheet $LabelCell
    $outDir = Join-Path $ImagesRoot $sid
    # pipe through Where-Object: flattens the returned array and drops any
    # empty entry so Invoke-WinOcrFile never sees an empty -Path (which
    # would die on Test-Path -LiteralPath with a cryptic binding error)
    $pngs = @(Export-SheetPicturesToPng $Workbook $SheetName $outDir $sid $bounds.Top $bounds.Bottom |
        Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($pngs.Count -eq 0) {
        Write-Host ("  [OCR] no pictures in the {0} section of sheet '{1}'; manual check only." -f $sid, $SheetName) -ForegroundColor Yellow
        return @{ Meta = $null; Verdict = 'none' }
    }

    # dual-engine pass: the ja recognizer garbles digit runs on the host
    # terminal font ('00' -> a kanji box, '7' -> '?'); en-US reads the same
    # digits cleanly, so each image is OCR'd with both and the row lines
    # are merged (matchers scan all of them; extra lines are harmless).
    $ocrLangs = @($LanguageTag)
    try {
        if (((Get-WinOcrLanguageTags) -contains 'en-US') -and ($LanguageTag -ne 'en-US')) { $ocrLangs += 'en-US' }
    } catch {}

    $imageSets = @()
    $allLines = @()
    foreach ($p in $pngs) {
        $lines = @()
        foreach ($lg in $ocrLangs) {
            $res = Invoke-WinOcrFile -Path $p -LanguageTag $lg
            # row reconstruction from word boxes: the engine fragments one
            # terminal row into several OCR lines, separating label and record
            $lines += @(ConvertTo-SendRowLines $res.Lines)
        }
        $imageSets += ,@($lines)
        $allLines += $lines
    }
    # audit/debug dump of what the matcher actually saw
    try {
        $dumpSb = New-Object System.Text.StringBuilder
        for ($di = 0; $di -lt $imageSets.Count; $di++) {
            [void]$dumpSb.AppendLine(('===== image {0:D2} =====' -f ($di + 1)))
            foreach ($dl in @($imageSets[$di])) { [void]$dumpSb.AppendLine([string]$dl) }
        }
        $dumpPath = Join-Path $outDir ($sid + '_ocr.txt')
        [System.IO.File]::WriteAllText($dumpPath, $dumpSb.ToString(), (New-Object System.Text.UTF8Encoding($false)))
    } catch {}

    if ($allLines.Count -eq 0) {
        Write-Host ("  [OCR] no text recognized on any of the {0} exported image(s) - check {1}" -f $pngs.Count, $outDir) -ForegroundColor Yellow
        Write-Host ("        diagnose with: .\OcrTool.ps1 -Diag -Path {0}" -f $outDir) -ForegroundColor Yellow
    }
    $meta = Build-SendMetadataRecord -CorrelIdS $sid -ExcelName ([string]$Row.Excel_NAME) `
        -ImageCount $pngs.Count -TextLines $allLines -ZeroBytePattern $ZeroPattern
    $cmp = Compare-SendGiftEvidence -GiftRow $GiftRows[0] -ImageTextSets $imageSets -ZeroBytePattern $ZeroPattern
    Write-Host ("  [OCR] images={0} lines={1} (engines: {2})" -f $pngs.Count, $allLines.Count, ($ocrLangs -join '+')) -ForegroundColor Cyan
    foreach ($c in $cmp.Checks) {
        $color = switch ($c.Status) { 'match' { 'Green' } 'fuzzy' { 'DarkGreen' } 'mismatch' { 'Red' } default { 'DarkGray' } }
        Write-Host ("    {0,-14} send='{1}' gift='{2}' -> {3}" -f $c.Name, $c.Send, $c.Gift, $c.Status) -ForegroundColor $color
    }
    $verdict = [string]$cmp.Verdict

    # Zero-byte template fallback: OCR keeps missing the small ': 0' value
    # on the dataset-info screen, so an operator-cropped template (cut from
    # one of THESE exported PNGs - same pipeline, same pixels) can prove
    # the 0-byte evidence via Locate-ByImage pixel matching instead.
    $giftZero = $false
    try { $giftZero = ([long]$GiftRows[0].SizeBytes -eq 0) } catch {}
    if ($giftZero -and $verdict -eq 'unknown' -and
        -not [string]::IsNullOrWhiteSpace($ZeroTemplatePath) -and (Test-Path -LiteralPath $ZeroTemplatePath)) {
        $locator = Join-Path $PSScriptRoot 'Locate-ByImage.ps1'
        $tplSeen = $false
        foreach ($p in $pngs) {
            try {
                $box = & $locator -SourcePath $p -TemplatePath $ZeroTemplatePath -Quiet
                if ($null -ne $box) { $tplSeen = $true; break }
            } catch {}
        }
        if ($tplSeen) {
            Write-Host ("    {0,-14} send='template {1} FOUND' gift='0 bytes' -> match" -f 'ZeroByteTpl', (Split-Path -Leaf $ZeroTemplatePath)) -ForegroundColor Green
            $verdict = 'ok'
        } else {
            Write-Host ("    {0,-14} send='template {1} not found' gift='0 bytes' -> unknown" -f 'ZeroByteTpl', (Split-Path -Leaf $ZeroTemplatePath)) -ForegroundColor DarkGray
        }
    }

    $vColor = switch ($verdict) { 'ok' { 'Green' } 'ng' { 'Red' } default { 'Yellow' } }
    Write-Host ("  [OCR] verdict: {0}" -f $verdict) -ForegroundColor $vColor
    return @{ Meta = $meta; Verdict = $verdict }
}

# ---- WorkDir / Owner resolution (standalone launch friendly) ----
# VerifyTool.ps1 always passes both; when run directly we fall back to
# verify_session.json, then to the single mapping_*.csv in the work
# folder, then to a prompt - no more silent 'mapping_.csv' failure.
$sessionData = @{}
try {
    $sessionFile = Join-Path $PSScriptRoot 'verify_session.json'
    if (Test-Path -LiteralPath $sessionFile) {
        $obj = Get-Content -LiteralPath $sessionFile -Raw -Encoding UTF8 | ConvertFrom-Json
        foreach ($p in $obj.PSObject.Properties) { $sessionData[$p.Name] = $p.Value }
    }
} catch {}

if ([string]::IsNullOrWhiteSpace($WorkDir)) {
    $candidate = ''
    if ($sessionData.ContainsKey('WorkDir')) { $candidate = [string]$sessionData['WorkDir'] }
    if (-not [string]::IsNullOrWhiteSpace($candidate) -and (Test-Path -LiteralPath $candidate)) {
        $v = Read-Host ("WorkDir path [{0}]" -f $candidate)
        $WorkDir = if ([string]::IsNullOrWhiteSpace($v)) { $candidate } else { $v }
    } else {
        $WorkDir = Read-Host 'WorkDir path'
    }
}
if (-not (Test-Path -LiteralPath $WorkDir)) { throw "WorkDir not found: $WorkDir" }

if ([string]::IsNullOrWhiteSpace($Owner)) {
    $sessionOwner = ''
    if ($sessionData.ContainsKey('Owner')) { $sessionOwner = [string]$sessionData['Owner'] }
    if (-not [string]::IsNullOrWhiteSpace($sessionOwner) -and
        (Test-Path -LiteralPath (Join-Path $WorkDir ("mapping_{0}.csv" -f $sessionOwner)))) {
        $Owner = $sessionOwner
        Write-Host ("[INFO] Owner from session: {0}" -f $Owner) -ForegroundColor DarkGray
    } else {
        $maps = @(Get-ChildItem -LiteralPath $WorkDir -Filter 'mapping_*.csv' -File -ErrorAction SilentlyContinue)
        if ($maps.Count -eq 1 -and $maps[0].Name -match '^mapping_(.+)\.csv$') {
            $Owner = $Matches[1]
            Write-Host ("[INFO] Owner from work folder: {0} ({1})" -f $Owner, $maps[0].Name) -ForegroundColor DarkGray
        } else {
            $Owner = Read-Host 'Owner suffix (mapping_<Owner>.csv)'
        }
    }
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

# zero-byte template: relative paths resolve against the work folder
if (-not [string]::IsNullOrWhiteSpace($ZeroTemplate) -and -not [System.IO.Path]::IsPathRooted($ZeroTemplate)) {
    $ZeroTemplate = Join-Path $WorkDir $ZeroTemplate
}

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

function Set-SendVsGiftValue([string]$Sid, [string]$Value) {
    foreach ($row in $script:allRows) {
        if ([string]$row.Correl_ID_S -eq $Sid) { $row.SendVsGift = $Value }
    }
    $script:allRows | Export-Csv -LiteralPath $script:mappingPath -Encoding UTF8 -NoTypeInformation -Force
}

# pending = anything not yet OK (covers '', '0' and NG rows marked '2')
$selectedRows = @($allRows | Where-Object { Test-TargetRow $_ $targetSet })
$pendingRows = @($selectedRows | Where-Object { $forceFlag -or [string]$_.SendVsGift -ne '1' })
if ($pendingRows.Count -eq 0) {
    Write-Host '[INFO] no pending SendVsGift rows.' -ForegroundColor Green
    return
}

Write-Host ("[INFO] pending SendVsGift rows: {0}" -f $pendingRows.Count) -ForegroundColor Cyan
if ($dryRunFlag) { return }

# Stage 2 OCR setup (no-op unless -Ocr): availability probe + send-side
# metadata store. OCR runs on Windows only; elsewhere we warn and the
# manual flow continues untouched.
$ocrReady = $false
$sendMetaPath = Join-Path (Join-Path $WorkDir 'data') 'send_metadata.csv'
$sendImagesRoot = Join-Path (Join-Path $WorkDir 'data') 'send_images'
$sendMetaRows = @()
if ([string]::IsNullOrWhiteSpace($SendSheetName)) {
    $labels = Get-ProjectLabels
    $SendSheetName = [string]$labels['SheetSoshinData']
}
if ($ocrFlag) {
    $ocrReady = Test-WinOcrAvailable
    if ($ocrReady) {
        if (Test-Path -LiteralPath $sendMetaPath) {
            $sendMetaRows = @(Import-Csv -LiteralPath $sendMetaPath -Encoding UTF8)
        }
        Write-Host ("[OCR] engine ready (languages: {0}); send sheet: {1}; auto-mark: {2}" -f `
            ((Get-WinOcrLanguageTags) -join ', '), $SendSheetName, $autoMark) -ForegroundColor Green
    } else {
        Write-Host ("[WARN] -Ocr requested but Windows OCR is unavailable: {0}" -f (Get-WinOcrInitError)) -ForegroundColor Yellow
        Write-Host '       continuing with the manual compare flow only.' -ForegroundColor Yellow
    }
}
# OCR mode runs hands-off: never steal the OS foreground for Excel and
# never save the workbooks on close (content is untouched by this phase)
$quietOcr = ($ocrFlag -and $ocrReady)

# Group pending rows per evidence workbook so each file is opened once
# and the cursor just moves between correl labels inside it.
$groupList = @()
$groupMap = @{}
foreach ($r in $pendingRows) {
    $prefix = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
    $fullStem = Get-ExcelFullStem -Prefix $prefix -Name ([string]$r.Excel_NAME)
    $file = Get-EvidencePath $EvidenceDir $fullStem
    if (-not $groupMap.ContainsKey($file)) {
        $g = @{ File = $file; ExcelName = [string]$r.Excel_NAME; Rows = @() }
        $groupMap[$file] = $g
        $groupList += $g
    }
    $groupMap[$file].Rows += $r
}

$ngList = @()
$quitAll = $false
$rowIdx = 0
$excel = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true
    $excel.DisplayAlerts = $false
    try { if ($maximizeFlag) { $excel.WindowState = -4137 } } catch {}

    foreach ($g in $groupList) {
        if ($quitAll) { break }
        if (-not (Test-Path -LiteralPath $g.File)) {
            $rowIdx += @($g.Rows).Count
            Write-Host ("[MISSING] workbook for {0} ({1} row(s)): {2}" -f $g.ExcelName, @($g.Rows).Count, $g.File) -ForegroundColor Yellow
            continue
        }

        $wb = $null
        try {
            Write-Host ''
            Write-Host ("[OPEN] {0}  ({1} correl row(s))" -f $g.File, @($g.Rows).Count) -ForegroundColor Cyan
            $wb = $excel.Workbooks.Open($g.File, 0, $false)
            $sendWs = Find-WorksheetByName $wb $SendSheetName
            if ($null -eq $sendWs) {
                Write-Host ("  [WARN] send sheet '{0}' not found; cursor stays at {1}." -f $SendSheetName, $CursorCell) -ForegroundColor Yellow
            }
            $markedAny = $false

            foreach ($r in $g.Rows) {
                $rowIdx++
                $sid = [string]$r.Correl_ID_S
                $giftRows = @(Find-GiftMetadataForRow $r $metadataRows)
                Show-MetadataBlock $r $giftRows
                Write-Host ("[{0}/{1}] {2}" -f $rowIdx, $pendingRows.Count, $sid) -ForegroundColor Cyan

                # cursor to this correl's label cell in column A (review rule:
                # the evidence pictures sit right below the label)
                $labelCell = $null
                if ($null -ne $sendWs) {
                    try { $sendWs.Activate() | Out-Null } catch {}
                    $labelCell = Find-SendCorrelCell $sendWs $sid
                    if ($null -ne $labelCell) {
                        try { $excel.Goto($labelCell, $true) | Out-Null } catch {
                            try { $labelCell.Select() | Out-Null } catch {}
                        }
                        Write-Host ("  cursor -> {0}!A{1}" -f $SendSheetName, [int]$labelCell.Row) -ForegroundColor DarkGray
                    } else {
                        Write-Host ("  [WARN] '{0}' not found in column A of sheet '{1}'." -f $sid, $SendSheetName) -ForegroundColor Yellow
                        try { $sendWs.Range($CursorCell).Select() | Out-Null } catch {}
                    }
                }
                # OCR runs hands-off: do not steal the foreground from the
                # console; the operator alt-tabs to Excel only when needed
                if (-not $quietOcr) { Set-ExcelForeground $excel }

                $verdict = 'none'
                if ($ocrFlag -and $ocrReady) {
                    try {
                        $res = Invoke-SendOcrReview $wb $sendWs $labelCell $r $giftRows $SendSheetName $sendImagesRoot $OcrLanguage $ZeroBytePattern $ZeroTemplate
                        if ($null -ne $res.Meta) {
                            $sendMetaRows = @($sendMetaRows | Where-Object { [string]$_.CorrelIdS -ne [string]$res.Meta.CorrelIdS }) + @($res.Meta)
                            $sendMetaRows | Export-Csv -LiteralPath $sendMetaPath -Encoding UTF8 -NoTypeInformation -Force
                        }
                        $verdict = [string]$res.Verdict
                    } catch {
                        Write-Host ("  [WARN] OCR compare failed; manual check continues: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
                        foreach ($frame in @([string]$_.ScriptStackTrace -split "`r?`n")) {
                            if (-not [string]::IsNullOrWhiteSpace($frame)) {
                                Write-Host ("         {0}" -f $frame.Trim()) -ForegroundColor DarkGray
                            }
                        }
                    }
                }

                if ($autoMark -and $verdict -eq 'ok') {
                    Set-SendVsGiftValue $sid '1'
                    $markedAny = $true
                    Write-Host '  [DONE] SendVsGift=1 (OCR auto)' -ForegroundColor Green
                    continue
                }
                if ($autoMark -and $verdict -eq 'ng') {
                    Set-SendVsGiftValue $sid '2'
                    $markedAny = $true
                    $ngList += [pscustomobject]@{ CorrelIdS = $sid; ExcelName = [string]$r.Excel_NAME; File = $g.File }
                    Write-Host '  [NG] SendVsGift=2 (OCR mismatch) - check this one manually.' -ForegroundColor Red
                    continue
                }
                if ($verdict -eq 'unknown') {
                    Write-Host '  [OCR] verdict unknown - your call.' -ForegroundColor Yellow
                }

                $ans = [string](Read-Host 'Check the workbook and console metadata. Enter=mark SendVsGift=1, n=mark 2(NG), s=skip, q=quit')
                $key = $ans.Trim().ToLower()
                if ($key -eq 'q') { $quitAll = $true; break }
                if ($key -eq 's') { Write-Host '  [SKIP]' -ForegroundColor Yellow; continue }
                if ($key -eq 'n') {
                    Set-SendVsGiftValue $sid '2'
                    $markedAny = $true
                    $ngList += [pscustomobject]@{ CorrelIdS = $sid; ExcelName = [string]$r.Excel_NAME; File = $g.File }
                    Write-Host '  [NG] SendVsGift=2' -ForegroundColor Red
                    continue
                }
                Set-SendVsGiftValue $sid '1'
                $markedAny = $true
                Write-Host '  [DONE] SendVsGift=1' -ForegroundColor Green
            }

            if ($markedAny -and -not $quietOcr) {
                # manual flow: persist the reviewed cursor positions.
                # OCR mode never saves - the workbook content is untouched
                # and a dirty save prompt / file timestamp change is noise.
                Set-WorkbookCursorAllSheets $wb $CursorCell
                $wb.Save()
                Start-Sleep -Milliseconds $SaveWaitMs
            }
            $wb.Close($false)
            $wb = $null
        } catch {
            Write-Host ("  [ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red
            if ($wb) { try { $wb.Close($false) } catch {} }
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

if ($ngList.Count -gt 0) {
    Write-Host ''
    Write-Host ("===== SendVsGift NG rows (marked 2): {0} =====" -f $ngList.Count) -ForegroundColor Red
    foreach ($n in $ngList) {
        Write-Host ("  {0,-12} {1,-16} {2}" -f $n.CorrelIdS, $n.ExcelName, $n.File) -ForegroundColor Red
    }
    Write-Host '  Re-run SendVsGift (NG rows stay pending) after checking them.' -ForegroundColor Yellow
}
