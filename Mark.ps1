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
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),
    [switch]$Force,

    [string]$CommonScript = '',
    [string]$ExcelHelpersScript = '',

    [hashtable]$BoxesConfig = @{},
    [string]$NamePrefix = 'verifyMark_',
    [double]$LineWeight = 1.5,

    # GFIX log yellow-highlight settings. Folded in from the old standalone
    # MarkGfixLog phase: in -Mode Gfix the log "Command:" row is highlighted in
    # the same pass that draws the red rectangles (one workbook open, one bit).
    [string]$GfixLogAnchor = '',
    [string]$GfixLogCommandPattern = "Command:\s*'/appl/[A-Za-z0-9]+/shell/",
    [long]$GfixLogHighlightColor = 65535,
    [int]$GfixLogColStart = 2,
    [int]$GfixLogColEnd   = 51
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

# ── Dot-source ExcelHelpers.ps1 ─────────────────────────────
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)
$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path; break
    }
}
if (-not $helpersPath) { Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red; exit 1 }
. $helpersPath
if (-not (Get-Command -Name 'Add-RedRectangle' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] ExcelHelpers dot-source failed (Add-RedRectangle not loaded).' -ForegroundColor Red; exit 1
}

# ── Target filter ───────────────────────────────────────────
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

# ── Mode config (sheet names + which folders carry marks) ───
$sheetGiftRecv = "GIFT" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GIFT受信結果
$sheetGfixRecv = "GFIX" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GFIX受信結果
$sheetDfDiff   = "GIFT" + [char]0x30C7 + [char]0x30FC + [char]0x30BF +
                 "vs" + "GFIX" + [char]0x30C7 + [char]0x30FC + [char]0x30BF          # GIFTデータvsGFIXデータ

$modeCfg = switch ($Mode) {
    'Gift' { @{
        Sheet = $sheetGiftRecv
        Bit   = 1
        Folders = @('excel','GIFT_HM','GIFT_MQ','GIFT_Jenkins')
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

# ── Header ──────────────────────────────────────────────────
$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
$evDir       = Join-Path $WorkDir 'evidence'

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

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
if ($configuredFolders.Count -eq 0) {
    Write-Host '[WARN] No Boxes configured for any folder in this mode.' -ForegroundColor Yellow
    Write-Host '       Edit VerifyConfig.psd1 -> Mark.Boxes after probing with ProbeShapes.' -ForegroundColor DarkGray
    Write-Host '       Nothing to do; exiting.' -ForegroundColor Yellow
    return
}
Write-Host ("  Boxes     : {0}" -f ($configuredFolders -join ', ')) -ForegroundColor DarkGray
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

# ── Main loop ───────────────────────────────────────────────
$excel = New-ExcelApp
$cntDone = 0
$cntSkip = 0
$cntFail = 0

try {
    foreach ($g in $groups) {
        $first = $g.Group | Select-Object -First 1
        $excelName   = [string]$first.Excel_NAME
        if ([string]::IsNullOrWhiteSpace($excelName)) { continue }
        $excelPrefix = if ($first.PSObject.Properties.Name -contains 'Excel_Prefix') { [string]$first.Excel_Prefix } else { '' }
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

                        $left = 0.0; $top = 0.0; $bw = 100.0; $bh = 20.0
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
                            $ox = 0.0; $oy = 0.0
                            try { $ox = [double]$b.OffsetX } catch {}
                            try { $oy = [double]$b.OffsetY } catch {}
                            try { $bw = [double]$b.Width } catch {}
                            try { $bh = [double]$b.Height } catch {}
                            $left = $picLeft + $ox
                            $top  = $picTop  + $oy
                        }

                        $name = ("{0}{1}_{2}_{3}" -f $NamePrefix, $folder, $cid, $idx)

                        try {
                            Add-RedRectangle $ws $left $top $bw $bh $name $lw | Out-Null
                            $marksDrawn++
                            Write-Host ("  [MARK] {0,-16} {1,-12} [{2,3}] L={3,6:0.0} T={4,6:0.0} W={5,5:0.0} H={6,5:0.0}" -f $folder, $cid, $idx, $left, $top, $bw, $bh) -ForegroundColor Green
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
                        -ColStart $GfixLogColStart -ColEnd $GfixLogColEnd
                    foreach ($w in @($hl.Warnings)) { Write-Host ("  [GfixLog WARN] {0}" -f $w) -ForegroundColor Yellow }
                    Write-Host ("  [GfixLog] highlights applied: {0} (anchors: {1})" -f $hl.Applied, $hl.Anchors) -ForegroundColor DarkGray
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
