# ============================================================
# FillCheckSheet.ps1 - Phase CheckSheet
#
# Appends one row per evidence Excel (grouped by Excel_NAME) to the shared
# review check sheet workbook, sheet "Check Sheet_J4".
#
# The check sheet is a PUBLIC document, so the write is double-checked:
#   1. Snapshot the original file's timestamp + size.
#   2. Copy it to a TEMP file, apply the planned rows there, open it visible
#      so Misaki can eyeball the result.
#   3. She presses Enter to commit, or q to abort.
#   4. On commit the original is re-stat'd; if it changed at all during the
#      preview the write is HELD (nothing is overwritten) so she can re-check.
#      Otherwise the identical edits are applied to the original and saved.
#
# Columns written (1-indexed, configurable; Japanese headers in parens):
#   A No.            : continued from the last numeric No. (only if blank)
#   B (kinyuubi)     : today's date, number format copied from the row above
#   C (COBOL/JAVA)   : Language   (JAVA)
#   D (resource id)  : left blank
#   E (kakunin phase): Phase      (J4 internal review label)
#   F (review target): full evidence filename  <Workbook.ExcelPrefix>_<Excel_NAME>.xlsx
#                      (if no prefix is configured/overridden, EvidenceDir is
#                      checked for the real on-disk filename's prefix first --
#                      see the Resolve-ExcelPrefix / Find-WorkbookByExcelName
#                      block below -- so this never lists a bare name that
#                      doesn't match the actual workbook.)
#   G (tantou)       : Owner
#   H (kakuninsha)   : Viewer     (reviewer short name)
#   I (kanryou kibou-bi) / J~ : left blank
#
# All Japanese (sheet label, phase label, reviewer name) arrives as
# parameters from VerifyConfig.psd1 (BOM) so this source stays pure ASCII.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string[]]$TargetIds = @(),
    [string]$EvidenceDir = '',

    [string]$CheckSheetPath = '',
    [string]$ExcelPrefix = '',
    [string]$SheetName = 'Check Sheet_J4',
    [string]$Language = 'JAVA',
    [string]$Phase = '',
    [string]$Viewer = '',

    [int]$ColNo = 1,
    [int]$ColDate = 2,
    [int]$ColLang = 3,
    [int]$ColResourceId = 4,
    [int]$ColPhase = 5,
    [int]$ColTarget = 6,
    [int]$ColOwner = 7,
    [int]$ColViewer = 8,

    [string]$DateFormat = 'yyyy/m/d',

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

# -- Dot-source ExcelHelpers.ps1 + shared libs ---------------
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

function Get-CellText($ws, [int]$row, [int]$col) {
    $v = $null
    try { $v = $ws.Cells.Item($row, $col).Value2 } catch { return '' }
    if ($null -eq $v) { return '' }
    return ([string]$v).Trim()
}

# Cell/sheet context appended to a write-verify warning so the office-PC
# console log alone can tell apart the known ways a Value2 write comes back
# different from what was written: a protected sheet, a merged cell, and a
# text-formatted ('@') column. Every probe is individually guarded -- a
# diagnostics read must never turn a warning into a new exception.
function Get-CellDiagSuffix($ws, $cell) {
    $parts = @()
    try { if ($null -ne $cell) { $parts += ('addr=' + [string]$cell.Address($false, $false)) } } catch {}
    try { if ($null -ne $cell) { $parts += ('fmt=' + [string]$cell.NumberFormat) } } catch {}
    try { if ($null -ne $cell -and [bool]$cell.MergeCells) { $parts += 'merged=True' } } catch {}
    try { if ([bool]$ws.ProtectContents) { $parts += 'sheetProtected=True' } } catch {}
    if ($parts.Count -eq 0) { return '' }
    return (' [{0}]' -f ($parts -join ', '))
}

function Format-CellReadback($v) {
    if ($null -eq $v) { return '(null)' }
    return ('{0} ({1})' -f $v, $v.GetType().Name)
}

# Set-RangeValue2 (retry-via-InvokeMember Value2 write, fixing PS 5.1's COM
# binder cached-conversion-rule bug -- office-PC log 2026-07-09) now lives in
# ExcelHelpers.ps1 (dot-sourced above), shared with ProcessTime.ps1's
# Write-ProcessTimeWorkbook (v2.14.1), which hit the same bug.

# Writes one cell and reads it back to confirm the value actually stuck --
# a bare try{}catch{} around Value2 assignment swallows both real COM
# exceptions AND silent no-ops (e.g. a locked/merged cell that accepts the
# assignment but doesn't change), either of which used to leave column B
# blank while the run still reported "[OK] wrote N row(s)". Appends a
# message to $Warnings and returns $false on any mismatch/exception instead
# of failing silently. The warning carries the written value, the raw
# readback (value + type) and the cell context (address / NumberFormat /
# merged / sheet protection), so a failure is diagnosable from the console
# log alone. A numeric write whose readback comes back as TEXT (e.g. the
# cell is formatted '@') is compared by parsed value, not by a blind
# [double] cast that would itself throw and mask the real mismatch.
# $AcceptSerial (optional): a date serial the readback may ALSO verify
# against -- used when the date is written as TEXT and Excel parses it into
# the real date value, so Value2 reads back the serial, not the text.
function Set-CellChecked($ws, [int]$row, [int]$col, $value, [string]$label, [System.Collections.Generic.List[string]]$Warnings, $AcceptSerial = $null) {
    $cell = $null
    $via  = ''
    try {
        $cell = $ws.Cells.Item($row, $col)
        $via  = Set-RangeValue2 $cell $value
        $after = $cell.Value2
        $ok = $false
        if ($value -is [double] -or $value -is [int]) {
            if ($after -is [double] -or $after -is [int]) {
                $ok = ([double]$after) -eq ([double]$value)
            } elseif ($null -ne $after) {
                $afterNum = 0.0
                if ([double]::TryParse(([string]$after).Trim(), [ref]$afterNum)) {
                    $ok = $afterNum -eq [double]$value
                }
            }
        } else {
            $ok = ([string]$after).Trim() -eq ([string]$value).Trim()
        }
        if (-not $ok -and $null -ne $AcceptSerial) {
            $acceptNum = [double]$AcceptSerial
            if ($after -is [double] -or $after -is [int]) {
                $ok = ([double]$after) -eq $acceptNum
            } elseif ($null -ne $after) {
                $afterNum = 0.0
                if ([double]::TryParse(([string]$after).Trim(), [ref]$afterNum)) {
                    $ok = $afterNum -eq $acceptNum
                }
            }
        }
        if (-not $ok) {
            $Warnings.Add(("{0} did not verify after write (row {1}): wrote <{2}> via {3}, read back <{4}>{5}" -f `
                $label, $row, $value, $via, (Format-CellReadback $after), (Get-CellDiagSuffix $ws $cell)))
        }
        return $ok
    } catch {
        $viaNote = ''
        if (-not [string]::IsNullOrEmpty($via)) { $viaNote = (' (via {0})' -f $via) }
        $Warnings.Add(("{0} write failed (row {1}): {2}{3}{4}" -f `
            $label, $row, $_.Exception.Message, $viaNote, (Get-CellDiagSuffix $ws $cell)))
        return $false
    }
}

# Computes and (optionally) writes the new rows on $ws. Returns an array of
# summary objects { Row; No; Target; Action }. Re-running computes identical
# results on the original as on the temp copy (same content), so the preview
# and the commit stay in lock-step.
function Apply-CheckSheetRows($ws, [object[]]$Candidates, [bool]$WriteCells) {
    $summary = New-Object System.Collections.Generic.List[object]

    $xlUp = -4162
    $lastB = 0
    try { $lastB = [int]$ws.Cells.Item($ws.Rows.Count, $ColDate).End($xlUp).Row } catch { $lastB = 0 }
    # End(xlUp) lands on row 1 even when empty; treat a blank B1 as "no data".
    if ($lastB -eq 1 -and [string]::IsNullOrWhiteSpace((Get-CellText $ws 1 $ColDate))) { $lastB = 0 }
    $startRow = $lastB + 1

    # Existing review targets already listed (column F), for de-dup.
    $existing = @{}
    for ($r = 1; $r -le $lastB; $r++) {
        $t = Get-CellText $ws $r $ColTarget
        if (-not [string]::IsNullOrWhiteSpace($t)) { $existing[$t.ToLower()] = $true }
    }

    # Last numeric No. above the insertion point (skip the "rei" sample row).
    $lastNo = 0
    for ($r = $lastB; $r -ge 1; $r--) {
        $a = Get-CellText $ws $r $ColNo
        $n = 0
        if ([int]::TryParse($a, [ref]$n)) { $lastNo = $n; break }
    }

    # Date format to mirror from the last filled B cell. '@' (text) and
    # 'General' are never mirrored: writing the OADate serial into either
    # renders as a bare number (or stores text), which is exactly the
    # blank/garbled column-B symptom -- fall back to $DateFormat instead.
    $dateFmt = $DateFormat
    if ($lastB -ge 1) {
        try {
            $f = [string]$ws.Cells.Item($lastB, $ColDate).NumberFormat
            if (-not [string]::IsNullOrWhiteSpace($f) -and $f -ne '@' -and $f -ne 'General') { $dateFmt = $f }
        } catch {}
    }

    $today    = (Get-Date).Date
    $todaySer = [double]$today.ToOADate()
    $todayText = $today.ToString('yyyy/MM/dd')
    $nextNo   = $lastNo + 1
    $row      = $startRow

    foreach ($cand in $Candidates) {
        $target = [string]$cand.TargetFile
        if ((-not $forceFlag) -and $existing.ContainsKey($target.ToLower())) {
            $summary.Add([pscustomobject]@{ Row = 0; No = ''; Date = ''; Target = $target; Action = 'skip (already listed)' })
            continue
        }

        # No. -- only fill when blank; always keep numbering monotonic.
        $noText = Get-CellText $ws $row $ColNo
        $thisNo = $noText
        if ([string]::IsNullOrWhiteSpace($noText)) {
            $thisNo = [string]$nextNo
            if ($WriteCells) { try { $ws.Cells.Item($row, $ColNo).Value2 = $nextNo } catch {} }
            $nextNo++
        } else {
            $parsed = 0
            if ([int]::TryParse($noText, [ref]$parsed) -and $parsed -ge $nextNo) { $nextNo = $parsed + 1 }
        }

        $rowOk = $true
        if ($WriteCells) {
            $warnings = New-Object System.Collections.Generic.List[string]
            # Set the date NumberFormat BEFORE writing the value: writing the
            # OADate serial into a cell still formatted '@' (text) stores the
            # digits as text, and re-formatting afterwards does NOT convert it
            # back -- the cell keeps showing "46212"-style text. Format-first
            # makes the serial land as a real date in one pass.
            try { $ws.Cells.Item($row, $ColDate).NumberFormat = $dateFmt } catch {
                Write-Host ("    [WARN] row {0}: date NumberFormat set failed: {1}" -f $row, $_.Exception.Message) -ForegroundColor Yellow
            }
            # Date write, two tiers. Tier 1: the OADate serial as a Double
            # (Set-RangeValue2 already retries the PS-COM-binder
            # InvalidCastException via InvokeMember). Tier 2, last resort:
            # write the date as TEXT -- the cell already carries the date
            # NumberFormat, so Excel parses it into a real date value;
            # AcceptSerial lets the verify pass on the parsed serial. Tier
            # 1's warning is only surfaced when tier 2 also fails, so a
            # recovered date doesn't leave a scary-but-stale WARN behind.
            $dateTier1 = New-Object System.Collections.Generic.List[string]
            $dateOk = Set-CellChecked $ws $row $ColDate $todaySer 'date' $dateTier1
            if (-not $dateOk) {
                $dateOk = Set-CellChecked $ws $row $ColDate $todayText 'date-as-text' $warnings $todaySer
                if ($dateOk) {
                    Write-Host ("    [INFO] row {0}: serial date write failed ({1}); recovered by writing the date as text." -f `
                        $row, ($dateTier1 -join ' / ')) -ForegroundColor DarkGray
                } else {
                    foreach ($w in $dateTier1) { $warnings.Add($w) }
                }
            }
            $langOk   = Set-CellChecked $ws $row $ColLang   $Language 'language' $warnings
            $phaseOk  = Set-CellChecked $ws $row $ColPhase  $Phase    'phase'    $warnings
            $targetOk = Set-CellChecked $ws $row $ColTarget $target   'target'   $warnings
            $ownerOk  = Set-CellChecked $ws $row $ColOwner  $Owner    'owner'    $warnings
            $viewerOk = Set-CellChecked $ws $row $ColViewer $Viewer   'viewer'   $warnings
            $rowOk = $dateOk -and $langOk -and $phaseOk -and $targetOk -and $ownerOk -and $viewerOk
            foreach ($w in $warnings) { Write-Host ("    [WARN] row {0}: {1}" -f $row, $w) -ForegroundColor Yellow }
        }

        $summary.Add([pscustomobject]@{ Row = $row; No = $thisNo; Date = $todayText; Target = $target; Action = 'add'; Ok = $rowOk })
        $row++
    }

    return $summary.ToArray()
}

function Show-Plan([object[]]$Plan) {
    Write-Host ''
    Write-Host '  Planned check-sheet rows:' -ForegroundColor Cyan
    foreach ($p in $Plan) {
        if ($p.Action -eq 'add') {
            Write-Host ("    row {0,-4} No.{1,-4} {2,-10} {3}" -f $p.Row, $p.No, $p.Date, $p.Target) -ForegroundColor White
        } else {
            Write-Host ("    {0,-14} {1}" -f $p.Action, $p.Target) -ForegroundColor DarkGray
        }
    }
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red; exit 1
}

# Resolve the check sheet path; prompt if the configured one does not exist.
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -or -not (Test-Path -LiteralPath $CheckSheetPath)) {
    if (-not [string]::IsNullOrWhiteSpace($CheckSheetPath)) {
        Write-Host ("[INFO] configured check sheet not found: {0}" -f $CheckSheetPath) -ForegroundColor Yellow
    }
    $CheckSheetPath = Read-Host 'Review check sheet (.xlsx) full path'
}
if ([string]::IsNullOrWhiteSpace($CheckSheetPath) -or -not (Test-Path -LiteralPath $CheckSheetPath)) {
    Write-Host "[ERROR] check sheet not found: $CheckSheetPath" -ForegroundColor Red; exit 1
}
$CheckSheetPath = (Resolve-Path -LiteralPath $CheckSheetPath).ProviderPath

$targets = @(ConvertTo-TargetIdList $TargetIds)

$allRows = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red; exit 1
}

# One candidate per Excel_NAME (mapping order), honoring the target filter.
$candidates = New-Object System.Collections.Generic.List[object]
$seen = @{}
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if ($seen.ContainsKey($name)) { continue }
    $seen[$name] = $true
    $prefix = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
    if ([string]::IsNullOrWhiteSpace($prefix)) {
        # Neither the mapping row (legacy Excel_Prefix) nor Workbook.ExcelPrefix
        # gave a prefix. Before listing a bare, unprefixed name on the shared
        # check sheet, recover the prefix the real evidence file already
        # carries on disk (shared WorkbookResolver helper; DeliverMail uses
        # the same one for the mail-body filename).
        $prefix = Resolve-ExcelPrefixWithDisk -Row $r -DefaultPrefix $ExcelPrefix -ExcelName $name -EvidenceDir $EvidenceDir
        if (-not [string]::IsNullOrWhiteSpace($prefix)) {
            Write-Host ("  [INFO] no configured prefix for {0}; using prefix found on disk: {1}" -f $name, $prefix) -ForegroundColor DarkGray
        }
    }
    $fullStem = Get-ExcelFullStem -Prefix $prefix -Name $name
    $candidates.Add([pscustomobject]@{ ExcelName = $name; TargetFile = (Get-ExcelDestLeaf $fullStem) })
}
if ($candidates.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow
    return
}

Write-Host ''
Write-Host '===== FillCheckSheet =====' -ForegroundColor Green
Write-Host ("  WorkDir    : {0}" -f $WorkDir)
Write-Host ("  EvidenceDir: {0}" -f $EvidenceDir)
Write-Host ("  CheckSheet : {0}" -f $CheckSheetPath)
Write-Host ("  Sheet      : {0}" -f $SheetName)
Write-Host ("  Candidates : {0}" -f $candidates.Count)
Write-Host ("  Owner      : {0}" -f $Owner)
Write-Host ("  Force      : {0}" -f $forceFlag)
if ($targets.Count -gt 0) { Write-Host ("  TargetIds  : {0}" -f ($targets -join ', ')) }

$excel = $null
$wbCs  = $null
$mappedLetter = $null

try {
    $excel = New-ExcelApp
    $excel.Visible = $true
    try { $excel.DisplayAlerts = $false } catch {}
    # New-ExcelApp turns ScreenUpdating OFF for headless speed. This phase
    # shows the workbook to the operator for a visual check, so turn it back
    # ON -- otherwise the visible window never repaints and the preview looks
    # blank (no name, no rows/columns) even though the content is correct.
    try { $excel.ScreenUpdating = $true } catch {}

    # Map temp drive for long UNC paths (PS5.1 MAX_PATH workaround). Applies
    # to both the DryRun read-only open and the real one below -- a long
    # check-sheet UNC path fails the same way in either mode.
    function Get-FreeDriveLetter {
        $used = @(Get-PSDrive -PSProvider FileSystem | Select-Object -ExpandProperty Name) + @('A','B','C')
        foreach ($l in 'Z','Y','X','W','V','U','T','S','R','Q','P') {
            if ($used -notcontains $l) { return $l }
        }
        return $null
    }

    $effectivePath = $CheckSheetPath
    $csParent = Split-Path $CheckSheetPath -Parent
    $csLeaf   = Split-Path $CheckSheetPath -Leaf

    if ($CheckSheetPath.Length -gt 200 -and $CheckSheetPath -match '^\\\\') {
        $letter = Get-FreeDriveLetter
        if ($letter) {
            $netOut = & net use "${letter}:" $csParent /persistent:no 2>&1
            if ($LASTEXITCODE -eq 0) {
                $mappedLetter  = $letter
                $effectivePath = "${letter}:\${csLeaf}"
                Write-Host ("  [INFO] mapped {0}: to shorten UNC path" -f $letter) -ForegroundColor DarkGray
            } else {
                Write-Host ("  [WARN] net use failed: {0}" -f ($netOut -join ' ')) -ForegroundColor Yellow
            }
        }
    }

    # -- DryRun: compute on a read-only open of the original, no writes --
    if ($dryRunFlag) {
        $wbRo = $excel.Workbooks.Open($effectivePath, 0, $true)   # ReadOnly
        try {
            $wsRo = Get-SheetByName $wbRo $SheetName
            if ($null -eq $wsRo) { Write-Host ("[ERROR] sheet not found: {0}" -f $SheetName) -ForegroundColor Red; return }
            $plan = Apply-CheckSheetRows $wsRo $candidates.ToArray() $false
            Show-Plan $plan
            Write-Host '  [DRY RUN] nothing written.' -ForegroundColor Yellow
        } finally {
            try { $wbRo.Close($false) } catch {}
        }
        return
    }

    $wbCs = $excel.Workbooks.Open($effectivePath)
    $wsCs = Get-SheetByName $wbCs $SheetName
    if ($null -eq $wsCs) {
        Write-Host ("[ERROR] sheet not found: {0}" -f $SheetName) -ForegroundColor Red
        return
    }
    try { $wsCs.Visible = -1 } catch {}
    $wsCs.Activate() | Out-Null

    # Preview (read-only pass, no writes)
    $plan = Apply-CheckSheetRows $wsCs $candidates.ToArray() $false
    $added = @($plan | Where-Object { $_.Action -eq 'add' })
    Show-Plan $plan

    if ($added.Count -eq 0) {
        Write-Host '  Nothing to add (all candidates already listed). Use -Force to add anyway.' -ForegroundColor Yellow
        return
    }

    try { $wsCs.Range($wsCs.Cells.Item($added[0].Row, 1), $wsCs.Cells.Item($added[-1].Row, $ColViewer)).Select() | Out-Null } catch {}
    Write-Host ''
    Write-Host '  Planned rows shown above. Check workbook, then confirm.' -ForegroundColor DarkGray
    $ans = Read-Host '  Enter=write to check sheet, q=abort'
    if ($ans.Trim().ToLower() -eq 'q') {
        Write-Host '  [ABORT] no changes written.' -ForegroundColor Yellow
        return
    }

    # Commit: write rows and save
    $plan2  = Apply-CheckSheetRows $wsCs $candidates.ToArray() $true
    $added2 = @($plan2 | Where-Object { $_.Action -eq 'add' })
    $wbCs.Save()

    $okRows     = @($added2 | Where-Object { $_.Ok })
    $failedRows = @($added2 | Where-Object { -not $_.Ok })
    Write-Host ("  [OK] wrote {0} row(s) to {1}" -f $okRows.Count, $CheckSheetPath) -ForegroundColor Green
    if ($failedRows.Count -gt 0) {
        Write-Host ("  [WARN] {0} row(s) had a write that failed to verify -- open the workbook and check before trusting this run:" -f $failedRows.Count) -ForegroundColor Yellow
        foreach ($p in $failedRows) { Write-Host ("    row {0} ({1})" -f $p.Row, $p.Target) -ForegroundColor Yellow }
    }
    foreach ($p in $added2) {
        $status = if ($p.Ok) { 'ok' } else { 'warn' }
        Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -JobName $p.Target -Action 'commit' -Status $status -Message ("No.{0}" -f $p.No)
    }

} catch {
    Write-Host ("[ERROR] {0}" -f $_.Exception.Message) -ForegroundColor Red
    Write-ProgressEvent -WorkDir $WorkDir -Phase 'CheckSheet' -Action 'commit' -Status 'fail' -Message $_.Exception.Message
} finally {
    if ($wbCs)  { try { $wbCs.Close($false) } catch {} }
    if ($excel) { Close-ExcelApp $excel }
    if ($mappedLetter) { try { & net use "${mappedLetter}:" /delete /y 2>&1 | Out-Null } catch {} }
}

Write-Host ''
Write-Host '===== FillCheckSheet Done =====' -ForegroundColor Green
