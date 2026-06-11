# ============================================================
# ReviewEvidence.ps1 - manual visual review driver
#
# Opens evidence workbooks one by one in editable mode. Misaki reviews/adjusts by hand.
#
# Important behavior:
#   - Opening a workbook jumps to the pending ID on the send-data sheet
#     (column A exact match) so the operator can review that ID first.
#   - Pressing Enter marks only the current ID. If the same workbook still has
#     pending IDs, the workbook stays open and the cursor jumps to the next ID.
#   - Only after all IDs in the workbook are reviewed, the script:
#       1. places cursor on each sheet from last to first
#          - A1 if sheet name contains the "=>" Excel sheet arrow
#          - otherwise CursorCell, default A3
#       2. sends Ctrl+S, waits, sends Esc to dismiss GenBa comment prompt
#       3. closes the workbook
#
# Read-only handling:
#   - The script requests read-write open explicitly.
#   - "Read-only recommended" is ignored.
#   - If the file has the Windows ReadOnly attribute, the attribute is cleared before open.
#   - If Excel still opens it as read-only, the item fails instead of silently saving nothing.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = '',
    [string]$EvidenceDir = '',
    [string]$ExcelPrefix = '',
    [string]$CursorCell = 'A3',
    [string]$ReviewField = 'isReviewed',
    [int]$ReviewBit = 7,
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

try {
    Get-ChildItem -LiteralPath $PSScriptRoot -Filter '*.ps1' -File -ErrorAction SilentlyContinue |
        ForEach-Object { Unblock-File -LiteralPath $_.FullName -ErrorAction SilentlyContinue }
} catch {}

$forceFlag = [bool]$Force.IsPresent
$dryRunFlag = [bool]$DryRun.IsPresent
$maximizeFlag = [bool]$Maximize.IsPresent

# ── Dot-source ExcelHelpers.ps1 ─────────────────────────────
$candidates = @()
if (-not [string]::IsNullOrWhiteSpace($ExcelHelpersScript)) { $candidates += $ExcelHelpersScript }
$candidates += @(
    (Join-Path $PSScriptRoot 'ExcelHelpers.ps1')
)

$helpersPath = $null
foreach ($c in $candidates) {
    if (-not [string]::IsNullOrWhiteSpace($c) -and (Test-Path -LiteralPath $c)) {
        $helpersPath = (Resolve-Path -LiteralPath $c).Path
        break
    }
}
if (-not $helpersPath) {
    Write-Host '[ERROR] ExcelHelpers.ps1 not found.' -ForegroundColor Red
    exit 1
}
. $helpersPath
. (Join-Path $PSScriptRoot 'ProjectLabels.ps1')   # no param() -> safe to dot-source

function Test-TargetRow($Row, [hashtable]$TargetSet) {
    if ($TargetSet.Count -eq 0) { return $true }
    return ($TargetSet.ContainsKey([string]$Row.Correl_ID_S) -or
            $TargetSet.ContainsKey([string]$Row.Correl_ID_M) -or
            $TargetSet.ContainsKey([string]$Row.JOB_NAME) -or
            $TargetSet.ContainsKey([string]$Row.Excel_NAME))
}

function Test-BitDone([array]$Rows, [string]$Field, [int]$Bit) {
    if ($Rows.Count -eq 0) { return $false }
    foreach ($r in $Rows) {
        $v = Get-BitValue $r $Field
        if (($v -band $Bit) -ne $Bit) { return $false }
    }
    return $true
}

. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

function Get-EvidencePath([string]$Dir, [string]$ExcelName) {
    $resolved = Find-WorkbookByExcelName -Dir $Dir -ExcelName $ExcelName
    if ($null -ne $resolved) { return $resolved }
    if ($ExcelName -match '\.xlsx$') { return (Join-Path $Dir $ExcelName) }
    return (Join-Path $Dir ($ExcelName + '.xlsx'))
}

function Get-ExcelProcessId($Excel) {
    try {
        $hwnd = [IntPtr]([int]$Excel.Hwnd)
        $p = Get-Process | Where-Object { $_.MainWindowHandle -eq $hwnd } | Select-Object -First 1
        if ($p) { return [int]$p.Id }
    } catch {}
    return 0
}

function Activate-ExcelWindow($Shell, $Excel, $Workbook) {
    $excelPid = Get-ExcelProcessId $Excel
    if ($excelPid -gt 0) {
        try {
            if ($Shell.AppActivate($excelPid)) { return $true }
        } catch {}
    }
    try {
        if ($Workbook -and $Shell.AppActivate([string]$Workbook.Name)) { return $true }
    } catch {}
    try {
        if ($Shell.AppActivate('Excel')) { return $true }
    } catch {}
    return $false
}

function Clear-FileReadOnlyAttribute([string]$File) {
    try {
        $item = Get-Item -LiteralPath $File -ErrorAction Stop
        if (($item.Attributes -band [System.IO.FileAttributes]::ReadOnly) -ne 0) {
            $item.Attributes = ($item.Attributes -band (-bnot [System.IO.FileAttributes]::ReadOnly))
            Write-Host '  [INFO] Windows ReadOnly attribute cleared before open.' -ForegroundColor DarkGray
        }
    } catch {
        Write-Host ("  [WARN] failed to clear Windows ReadOnly attribute: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Open-WorkbookEditable($Excel, [string]$File) {
    Clear-FileReadOnlyAttribute $File

    $missing = [Type]::Missing

    # Workbooks.Open parameters used here:
    #   Filename, UpdateLinks, ReadOnly, Format, Password, WriteResPassword,
    #   IgnoreReadOnlyRecommended, Origin, Delimiter, Editable, Notify,
    #   Converter, AddToMru, Local
    #
    # Points:
    #   - ReadOnly=$false requests editable open.
    #   - IgnoreReadOnlyRecommended=$true avoids Excel's "read-only recommended" prompt/default.
    #   - Notify=$false prevents Excel from opening a locked workbook as read-only while waiting.
    #   - AddToMru=$false avoids polluting Recent files.
    return $Excel.Workbooks.Open(
        $File,
        0,
        $false,
        $missing,
        $missing,
        $missing,
        $true,
        $missing,
        $missing,
        $true,
        $false,
        $missing,
        $false,
        $true
    )
}

# Splits raw review input into an action ('' = done, 's' = skip, 'q' = quit)
# and an optional comment introduced by  -m "comment"  (single or double
# quotes optional). Examples:
#   ''             -> done,  no comment
#   's'            -> skip,  no comment
#   '-m "looks ok"'-> done,  comment="looks ok"
#   's -m fix font'-> skip,  comment="fix font"
#   'q -m "stop"'  -> quit,  comment="stop"
function Parse-ReviewInput([string]$Raw) {
    $comment = ''
    $action  = ''
    if ($null -ne $Raw) {
        $s = $Raw.Trim()
        $m = [regex]::Match($s, '(?:^|\s)-m\s+(.+)$')
        if ($m.Success) {
            $comment = $m.Groups[1].Value.Trim()
            if ($comment.Length -ge 2 -and
                (($comment[0] -eq '"' -and $comment[-1] -eq '"') -or
                 ($comment[0] -eq "'" -and $comment[-1] -eq "'"))) {
                $comment = $comment.Substring(1, $comment.Length - 2)
            }
            $s = $s.Substring(0, $m.Index).Trim()
        }
        $action = $s.ToLower()
    }
    return [pscustomobject]@{ Action = $action; Comment = $comment }
}

function Read-ReviewChoice([datetime]$OpenedAt) {
    $raw    = Read-Host '  Enter=save+close+mark, s=skip, q=quit   ( add  -m "comment"  to record a note )'
    $parsed = Parse-ReviewInput $raw

    # Safety guard:
    # A blank Enter (no action AND no comment) right after open is often an
    # accidental/stale keypress. Ask once more so the workbook does not close
    # immediately after opening.
    if ([string]::IsNullOrWhiteSpace($parsed.Action) -and [string]::IsNullOrWhiteSpace($parsed.Comment)) {
        $elapsed = ((Get-Date) - $OpenedAt).TotalSeconds
        if ($elapsed -lt 2) {
            Write-Host '  [WARN] Enter was received immediately after open. Confirm once more to close this workbook.' -ForegroundColor Yellow
            $raw    = Read-Host '  Enter again=save+close+mark, s=skip, q=quit   ( -m "comment" to note )'
            $parsed = Parse-ReviewInput $raw
        }
    }

    return $parsed
}

# Activates the sheet relevant to the review mode (e.g. ReviewGift -> GIFT
# jushin kekka) so it is the front sheet when Misaki starts reviewing.

function Test-RowBitDone($Row, [string]$Field, [int]$Bit) {
    $v = Get-BitValue $Row $Field
    return (($v -band $Bit) -eq $Bit)
}

function Get-RowReviewId($Row) {
    foreach ($prop in @('Correl_ID_S', 'Correl_ID_M', 'JOB_NAME')) {
        if ($Row.PSObject.Properties.Name -contains $prop) {
            $v = [string]$Row.$prop
            if (-not [string]::IsNullOrWhiteSpace($v)) { return $v.Trim() }
        }
    }
    return ''
}

function Get-PendingReviewRows([array]$Rows, [string]$Field, [int]$Bit, [bool]$ForceFlag) {
    if ($ForceFlag) { return @($Rows) }
    return @($Rows | Where-Object { -not (Test-RowBitDone $_ $Field $Bit) })
}

function Get-WorksheetByName($Workbook, [string]$SheetName) {
    if ([string]::IsNullOrWhiteSpace($SheetName)) { return $null }
    foreach ($s in $Workbook.Worksheets) {
        if ([string]$s.Name -eq $SheetName) { return $s }
    }
    return $null
}

function Move-ToSendDataId($Workbook, [string]$Id, [string]$FallbackCell, [string]$SendSheetName) {
    $cell = $FallbackCell
    if ([string]::IsNullOrWhiteSpace($cell)) { $cell = 'A3' }

    $ws = Get-WorksheetByName $Workbook $SendSheetName
    if ($null -eq $ws) {
        Write-Host ("  [WARN] send-data sheet not found, using current sheet/cell {0}: {1}" -f $cell, $SendSheetName) -ForegroundColor Yellow
        try { $Workbook.ActiveSheet.Range($cell).Select() | Out-Null } catch {}
        return $false
    }

    try { $ws.Visible = -1 } catch {}
    try { $ws.Activate() | Out-Null } catch {}

    if ([string]::IsNullOrWhiteSpace($Id)) {
        try { $ws.Range($cell).Select() | Out-Null } catch {}
        return $false
    }

    $lastRow = 1
    try { $lastRow = [int]$ws.Cells.Item($ws.Rows.Count, 1).End(-4162).Row } catch { $lastRow = 1 } # xlUp
    if ($lastRow -lt 1) { $lastRow = 1 }

    for ($row = 1; $row -le $lastRow; $row++) {
        $text = ''
        try { $text = [string]$ws.Cells.Item($row, 1).Text } catch {}
        if ([string]::IsNullOrWhiteSpace($text)) {
            try { $text = [string]$ws.Cells.Item($row, 1).Value2 } catch {}
        }
        if ($text.Trim() -eq $Id) {
            $addr = ('A{0}' -f $row)
            try { $ws.Range($addr).Select() | Out-Null } catch {}
            Write-Host ("  [CURSOR] {0}!{1} for ID {2}" -f $SendSheetName, $addr, $Id) -ForegroundColor DarkGray
            return $true
        }
    }

    try { $ws.Range($cell).Select() | Out-Null } catch {}
    Write-Host ("  [WARN] ID not found in {0} column A: {1}; cursor={2}" -f $SendSheetName, $Id, $cell) -ForegroundColor Yellow
    return $false
}

function Open-SheetForReview($Workbook, [string]$SheetName, [string]$Cell) {
    if ([string]::IsNullOrWhiteSpace($SheetName)) { return }
    try {
        $ws = $null
        foreach ($s in $Workbook.Worksheets) { if ([string]$s.Name -eq $SheetName) { $ws = $s; break } }
        if ($null -eq $ws) {
            Write-Host ("  [WARN] review sheet not found, leaving default: {0}" -f $SheetName) -ForegroundColor Yellow
            return
        }
        try { $ws.Visible = -1 } catch {}   # xlSheetVisible
        $ws.Activate() | Out-Null
        try { $ws.Range($Cell).Select() | Out-Null } catch {}
        Write-Host ("  [SHEET] opened: {0}" -f $SheetName) -ForegroundColor DarkGray
    } catch {
        Write-Host ("  [WARN] could not open review sheet {0}: {1}" -f $SheetName, $_.Exception.Message) -ForegroundColor Yellow
    }
}

function Set-ReviewCursorAllSheets($Workbook, [string]$DefaultCell) {
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

function Save-BySendKeys($Shell, $Excel, $Workbook, [int]$WaitMs) {
    if ($Workbook.ReadOnly) {
        throw "Workbook is read-only. Save is blocked: $($Workbook.FullName)"
    }

    [void](Activate-ExcelWindow $Shell $Excel $Workbook)
    Start-Sleep -Milliseconds 200

    # GenBa macro prompt appears immediately after Ctrl+S.
    # Dismiss it first, then wait for network-share save to settle.
    $Shell.SendKeys('^s')
    Start-Sleep -Milliseconds 300
    $Shell.SendKeys('{ESC}')
    Start-Sleep -Milliseconds ([Math]::Max(800, $WaitMs))
}

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir 'evidence' }
elseif (-not [System.IO.Path]::IsPathRooted($EvidenceDir)) { $EvidenceDir = Join-Path $WorkDir $EvidenceDir }

if (-not (Test-Path -LiteralPath $EvidenceDir)) {
    Write-Host "[ERROR] EvidenceDir not found: $EvidenceDir" -ForegroundColor Red
    exit 1
}

if ([string]::IsNullOrWhiteSpace($ReviewField)) { $ReviewField = 'isReviewed' }
if ($ReviewBit -le 0) { $ReviewBit = 7 }

$mappingPath = Join-Path $WorkDir ("mapping_{0}.csv" -f $Owner)
if (-not (Test-Path -LiteralPath $mappingPath)) {
    Write-Host "[ERROR] mapping not found: $mappingPath" -ForegroundColor Red
    exit 1
}

$targetSet = @{}
foreach ($rawId in @($TargetIds)) {
    if ($null -eq $rawId) { continue }
    foreach ($part in ($rawId.ToString() -split ',')) {
        $v = $part.Trim()
        if (-not [string]::IsNullOrWhiteSpace($v)) { $targetSet[$v] = $true }
    }
}

$allRows = @(Import-Csv -LiteralPath $mappingPath -Encoding UTF8)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red
    exit 1
}
Ensure-Column $allRows $ReviewField '0'
Ensure-Column $allRows 'ReviewComment' ''

# Sheet to bring to front per review mode (ReviewGift -> GIFT jushin kekka,
# ReviewGfix -> GFIX jushin kekka, ReviewDf -> DF compare). ReviewEvidence
# (bit 7 = all) leaves the workbook's default sheet alone.
$labels = Get-ProjectLabels
$sendDataSheetName = $labels['SheetSoshinData']
$openSheetName = switch ($ReviewBit) {
    1 { $labels['SheetGiftRecv'] }
    2 { $labels['SheetGfixRecv'] }
    4 { $labels['SheetDfCompare'] }
    default { '' }
}

$selectedRows = @($allRows | Where-Object { Test-TargetRow $_ $targetSet })
if ($selectedRows.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow
    return
}

$excelNames  = New-Object System.Collections.Generic.List[string]
$prefixByName = @{}
foreach ($r in $selectedRows) {
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $excelNames.Contains($name)) {
        $excelNames.Add($name)
        $prefixByName[$name] = Resolve-ExcelPrefix -Row $r -DefaultPrefix $ExcelPrefix
    }
}

Write-Host ''
Write-Host '===== ReviewEvidence =====' -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Mapping     : {0}" -f $mappingPath)
Write-Host ("  EvidenceDir : {0}" -f $EvidenceDir)
Write-Host ("  Field       : {0}" -f $ReviewField)
Write-Host ("  ReviewBit   : {0}" -f $ReviewBit)
Write-Host ("  Workbooks   : {0}" -f $excelNames.Count)
Write-Host ("  Force       : {0}" -f $forceFlag)
if ($targetSet.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f (($targetSet.Keys | Sort-Object) -join ', ')) }

if ($dryRunFlag) {
    foreach ($name in $excelNames) {
        $groupRows = @($allRows | Where-Object { [string]$_.Excel_NAME -eq $name })
        $done = Test-BitDone $groupRows $ReviewField $ReviewBit
        Write-Host ("  [DRY] {0} rows={1} done={2}" -f $name, $groupRows.Count, $done)
    }
    return
}

$excel = $null
$shell = $null
$cntDone = 0
$cntSkip = 0
$cntMissing = 0
$cntFail = 0

try {
    $excel = New-ExcelApp
    $excel.Visible = $true
    try { $excel.DisplayAlerts = $false } catch {}
    try { $excel.ScreenUpdating = $true } catch {}
    if ($maximizeFlag) {
        try { $excel.WindowState = -4137 } catch {} # xlMaximized
    }

    $shell = New-Object -ComObject WScript.Shell

    for ($idx = 0; $idx -lt $excelNames.Count; $idx++) {
        $name = $excelNames[$idx]
        $groupRows = @($allRows | Where-Object { [string]$_.Excel_NAME -eq $name })

        if ((Test-BitDone $groupRows $ReviewField $ReviewBit) -and -not $forceFlag) {
            Write-Host ("[{0}/{1}] SKIP reviewed: {2}" -f ($idx + 1), $excelNames.Count, $name) -ForegroundColor DarkGray
            $cntSkip++
            continue
        }

        $fullStem = Get-ExcelFullStem -Prefix ($prefixByName[$name]) -Name $name
        $file = Get-EvidencePath $EvidenceDir $fullStem
        if (-not (Test-Path -LiteralPath $file)) {
            Write-Host ("[{0}/{1}] MISSING: {2}" -f ($idx + 1), $excelNames.Count, $file) -ForegroundColor Yellow
            $cntMissing++
            continue
        }

        $wb = $null
        try {
            Write-Host ''
            Write-Host ("[{0}/{1}] OPEN: {2}" -f ($idx + 1), $excelNames.Count, $file) -ForegroundColor Cyan
            $openedAt = Get-Date
            $wb = Open-WorkbookEditable $excel $file

            if ($wb.ReadOnly) {
                try { $wb.Close($false) } catch {}
                $wb = $null
                throw "Workbook opened as read-only. Possible causes: another Excel window/user is locking it, write permission is missing, or the file/share is read-only: $file"
            }

            $workbookCompleted = $false
            $sessionDone = @{}
            while ($true) {
                $pendingGroupRows = @(Get-PendingReviewRows $groupRows $ReviewField $ReviewBit $forceFlag | Where-Object { -not $sessionDone.ContainsKey((Get-RowReviewId $_)) })
                if ($pendingGroupRows.Count -eq 0) {
                    $workbookCompleted = ($forceFlag -or (Test-BitDone $groupRows $ReviewField $ReviewBit))
                    break
                }

                $currentRow = $pendingGroupRows[0]
                $currentId = Get-RowReviewId $currentRow
                Write-Host ("  [ID] {0}/{1}: {2}" -f (($groupRows.Count - $pendingGroupRows.Count) + 1), $groupRows.Count, $currentId) -ForegroundColor Cyan

                # Bring the current ID to the front before review starts.
                [void](Move-ToSendDataId $wb $currentId $CursorCell $sendDataSheetName)
                [void](Activate-ExcelWindow $shell $excel $wb)

                # Surface any prior comment recorded for this row.
                $priorComment = ''
                if ($currentRow.PSObject.Properties.Name -contains 'ReviewComment') { $priorComment = [string]$currentRow.ReviewComment }
                if (-not [string]::IsNullOrWhiteSpace($priorComment)) {
                    Write-Host ("  [COMMENT] {0}" -f $priorComment) -ForegroundColor Yellow
                }

                Write-Host '  Workbook opened in editable mode. Review/edit this ID in Excel, then return here manually.' -ForegroundColor DarkGray
                $choice = Read-ReviewChoice $openedAt
                $openedAt = Get-Date

                # Record a -m comment (if given) on the current row only, and
                # persist immediately for all outcomes (done / skip / quit).
                if (-not [string]::IsNullOrWhiteSpace($choice.Comment)) {
                    if (-not ($currentRow.PSObject.Properties.Name -contains 'ReviewComment')) {
                        $currentRow | Add-Member -NotePropertyName 'ReviewComment' -NotePropertyValue '' -Force
                    }
                    $currentRow.ReviewComment = $choice.Comment
                    $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
                    Write-Host ("  [COMMENT] recorded: {0}" -f $choice.Comment) -ForegroundColor DarkCyan
                }

                if ($choice.Action -eq 'q') {
                    try { $wb.Close($false) } catch {}
                    $wb = $null
                    break
                }
                if ($choice.Action -eq 's') {
                    Write-Host ("  [SKIP] ID left pending: {0}" -f $currentId) -ForegroundColor Yellow
                    $sessionDone[$currentId] = $true
                    $cntSkip++
                    if ($pendingGroupRows.Count -le 1) { break }
                    continue
                }

                Set-BitValue $currentRow $ReviewField $ReviewBit
                $sessionDone[$currentId] = $true
                $allRows | Export-Csv -LiteralPath $mappingPath -Encoding UTF8 -NoTypeInformation -Force
                Write-Host ("  [OK] reviewed bit set for ID {0} -> {1}" -f $currentId, $ReviewBit) -ForegroundColor Green

                $actualRemaining = @(Get-PendingReviewRows $groupRows $ReviewField $ReviewBit $false)
                $remaining = @(Get-PendingReviewRows $groupRows $ReviewField $ReviewBit $forceFlag | Where-Object { -not $sessionDone.ContainsKey((Get-RowReviewId $_)) })
                if (($forceFlag -and $remaining.Count -gt 0) -or ((-not $forceFlag) -and $actualRemaining.Count -gt 0 -and $remaining.Count -gt 0)) {
                    $nextId = Get-RowReviewId $remaining[0]
                    Write-Host ("  [NEXT] workbook still has {0} pending ID(s); jumping to {1} without save/close/reset." -f $remaining.Count, $nextId) -ForegroundColor DarkGray
                    continue
                }

                $workbookCompleted = if ($forceFlag) { $remaining.Count -eq 0 } else { $actualRemaining.Count -eq 0 }
                break
            }

            if ($null -eq $wb) { break }

            if (-not $workbookCompleted) {
                Write-Host ("  [INFO] workbook still has pending ID(s); leaving without final save/reset: {0}" -f $name) -ForegroundColor Yellow
                try { $wb.Close($false) } catch {}
                $cntSkip++
                continue
            }

            Set-ReviewCursorAllSheets $wb $CursorCell
            Save-BySendKeys $shell $excel $wb $SaveWaitMs

            if (-not $wb.Saved) {
                Write-Host '  [WARN] Excel still reports unsaved after Ctrl+S/Esc/wait. Closing without second save to avoid macro loop.' -ForegroundColor Yellow
            }

            try { $wb.Close($false) } catch {}
            $wb = $null

            Write-Host ("  [OK] workbook reviewed/saved: {0} -> {1}" -f $name, $ReviewBit) -ForegroundColor Green
            $cntDone++
        } catch {
            Write-Host ("  [ERROR] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            try { if ($wb) { $wb.Close($false) } } catch {}
            $cntFail++
        }
    }
} finally {
    if ($shell) {
        try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($shell) } catch {}
    }
    if ($excel) {
        try { $excel.DisplayAlerts = $true } catch {}
        Close-ExcelApp $excel
    }
}

Write-Host ''
Write-Host '===== ReviewEvidence Done =====' -ForegroundColor Green
Write-Host ("  Done    : {0}" -f $cntDone)
Write-Host ("  Skipped : {0}" -f $cntSkip)
Write-Host ("  Missing : {0}" -f $cntMissing)
Write-Host ("  Failed  : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
Write-Host ("  Mapping : {0}" -f $mappingPath)
