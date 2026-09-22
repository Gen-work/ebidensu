#Requires -Version 5.1
# ============================================================
#  GiftMqProcessTime.ps1   (standalone; NOT a VerifyTool phase)
#  -- UTF-8, NO BOM, ASCII source.
#
#  Fills the shori-jikan workbook (<label>(<Tag>).xlsx: No. / GIFT-GFIX /
#  correl / start / end / duration / count / job) from three sources the
#  operator used to read by hand every day:
#
#    START  = the GIFT MQ "Transfer status inquiry results" LIST page:
#             each record's Send date. Captured once with Ctrl+A/Ctrl+C
#             (Read-PageText.ps1) or read from a saved text file.
#    END    = that record's Detail page INSERTDATETIME (the receive-log
#             insertion stamp). Reached per record by keyboard: Ctrl+F
#             the record's Send date text -> Esc -> Tab (the row's
#             Detail button is the next focusable) -> Enter; or, with
#             -DetailNav Tab, Ctrl+F the page title -> Esc -> Tab N times
#             (N = record No, the Detail buttons are the only focusable
#             controls in the list). EVERY detail page reached is checked
#             against the record (CORRELID(CHAR) + SENDDATETIME) before
#             its INSERTDATETIME is trusted; a mismatch falls back to the
#             other navigation (with -TabsBeforeFirstDetail auto-tried
#             0/1/2) and finally asks the operator.
#    COUNT  = only in the team's Teams chat. Paste the chat text into a
#             file and pass -TeamsTextFile; every '<jobu>:<W-name> ...
#             (<soushin-yotei>:N<ken>)' line is picked up. Without it the
#             count cell stays blank and the end-of-run summary lists the
#             jobs still needing one, so the operator types only those.
#
#  WHICH job a list record belongs to comes from the operator's own
#  mapping.xlsx (sheet 'mapping'): JOB, owner, GIFT run date, GIFT TIME
#  (the leader's daily schedule). The scheduled time is a window (+-
#  ToleranceMinutes, default 10) and the day's order breaks ties
#  (Resolve-GiftMqJobMatches, modules/verify/GiftMqProcessTime.ps1).
#
#  Output rows: an existing GIFT row for the job (col H == job, col B ==
#  GIFT) is UPDATED, blank cells only (-Force rewrites); otherwise a
#  GIFT+GFIX pair is appended (into the sheet's spare placeholder rows
#  when it has them, else after the last row, formats copied). Start/end
#  are written as real Excel date/time values (yyyy/mm/dd hh:mm:ss),
#  duration stays the sheet's =E-D formula, count as '<n><ken>' text.
#  The GFIX side is never touched by this tool.
#
#  Everything pure is in modules/verify/GiftMqProcessTime.ps1 and unit-
#  tested; this file is COM + SendKeys glue (static-checked only in the
#  dev environment -- confirm on an office PC, see docs/GiftMqProcessTime.md).
#
#  Usage (office PC, Edge showing the LIST page for the wanted dates):
#    powershell -File GiftMqProcessTime.ps1 -MappingXlsx C:\work\mapping.xlsx `
#        -OutputXlsx "C:\work\<label>(BIX).xlsx" -FromDate 2026-09-15
#    add -TeamsTextFile C:\work\teams.txt for counts,
#        -PageTextFile C:\work\giftmq_text\list_....txt to reuse a capture,
#        -NoDetail to fill start times only, -DryRun to write nothing.
# ============================================================
param(
    [Parameter(Mandatory = $true)][string]$MappingXlsx,
    [string]$MappingSheet = 'mapping',
    [Parameter(Mandatory = $true)][string]$OutputXlsx,
    # '' = the workbook's first worksheet.
    [string]$OutputSheet = '',

    # Reuse a saved LIST-page Ctrl+A text instead of capturing it from Edge.
    [string]$PageTextFile = '',
    # Where captured page texts are archived ('' = <OutputXlsx dir>\giftmq_text).
    [string]$ArchiveDir = '',
    # Pasted Teams chat text (counts). Optional.
    [string]$TeamsTextFile = '',

    # Schedule filters: owner cell value, explicit JOB names, run-date window.
    [string]$Owner = '',
    [string[]]$Jobs = @(),
    [string]$FromDate = '',
    [string]$ToDate = '',
    # Only list records with this Correlid are considered ('' = any).
    [string]$CorrelId = '',
    [int]$ToleranceMinutes = 10,

    # Detail-page navigation (see header). 'Find' = Ctrl+F the record's own
    # Send date; 'Tab' = Ctrl+F the title then Tab No times.
    [ValidateSet('Find', 'Tab')][string]$DetailNav = 'Find',
    # Focusable controls before the first Detail button in Tab mode
    # (-1 = try 0, 1, 2 on the first record and lock in what verifies).
    [int]$TabsBeforeFirstDetail = -1,
    # How to leave a detail page: the page's own Back button (first
    # focusable after the title) or the browser's Alt+Left.
    [ValidateSet('Button', 'AltLeft')][string]$BackMethod = 'Button',
    # Ctrl+F focus anchor: the title both pages carry.
    [string]$AnchorText = 'Transfer status inquiry results',
    # Detail field that is the END time.
    [string]$EndTimeKey = 'INSERTDATETIME',

    [switch]$NoDetail,
    [switch]$Force,
    [switch]$DryRun,
    [switch]$NonInteractive,

    [int]$ActionWaitMs = 500,
    [int]$ResultWaitSec = 2,
    [int]$PollTimeoutSec = 15,
    [string]$CommonScript = ''
)

$ErrorActionPreference = 'Stop'
# NOTE: do NOT force [Console]::OutputEncoding to UTF-8 here. The office PC's
# console runs the JP codepage (932); forcing UTF-8 output there makes every
# non-ASCII byte we print render as mojibake -- the first real run showed
# the output workbook's own path as 'C:\...\<garbage>BIX.xlsx', which reads
# like the tool opened the wrong file when it had in fact opened the right
# one. The console's own encoding already renders a [char]-built Japanese
# string correctly on both a CP932 and a UTF-8 console, so leave it alone.

# Capture switches BEFORE any dot-source (CLAUDE.md switch-flag pattern).
$noDetailFlag = [bool]$NoDetail.IsPresent
$forceFlag    = [bool]$Force.IsPresent
$dryRunFlag   = [bool]$DryRun.IsPresent
$interactive  = -not [bool]$NonInteractive.IsPresent

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
if ([string]::IsNullOrWhiteSpace($CommonScript)) { $CommonScript = Join-Path $scriptDir 'Common.ps1' }
if (-not (Test-Path -LiteralPath $CommonScript)) { Write-Host "[ERROR] Common.ps1 not found: $CommonScript" -ForegroundColor Red; exit 1 }

$savedEAP = $ErrorActionPreference
$ErrorActionPreference = 'SilentlyContinue'
. $CommonScript
$ErrorActionPreference = $savedEAP
. (Join-Path $scriptDir 'ExcelHelpers.ps1')
. (Join-Path $scriptDir 'modules/verify/GiftMqProcessTime.ps1')
$pageTextScript = Join-Path $scriptDir 'Read-PageText.ps1'

if (-not (Get-Command -Name 'Wait-PagePrepared' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] Common.ps1 dot-source failed (Wait-PagePrepared not found).' -ForegroundColor Red; exit 1
}
if (-not (Get-Command -Name 'Resolve-GiftMqJobMatches' -ErrorAction SilentlyContinue)) {
    Write-Host '[ERROR] modules/verify/GiftMqProcessTime.ps1 dot-source failed.' -ForegroundColor Red; exit 1
}

$Global:Timing = @{
    ActionWaitMs  = $ActionWaitMs
    ResultWaitSec = $ResultWaitSec
    ResultWaitMs  = [Math]::Max(200, $ResultWaitSec * 1000)
}
$L = Get-GiftMqLabels

# ============================================================
# Console / Edge helpers (mirror MqSnap.ps1's local ones)
# ============================================================

# Extra P/Invokes this script needs on top of Common.ps1's WinAPI/MouseAPI:
# which process owns the foreground window (so a click can never land on the
# wrong app) and the console's QuickEdit flag (see Disable-GiftMqQuickEdit).
# Guarded: re-running in the same session would throw "type already exists".
if (-not ('GiftMqWin' -as [type])) {
    Add-Type @"
using System;
using System.Runtime.InteropServices;

public class GiftMqWin {
    [DllImport("user32.dll")]
    public static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

    [DllImport("kernel32.dll")]
    public static extern IntPtr GetStdHandle(int nStdHandle);

    [DllImport("kernel32.dll")]
    public static extern bool GetConsoleMode(IntPtr hConsoleHandle, out uint lpMode);

    [DllImport("kernel32.dll")]
    public static extern bool SetConsoleMode(IntPtr hConsoleHandle, uint dwMode);
}
"@
}

# QuickEdit mode (on by default on Windows 10/11) turns any left-click inside
# the console into a text selection, and a console with an active selection
# BLOCKS the process's writes until the selection is cleared. This script
# drives a real mouse, so one click landing on the console instead of Edge
# used to freeze the run dead: the operator pressed Enter and nothing ever
# happened again, with no error to show for it. Turning QuickEdit off for the
# duration removes the freeze; Click-GiftMqPageCenter's foreground check below
# removes the stray click itself. Both, because either alone still leaves a
# silent failure mode.
$script:PrevConsoleMode = $null
function Disable-GiftMqQuickEdit {
    try {
        $h = [GiftMqWin]::GetStdHandle(-10)   # STD_INPUT_HANDLE
        if ($h -eq [IntPtr]::Zero) { return }
        $mode = 0
        if (-not [GiftMqWin]::GetConsoleMode($h, [ref]$mode)) { return }
        $script:PrevConsoleMode = $mode
        # clear ENABLE_QUICK_EDIT_MODE (0x0040), set ENABLE_EXTENDED_FLAGS (0x0080)
        $new = ($mode -band (-bnot 0x0040)) -bor 0x0080
        [void][GiftMqWin]::SetConsoleMode($h, $new)
    } catch {}
}

function Restore-GiftMqQuickEdit {
    try {
        if ($null -eq $script:PrevConsoleMode) { return }
        $h = [GiftMqWin]::GetStdHandle(-10)
        if ($h -eq [IntPtr]::Zero) { return }
        [void][GiftMqWin]::SetConsoleMode($h, [uint32]$script:PrevConsoleMode)
        $script:PrevConsoleMode = $null
    } catch {}
}

# Is the window the operator is looking at an Edge window? Everything this
# script does with the mouse and the keyboard is aimed at the MQ page, so
# anything else in the foreground means the action would hit the wrong app.
function Test-GiftMqForegroundIsEdge {
    try {
        $hWnd = [WinAPI]::GetForegroundWindow()
        if ($hWnd -eq [IntPtr]::Zero) { return $false }
        $pid32 = 0
        [void][GiftMqWin]::GetWindowThreadProcessId($hWnd, [ref]$pid32)
        if ($pid32 -eq 0) { return $false }
        $p = Get-Process -Id ([int]$pid32) -ErrorAction SilentlyContinue
        if ($null -eq $p) { return $false }
        return ($p.ProcessName -eq 'msedge')
    } catch { return $false }
}

function Bring-ShellToFront {
    try {
        $hwnd = (Get-Process -Id $PID).MainWindowHandle
        if ($hwnd -ne [IntPtr]::Zero) {
            [WinAPI]::ShowWindowAsync($hwnd, 9) | Out-Null
            [WinAPI]::SetForegroundWindow($hwnd) | Out-Null
            Start-Sleep -Milliseconds 200
        }
    } catch {}
}

# Click the centre of the Edge window: on the frameset MQ page that lands in
# frame_main, so Ctrl+A / Ctrl+F act on the result frame and not on the left
# navigation frame.
#
# It clicks ONLY when Edge is the foreground window. It used to click whatever
# was foreground, which meant that whenever the Alt+Tab in Switch-ToEdge did
# not land (synthetic Alt+Tab is unreliable and is ignored outright under some
# policies), the click went into this console instead -- selecting text there
# and freezing every later write. Returns $true when the click was made.
function Click-GiftMqPageCenter {
    if (-not (Test-GiftMqForegroundIsEdge)) { return $false }
    $hWnd = [WinAPI]::GetForegroundWindow()
    if ($hWnd -eq [IntPtr]::Zero) { return $false }
    $rect = New-Object WinAPI+RECT
    [WinAPI]::GetWindowRect($hWnd, [ref]$rect) | Out-Null
    $x = [int](($rect.Left + $rect.Right) / 2)
    $y = [int](($rect.Top + $rect.Bottom) / 2)
    [MouseAPI]::SetCursorPos($x, $y) | Out-Null
    Start-Sleep -Milliseconds 100
    [MouseAPI]::mouse_event(0x0002, 0, 0, 0, [UIntPtr]::Zero)  # LEFTDOWN
    Start-Sleep -Milliseconds 50
    [MouseAPI]::mouse_event(0x0004, 0, 0, 0, [UIntPtr]::Zero)  # LEFTUP
    Start-Sleep -Milliseconds 300
    return $true
}

# Put Edge in front and say so. Switch-ToEdge's Alt+Tab is best-effort, so the
# result is CHECKED rather than assumed; when it fails the operator is asked to
# click the Edge window, because a run that proceeds without the page in front
# just sends keystrokes into whatever else is there.
function Confirm-GiftMqEdgeForeground {
    param([string]$What = 'the GIFT MQ page')
    for ($i = 0; $i -lt 3; $i++) {
        Switch-ToEdge
        if (Test-GiftMqForegroundIsEdge) { return $true }
        if (-not $interactive) { break }
        Bring-ShellToFront
        Write-Host ("  [WARN] Edge is not in the foreground, so nothing can be read from {0}." -f $What) -ForegroundColor Yellow
        Write-Host '    Click the Edge window yourself, then come back here: Enter=retry / q=quit : ' -ForegroundColor Magenta -NoNewline
        if ((Read-Host).Trim() -eq 'q') { return $false }
    }
    return (Test-GiftMqForegroundIsEdge)
}

function Read-GiftMqPageText {
    if (-not (Click-GiftMqPageCenter)) { return '' }
    $txt = & $pageTextScript -SelectWaitMs $ActionWaitMs -CopyWaitMs $ActionWaitMs
    if ($null -eq $txt) { return '' }
    return [string]$txt
}

# Poll the page text until $Test returns $true or the timeout elapses.
# Returns @{ Text; Ok }.
# Poll the page text until $Test passes or the timeout elapses. Every attempt
# prints one line: a poll that says nothing for 15 seconds is indistinguishable
# from a hang, which is exactly how the first freeze presented itself.
function Wait-GiftMqPageText {
    param([scriptblock]$Test, [int]$TimeoutSec, [string]$What = 'page')
    $deadline = (Get-Date).AddSeconds([Math]::Max(1, $TimeoutSec))
    $text = ''
    $n = 0
    do {
        $n++
        if (-not (Test-GiftMqForegroundIsEdge)) {
            Write-Host ("    read {0} (try {1}): Edge is not in the foreground" -f $What, $n) -ForegroundColor DarkYellow
        } else {
            $text = Read-GiftMqPageText
            if (& $Test $text) {
                Write-Host ("    read {0} (try {1}): {2} chars, recognised" -f $What, $n, $text.Length) -ForegroundColor DarkGray
                return @{ Text = $text; Ok = $true }
            }
            Write-Host ("    read {0} (try {1}): {2} chars, not the page yet" -f $What, $n, $text.Length) -ForegroundColor DarkGray
        }
        Start-Sleep -Milliseconds 700
    } while ((Get-Date) -lt $deadline)
    return @{ Text = $text; Ok = $false }
}

# Ctrl+F <text>, wait, Esc. Edge leaves the sequential-focus starting point
# at the match, so the next Tab lands on the first focusable AFTER it.
function Invoke-GiftMqFind {
    param([string]$Text)
    Send-CtrlF
    Paste-Replace $Text
    Start-Sleep -Milliseconds $Global:Timing.ResultWaitMs
    Send-Key '{ESC}' 300
}

function Test-GiftMqIsListText([string]$t) { return ($t -match 'Number of records') }
function Test-GiftMqIsDetailText([string]$t) { return ($t -match 'CORRELID\(CHAR\)' -or $t -match $EndTimeKey) }

function Save-GiftMqArchive {
    param([string]$Name, [string]$Text)
    if ([string]::IsNullOrWhiteSpace($script:archiveDir)) { return '' }
    try {
        Ensure-Dir $script:archiveDir
        $p = Join-Path $script:archiveDir $Name
        [System.IO.File]::WriteAllText($p, $Text, (New-Object System.Text.UTF8Encoding($false)))
        return $p
    } catch {
        Write-Host ("    [WARN] archive failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow
        return ''
    }
}

# Open record $Record's detail page. Returns @{ Ok; Detail; Text; Mode }.
# Never trusts the page it lands on: Test-GiftMqDetailMatchesRecord decides.
function Open-GiftMqDetail {
    param($Record, [string]$Mode, [int]$TabsBefore)
    Click-GiftMqPageCenter
    if ($Mode -eq 'Find') {
        Invoke-GiftMqFind $Record.SendDateText
        Send-Tab 1
    } else {
        Invoke-GiftMqFind $AnchorText
        Send-Tab (Get-GiftMqDetailTabCount -No ([int]$Record.No) -TabsBeforeFirstDetail $TabsBefore)
    }
    Send-Enter
    Start-Sleep -Seconds $ResultWaitSec
    $r = Wait-GiftMqPageText -Test { param($t) Test-GiftMqIsDetailText $t } -TimeoutSec $PollTimeoutSec -What 'detail page'
    $detail = $null
    $ok = $false
    if ($r.Ok) {
        $detail = ConvertFrom-GiftMqDetailText -Text $r.Text
        $ok = Test-GiftMqDetailMatchesRecord -Detail $detail -Record $Record
    }
    return @{ Ok = $ok; Detail = $detail; Text = [string]$r.Text; Mode = $Mode; Reached = [bool]$r.Ok }
}

# Leave the detail page. Returns $true when the LIST page text is back.
function Return-GiftMqList {
    Click-GiftMqPageCenter
    if ($BackMethod -eq 'AltLeft') {
        Send-Key '%{LEFT}' 500
    } else {
        Invoke-GiftMqFind $AnchorText
        Send-Tab 1
        Send-Enter
    }
    Start-Sleep -Seconds $ResultWaitSec
    $r = Wait-GiftMqPageText -Test { param($t) Test-GiftMqIsListText $t } -TimeoutSec $PollTimeoutSec -What 'LIST page'
    return [bool]$r.Ok
}

# ============================================================
# Excel: read the operator's mapping sheet
# ============================================================
function Read-GiftMqMappingRows {
    param([string]$Path, [string]$SheetName)
    $excel = $null; $wb = $null
    $rows = New-Object System.Collections.ArrayList
    try {
        $excel = New-ExcelApp
        $wb = $excel.Workbooks.Open($Path, 0, $true)   # ReadOnly
        if ($null -eq $wb) { throw "Workbooks.Open returned null: $Path" }
        $ws = $null
        if (-not [string]::IsNullOrWhiteSpace($SheetName)) { $ws = Get-SheetByName $wb $SheetName }
        if ($null -eq $ws) { $ws = $wb.Worksheets.Item(1) }
        $used = $ws.UsedRange
        $vals = $used.Value2
        $firstRow = [int]$used.Row
        $nRows = [int]$used.Rows.Count
        $nCols = [int]$used.Columns.Count
        if ($nRows -lt 2) { return @($rows.ToArray()) }
        # header = the first used row
        $headers = @()
        for ($c = 1; $c -le $nCols; $c++) { $headers += [string]$vals.GetValue(1, $c) }
        $cols = Get-GiftMqMappingColumns -Headers $headers
        foreach ($k in @('Job', 'GiftDate')) {
            if ($cols[$k] -lt 1) { throw ("mapping header '{0}' not found on sheet '{1}' (row {2})" -f $k, $ws.Name, $firstRow) }
        }
        for ($r = 2; $r -le $nRows; $r++) {
            $get = {
                param($key)
                if ($cols[$key] -lt 1) { return $null }
                return $vals.GetValue($r, [int]$cols[$key])
            }
            [void]$rows.Add(@{
                Job      = [string](& $get 'Job')
                Excel    = [string](& $get 'Excel')
                Owner    = [string](& $get 'Owner')
                GiftDate = (& $get 'GiftDate')
                GiftTime = (& $get 'GiftTime')
                Row      = ($firstRow + $r - 1)
            })
        }
        return @($rows.ToArray())
    } finally {
        if ($null -ne $wb) { Close-Workbook $wb $false }
        if ($null -ne $excel) { Close-ExcelApp $excel }
    }
}

# ============================================================
# Excel: output workbook read / write
# ============================================================
function Get-GiftMqSheetLastRow($ws) {
    $used = $ws.UsedRange
    return ([int]$used.Row + [int]$used.Rows.Count - 1)
}

function Read-GiftMqSheetRows($ws) {
    $last = Get-GiftMqSheetLastRow $ws
    $rows = New-Object System.Collections.ArrayList
    if ($last -lt 2) { return @($rows.ToArray()) }
    $vals = $ws.Range(('A2:H{0}' -f $last)).Value2
    for ($r = 2; $r -le $last; $r++) {
        $i = $r - 1
        [void]$rows.Add([PSCustomObject]@{
            Row   = $r
            Side  = [string]$vals.GetValue($i, 2)
            Job   = [string]$vals.GetValue($i, 8)
            Start = $vals.GetValue($i, 4)
            End   = $vals.GetValue($i, 5)
            Count = [string]$vals.GetValue($i, 7)
            Correl = [string]$vals.GetValue($i, 3)
        })
    }
    return @($rows.ToArray())
}

# First spare placeholder pair (rows r, r+1 with C/D/E/G/H all blank),
# else 0.
function Find-GiftMqPlaceholderPair {
    param([object[]]$SheetRows, [int]$NotBelow = 2)
    $byRow = @{}
    foreach ($sr in $SheetRows) { $byRow[[int]$sr.Row] = $sr }
    $blank = {
        param($sr)
        return ([string]::IsNullOrWhiteSpace([string]$sr.Job) -and [string]::IsNullOrWhiteSpace([string]$sr.Correl) -and
                [string]::IsNullOrWhiteSpace([string]$sr.Start) -and [string]::IsNullOrWhiteSpace([string]$sr.End) -and
                [string]::IsNullOrWhiteSpace([string]$sr.Count))
    }
    foreach ($sr in ($SheetRows | Sort-Object Row)) {
        $r = [int]$sr.Row
        if ($r -lt $NotBelow) { continue }
        if (-not (& $blank $sr)) { continue }
        if ($byRow.ContainsKey($r + 1) -and (& $blank $byRow[$r + 1])) { return $r }
    }
    return 0
}

function Write-GiftMqDateCell($cell, $dt) {
    if ($null -eq $dt) { return }
    try { $cell.NumberFormat = 'yyyy/mm/dd hh:mm:ss' } catch {}
    Set-RangeValue2 $cell ([double]([datetime]$dt).ToOADate()) | Out-Null
}

# ============================================================
# 1. mapping -> schedules
# ============================================================
Write-Host ''
Write-Host '===== GiftMqProcessTime =====' -ForegroundColor Cyan
Write-Host ("  mapping : {0} [{1}]" -f $MappingXlsx, $MappingSheet)
Write-Host ("  output  : {0}" -f $OutputXlsx)
Write-Host ("  detail  : {0}{1}   back: {2}   tolerance: +-{3} min" -f $DetailNav, $(if ($noDetailFlag) { ' (skipped: -NoDetail)' } else { '' }), $BackMethod, $ToleranceMinutes)
Write-Host ("  force   : {0}   dry-run : {1}   interactive : {2}" -f $forceFlag, $dryRunFlag, $interactive)

# From here on a stray click can no longer freeze the console. Every exit
# below restores the flag; the engine-exit handler is the backstop for the
# paths that leave through Common.ps1's own `exit` (Wait-PagePrepared's 'q').
Disable-GiftMqQuickEdit
try { [void](Register-EngineEvent -SourceIdentifier ([System.Management.Automation.PsEngineEvent]::Exiting) -Action { Restore-GiftMqQuickEdit }) } catch {}

function Exit-GiftMq {
    param([int]$Code)
    Restore-GiftMqQuickEdit
    exit $Code
}

if (-not (Test-Path -LiteralPath $MappingXlsx)) { Write-Host "[ERROR] mapping workbook not found: $MappingXlsx" -ForegroundColor Red; exit 1 }
if (-not (Test-Path -LiteralPath $OutputXlsx))  { Write-Host "[ERROR] output workbook not found: $OutputXlsx" -ForegroundColor Red; exit 1 }
if (-not [string]::IsNullOrWhiteSpace($TeamsTextFile) -and -not (Test-Path -LiteralPath $TeamsTextFile)) {
    # Checked HERE, not where the counts are read: that happens AFTER the whole
    # detail-click loop, so a typo in the path used to surface a warning only
    # once the operator had already spent the run's slowest minutes.
    Write-Host ("[WARN] Teams text file not found: {0}" -f $TeamsTextFile) -ForegroundColor Yellow
    Write-Host '       Counts will stay blank; the end-of-run summary lists the jobs needing one.' -ForegroundColor Yellow
}
$script:archiveDir = $ArchiveDir
if ([string]::IsNullOrWhiteSpace($script:archiveDir)) {
    $script:archiveDir = Join-Path (Split-Path -Parent (Resolve-Path -LiteralPath $OutputXlsx).Path) 'giftmq_text'
}

$fromDt = $null; $toDt = $null
$inv = [System.Globalization.CultureInfo]::InvariantCulture
if (-not [string]::IsNullOrWhiteSpace($FromDate)) { $fromDt = [datetime]::Parse($FromDate, $inv) }
if (-not [string]::IsNullOrWhiteSpace($ToDate))   { $toDt   = [datetime]::Parse($ToDate, $inv) }

$mapRows = @(Read-GiftMqMappingRows -Path $MappingXlsx -SheetName $MappingSheet)
$schedules = @(ConvertTo-GiftMqSchedules -Rows $mapRows -Owner $Owner -Jobs $Jobs -FromDate $fromDt -ToDate $toDt)
Write-Host ("  mapping rows: {0}   scheduled (dated) jobs in scope: {1}" -f $mapRows.Count, $schedules.Count)
if ($schedules.Count -eq 0) {
    Write-Host '[INFO] No scheduled jobs in scope (check -FromDate/-ToDate/-Owner/-Jobs and the GIFT run-date column).' -ForegroundColor Yellow
    Exit-GiftMq 0
}
$schedMin = ($schedules | Sort-Object Scheduled | Select-Object -First 1).Scheduled.Date
$schedMax = ($schedules | Sort-Object Scheduled | Select-Object -Last 1).Scheduled.Date

# ============================================================
# 2. LIST page text (file or live)
# ============================================================
$listText = ''
$edgeLive = $false
if (-not [string]::IsNullOrWhiteSpace($PageTextFile)) {
    if (-not (Test-Path -LiteralPath $PageTextFile)) { Write-Host "[ERROR] page text file not found: $PageTextFile" -ForegroundColor Red; exit 1 }
    $listText = [System.IO.File]::ReadAllText($PageTextFile)
    Write-Host ("  list text : {0}" -f $PageTextFile)
} else {
    Bring-ShellToFront
    Write-Host ''
    Write-Host '  In Edge, open GIFT MQ > Transfer status > Inquiry and show the result LIST' -ForegroundColor Yellow
    Write-Host '  for the DAY(S) YOU WANT -- all their records on ONE page (raise rows-per-page).' -ForegroundColor Yellow
    if ($null -eq $fromDt -and $null -eq $toDt) {
        # Without an explicit window the page decides the scope, so printing the
        # mapping's whole dated span here read as "show me a month of records".
        Write-Host '  Only jobs scheduled on the days that page shows are matched. The mapping' -ForegroundColor Gray
        Write-Host ("  holds {0} dated job(s) between {1} and {2}; narrow with -FromDate/-ToDate." -f $schedules.Count, $schedMin.ToString('yyyy/MM/dd'), $schedMax.ToString('yyyy/MM/dd')) -ForegroundColor Gray
    } else {
        Write-Host ("  Scope: {0} job(s), {1} .. {2}." -f $schedules.Count, $schedMin.ToString('yyyy/MM/dd'), $schedMax.ToString('yyyy/MM/dd')) -ForegroundColor Gray
    }
    Write-Host '  Then click this console window again and press Enter (the answer is typed HERE,' -ForegroundColor Yellow
    Write-Host '  not in Edge). Do not touch the mouse or keyboard after that until it reports.' -ForegroundColor Yellow
    Wait-PagePrepared 'Press Enter when the LIST page is showing.'
    if (-not (Confirm-GiftMqEdgeForeground -What 'the LIST page')) {
        Write-Host '[ABORT] Edge never came to the foreground; nothing captured.' -ForegroundColor Yellow
        Exit-GiftMq 1
    }
    $edgeLive = $true
    $r = Wait-GiftMqPageText -Test { param($t) Test-GiftMqIsListText $t } -TimeoutSec $PollTimeoutSec -What 'LIST page'
    if (-not $r.Ok) {
        Write-Host '[ERROR] Could not read a LIST page (no "Number of records" in the Ctrl+A text).' -ForegroundColor Red
        if ([string]::IsNullOrWhiteSpace($r.Text)) {
            Write-Host '        Nothing was copied at all -- the Ctrl+A/Ctrl+C did not reach the page.' -ForegroundColor Red
        } else {
            $preview = if ($r.Text.Length -gt 160) { $r.Text.Substring(0, 160) } else { $r.Text }
            Write-Host ('        What was copied starts: ' + ($preview -replace '\s+', ' ')) -ForegroundColor Red
            Write-Host '        If that is the left navigation frame, click the result table once and rerun.' -ForegroundColor Red
        }
        Exit-GiftMq 1
    }
    $listText = [string]$r.Text
    $saved = Save-GiftMqArchive -Name ('list_{0}.txt' -f (Get-Date).ToString('yyyyMMdd_HHmmss')) -Text $listText
    if ($saved -ne '') { Write-Host ("  list text archived: {0}" -f $saved) -ForegroundColor DarkGray }
}

$list = ConvertFrom-GiftMqListText -Text $listText
Write-Host ("  page: Number of records {0}, parsed {1}" -f $list.NumRecords, $list.Records.Count)
if ($list.Records.Count -eq 0) { Write-Host '[ERROR] No records parsed from the LIST text.' -ForegroundColor Red; exit 1 }
if ($list.NumRecords -gt 0 -and $list.NumRecords -ne $list.Records.Count) {
    Write-Host ("  [WARN] page says {0} records but {1} parsed -- is the list split over pages?" -f $list.NumRecords, $list.Records.Count) -ForegroundColor Yellow
}

# No explicit window -> the page decides: only schedules on the days the
# page actually shows are in scope (older mapping rows would otherwise all
# report "no page record").
if ($null -eq $fromDt -and $null -eq $toDt) {
    $dated = @($list.Records | Where-Object { $null -ne $_.SendDate } | Sort-Object SendDate)
    if ($dated.Count -gt 0) {
        $pageMin = $dated[0].SendDate.Date
        $pageMax = $dated[$dated.Count - 1].SendDate.Date
        $schedules = @($schedules | Where-Object { $_.Scheduled.Date -ge $pageMin -and $_.Scheduled.Date -le $pageMax })
        Write-Host ("  scope narrowed to the page's days {0} .. {1}: {2} scheduled job(s)" -f $pageMin.ToString('yyyy/MM/dd'), $pageMax.ToString('yyyy/MM/dd'), $schedules.Count)
        if ($schedules.Count -eq 0) { Write-Host '[INFO] No scheduled jobs on the days this page shows.' -ForegroundColor Yellow; exit 0 }
        $schedMin = $pageMin; $schedMax = $pageMax
    }
}

# ============================================================
# 3. match
# ============================================================
$res = Resolve-GiftMqJobMatches -Schedules $schedules -Records $list.Records -ToleranceMinutes $ToleranceMinutes -CorrelId $CorrelId
Write-Host ''
Write-Host '  job       owner sched               status     No  send                 tmode reccnt' -ForegroundColor Gray
foreach ($m in $res.Matches) {
    $no = ''; $sd = ''; $tm = ''; $rc = ''
    if ($null -ne $m.Record) { $no = $m.Record.No; $sd = $m.Record.SendDateText; $tm = $m.Record.Tmode; $rc = $m.Record.RecCnt }
    $color = switch ($m.Status) { 'ok' { 'Green' } 'ambiguous' { 'Yellow' } default { 'DarkYellow' } }
    Write-Host ('  {0,-9} {1,-5} {2,-19} {3,-10} {4,3} {5,-20} {6,-5} {7}' -f $m.Job, $m.Owner, (Format-GiftMqStamp $m.Scheduled), $m.Status, $no, $sd, $tm, $rc) -ForegroundColor $color
}
if ($res.Unmatched.Count -gt 0) {
    Write-Host ''
    Write-Host ("  page records with no scheduled job ({0}):" -f $res.Unmatched.Count) -ForegroundColor DarkGray
    foreach ($u in ($res.Unmatched | Sort-Object SendDate)) {
        if ($u.SendDate.Date -lt $schedMin -or $u.SendDate.Date -gt $schedMax) { continue }
        Write-Host ('    No {0,3}  {1}  {2,-4} reccnt {3}' -f $u.No, $u.SendDateText, $u.Tmode, $u.RecCnt) -ForegroundColor DarkGray
    }
}

# Ambiguous / date-only schedules: let the operator pick.
foreach ($m in $res.Matches) {
    if ($m.Status -ne 'ambiguous' -and $m.Status -ne 'notime') { continue }
    if (-not $interactive) {
        if ($m.Status -eq 'notime') { Write-Host ("  [SKIP] {0}: date-only schedule, no record chosen (non-interactive)." -f $m.Job) -ForegroundColor Yellow }
        continue
    }
    Bring-ShellToFront
    Write-Host ''
    Write-Host ("  [{0}] {1} scheduled {2} -- candidates:" -f $m.Status.ToUpper(), $m.Job, (Format-GiftMqStamp $m.Scheduled)) -ForegroundColor Yellow
    foreach ($c in $m.Candidates) { Write-Host ('      No {0,3}  {1}  {2,-4} reccnt {3}' -f $c.No, $c.SendDateText, $c.Tmode, $c.RecCnt) }
    $default = if ($null -ne $m.Record) { 'Enter=No {0}' -f $m.Record.No } else { 'Enter=skip' }
    Write-Host ('    {0} / <No>=pick / s=skip : ' -f $default) -ForegroundColor Magenta -NoNewline
    $ans = (Read-Host).Trim()
    if ($ans -eq 's') { $m.Record = $null; $m.Status = 'skipped'; continue }
    if ($ans -match '^\d+$') {
        $pick = $m.Candidates | Where-Object { [int]$_.No -eq [int]$ans } | Select-Object -First 1
        if ($null -ne $pick) { $m.Record = $pick; $m.Status = 'ok' } else { Write-Host '    not a candidate; keeping the default.' -ForegroundColor Yellow }
    }
    if ($null -ne $m.Record -and $m.Status -eq 'ambiguous') { $m.Status = 'ok' }
}

# ============================================================
# 4. output sheet -> plan
# ============================================================
$excel = $null; $wb = $null; $ws = $null
$exitCode = 0
try {
    $excel = New-ExcelApp
    $wb = Open-Workbook $excel $OutputXlsx
    if ($null -eq $wb) { throw "Workbooks.Open returned null: $OutputXlsx" }
    if ([bool]$wb.ReadOnly) { throw "output workbook opened READ-ONLY (is it open in Excel?): $OutputXlsx" }
    if (-not [string]::IsNullOrWhiteSpace($OutputSheet)) { $ws = Get-SheetByName $wb $OutputSheet }
    if ($null -eq $ws) { $ws = $wb.Worksheets.Item(1) }
    $h8 = [string]$ws.Cells.Item(1, 8).Value2
    if (-not $h8.Contains($L.OutJob)) {
        Write-Host ("  [WARN] sheet '{0}' header H is '{1}', expected the job column -- writing A..H by position anyway." -f $ws.Name, $h8) -ForegroundColor Yellow
    }

    $sheetRows = @(Read-GiftMqSheetRows $ws)
    $plan = @(Get-GiftMqOutputPlan -SheetRows $sheetRows -Matches $res.Matches -Force:$forceFlag)
    Write-Host ''
    Write-Host ("  plan: {0} row(s) -- {1} update, {2} append" -f $plan.Count,
        @($plan | Where-Object { $_.Action -eq 'update' }).Count, @($plan | Where-Object { $_.Action -eq 'append' }).Count)

    # ============================================================
    # 5. detail pages -> end times
    # ============================================================
    $endTimes = @{}    # plan index -> [datetime]
    $detailWanted = @()
    for ($i = 0; $i -lt $plan.Count; $i++) { if ($plan[$i].NeedEnd) { $detailWanted += $i } }
    if ($noDetailFlag) {
        Write-Host ("  detail: skipped (-NoDetail); {0} row(s) keep a blank end time" -f $detailWanted.Count) -ForegroundColor DarkYellow
    } elseif ($detailWanted.Count -gt 0) {
        if (-not $edgeLive) {
            Bring-ShellToFront
            Write-Host ''
            Write-Host ("  {0} end time(s) to read. Show the same LIST page in Edge." -f $detailWanted.Count) -ForegroundColor Yellow
            Wait-PagePrepared 'Press Enter when the LIST page is showing (q=quit without end times).'
            if (-not (Confirm-GiftMqEdgeForeground -What 'the LIST page')) {
                Write-Host '  [WARN] Edge never came to the foreground; end times stay blank.' -ForegroundColor Yellow
                $detailWanted = @()
            }
            $edgeLive = $true
        }
        $tabsBefore = $TabsBeforeFirstDetail
        $mode = $DetailNav
        $n = 0
        foreach ($i in $detailWanted) {
            $n++
            $p = $plan[$i]
            $rec = $p.Match.Record
            Write-Host ("  [{0}/{1}] {2}  No {3}  {4}" -f $n, $detailWanted.Count, $p.Job, $rec.No, $rec.SendDateText) -ForegroundColor White
            $got = $null
            $attempt = 0
            $userQuit = $false
            do {
                $attempt++
                # Navigation attempts in order: the configured mode, then the other
                # mode; Tab mode auto-tries offsets 0/1/2 when not fixed.
                $tries = @()
                if ($mode -eq 'Find') {
                    $tries += @{ Mode = 'Find'; Tabs = 0 }
                    if ($tabsBefore -ge 0) { $tries += @{ Mode = 'Tab'; Tabs = $tabsBefore } } else { foreach ($k in 0, 1, 2) { $tries += @{ Mode = 'Tab'; Tabs = $k } } }
                } else {
                    if ($tabsBefore -ge 0) { $tries += @{ Mode = 'Tab'; Tabs = $tabsBefore } } else { foreach ($k in 0, 1, 2) { $tries += @{ Mode = 'Tab'; Tabs = $k } } }
                    $tries += @{ Mode = 'Find'; Tabs = 0 }
                }
                foreach ($t in $tries) {
                    $o = Open-GiftMqDetail -Record $rec -Mode $t.Mode -TabsBefore $t.Tabs
                    if ($o.Reached) {
                        [void](Save-GiftMqArchive -Name ('detail_{0:000}_{1}.txt' -f [int]$rec.No, $p.Job) -Text $o.Text)
                    }
                    if ($o.Ok) {
                        $got = $o
                        if ($t.Mode -eq 'Tab') { $tabsBefore = $t.Tabs }
                        $mode = $t.Mode
                        break
                    }
                    $why = if ($o.Reached) { 'a different record''s detail page' } else { 'no detail page' }
                    Write-Host ("    nav {0}{1}: {2}" -f $t.Mode, $(if ($t.Mode -eq 'Tab') { '(' + $t.Tabs + ')' } else { '' }), $why) -ForegroundColor DarkYellow
                    if ($o.Reached) {
                        if (-not (Return-GiftMqList)) { Write-Host '    [WARN] LIST page not back after Back.' -ForegroundColor Yellow }
                    } else {
                        # Wherever we are, try to get the list back before the next try.
                        $chk = Read-GiftMqPageText
                        if (-not (Test-GiftMqIsListText $chk)) { [void](Return-GiftMqList) }
                    }
                }
                if ($null -ne $got) { break }
                if (-not $interactive) { break }
                Bring-ShellToFront
                Write-Host ("    Could not reach No {0}'s detail page. Fix Edge (LIST page showing), then r=retry / s=skip / q=quit : " -f $rec.No) -ForegroundColor Magenta -NoNewline
                $ans = (Read-Host).Trim()
                if ($ans -eq 'q') { $userQuit = $true; break }
                if ($ans -eq 's') { break }
                [void](Confirm-GiftMqEdgeForeground -What 'the LIST page')
            } while ($attempt -lt 3)

            if ($userQuit) { Write-Host '[ABORT] User quit; rows read so far are still written.' -ForegroundColor Yellow; break }
            if ($null -eq $got) { Write-Host ("    -> end time NOT read for {0}" -f $p.Job) -ForegroundColor Yellow; continue }

            $endDt = Get-GiftMqDetailEndTime -Detail $got.Detail -Key $EndTimeKey
            if ($null -eq $endDt) {
                Write-Host ("    -> detail page has no usable {0}" -f $EndTimeKey) -ForegroundColor Yellow
            } else {
                $endTimes[$i] = $endDt
                Write-Host ("    -> {0} = {1}   (start {2}, {3:0}s)" -f $EndTimeKey, (Format-GiftMqStamp $endDt), $rec.SendDateText, ($endDt - $rec.SendDate).TotalSeconds) -ForegroundColor Green
            }
            if (-not (Return-GiftMqList)) {
                Write-Host '    [WARN] LIST page not back after Back; the next record may fail.' -ForegroundColor Yellow
            }
        }
    }

    # ============================================================
    # 6. Teams counts
    # ============================================================
    $counts = @{}
    if (-not [string]::IsNullOrWhiteSpace($TeamsTextFile)) {
        if (Test-Path -LiteralPath $TeamsTextFile) {
            $counts = Get-GiftMqTeamsCountMap -Text ([System.IO.File]::ReadAllText($TeamsTextFile))
            Write-Host ("  teams: {0} job count(s) read from {1}" -f $counts.Count, $TeamsTextFile)
        } else {
            Write-Host ("  [WARN] Teams text file not found: {0}" -f $TeamsTextFile) -ForegroundColor Yellow
        }
    }

    # ============================================================
    # 7. write
    # ============================================================
    Write-Host ''
    $written = 0
    $needCount = New-Object System.Collections.Generic.List[string]
    $needEnd   = New-Object System.Collections.Generic.List[string]
    $lastRow = Get-GiftMqSheetLastRow $ws
    $placeholder = Find-GiftMqPlaceholderPair -SheetRows $sheetRows
    $usedPlaceholders = @{}
    for ($i = 0; $i -lt $plan.Count; $i++) {
        $p = $plan[$i]
        $rec = $p.Match.Record
        $row = [int]$p.Row
        $isAppend = ($p.Action -eq 'append')
        if ($isAppend) {
            # next spare placeholder pair, else two fresh rows after the end
            $row = 0
            if ($placeholder -gt 0) {
                $cand = $placeholder
                while ($usedPlaceholders.ContainsKey($cand)) { $cand += 2 }
                $srCand = $sheetRows | Where-Object { [int]$_.Row -eq $cand } | Select-Object -First 1
                $srNext = $sheetRows | Where-Object { [int]$_.Row -eq ($cand + 1) } | Select-Object -First 1
                $blankPair = ($null -ne $srCand -and $null -ne $srNext -and
                    [string]::IsNullOrWhiteSpace([string]$srCand.Job) -and [string]::IsNullOrWhiteSpace([string]$srCand.Start) -and
                    [string]::IsNullOrWhiteSpace([string]$srNext.Job) -and [string]::IsNullOrWhiteSpace([string]$srNext.Start))
                if ($blankPair) { $row = $cand; $usedPlaceholders[$cand] = $true }
            }
            if ($row -eq 0) {
                $row = $lastRow + 1
                $lastRow += 2
                if (-not $dryRunFlag) {
                    # formats from the previous pair (GIFT row fill, borders, fonts)
                    try {
                        $ws.Range(('A{0}:H{1}' -f ($row - 2), ($row - 1))).Copy() | Out-Null
                        $ws.Range(('A{0}' -f $row)).PasteSpecial(-4122) | Out-Null   # xlPasteFormats
                        $excel.CutCopyMode = $false
                    } catch { Write-Host ("    [WARN] format copy failed: {0}" -f $_.Exception.Message) -ForegroundColor Yellow }
                }
            }
        }
        $startDt = $rec.SendDate
        $endDt = $null
        if ($endTimes.ContainsKey($i)) { $endDt = $endTimes[$i] }
        $cnt = ''
        if ($counts.ContainsKey($p.Job)) { $cnt = Format-GiftMqCount $counts[$p.Job] }

        $what = @()
        if ($p.NeedStart) { $what += 'start' }
        if ($p.NeedEnd -and $null -ne $endDt) { $what += 'end' }
        if ($p.NeedCount -and $cnt -ne '') { $what += 'count' }
        Write-Host ('  {0,-6} row {1,4}  {2,-9} {3,-20} {4,-20} {5,-8} [{6}]' -f $p.Action, $row, $p.Job, (Format-GiftMqStamp $startDt),
            $(if ($null -ne $endDt) { Format-GiftMqStamp $endDt } else { '' }), $cnt, ($what -join ',')) -ForegroundColor $(if ($dryRunFlag) { 'DarkGray' } else { 'White' })

        if ($p.NeedEnd -and $null -eq $endDt) { $needEnd.Add(('{0} (row {1}, No {2})' -f $p.Job, $row, $rec.No)) }
        if ($p.NeedCount -and $cnt -eq '') { $needCount.Add(('{0} (row {1})' -f $p.Job, $row)) }
        if ($dryRunFlag) { continue }

        if ($isAppend) {
            Set-RangeValue2 $ws.Cells.Item($row, 1) ($row - 1) | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row, 2) 'GIFT' | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row + 1, 1) $row | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row + 1, 2) 'GFIX' | Out-Null
            Set-RangeValue2 $ws.Cells.Item($row + 1, 8) $p.Job | Out-Null
            try {
                $ws.Cells.Item($row, 6).Formula = ('=E{0}-D{0}' -f $row)
                $ws.Cells.Item($row + 1, 6).Formula = ('=E{0}-D{0}' -f ($row + 1))
                $ws.Cells.Item($row, 6).NumberFormat = 'h:mm:ss'
                $ws.Cells.Item($row + 1, 6).NumberFormat = 'h:mm:ss'
            } catch {}
            Set-RangeValue2 $ws.Cells.Item($row, 8) $p.Job | Out-Null
        }
        if ($p.NeedStart) {
            Set-RangeValue2 $ws.Cells.Item($row, 3) ([string]$rec.CorrelId) | Out-Null
            Write-GiftMqDateCell $ws.Cells.Item($row, 4) $startDt
        }
        if ($p.NeedEnd -and $null -ne $endDt) { Write-GiftMqDateCell $ws.Cells.Item($row, 5) $endDt }
        if ($p.NeedCount -and $cnt -ne '') { Set-RangeValue2 $ws.Cells.Item($row, 7) $cnt | Out-Null }
        $written++
    }

    if (-not $dryRunFlag -and $written -gt 0) {
        $wb.Save()
        Write-Host ''
        Write-Host ("[OK] {0} row(s) written -> {1}" -f $written, $OutputXlsx) -ForegroundColor Green
    } elseif ($dryRunFlag) {
        Write-Host ''
        Write-Host '[DRY-RUN] nothing written.' -ForegroundColor Yellow
    } else {
        Write-Host ''
        Write-Host '[INFO] nothing to write.' -ForegroundColor Yellow
    }

    # ---- what the operator still has to do by hand ----
    if ($needEnd.Count -gt 0) {
        Write-Host ''
        Write-Host ("  End time still blank ({0}) -- open the Detail and copy {1}:" -f $needEnd.Count, $EndTimeKey) -ForegroundColor Yellow
        foreach ($s in $needEnd) { Write-Host ('    ' + $s) -ForegroundColor Yellow }
    }
    if ($needCount.Count -gt 0) {
        Write-Host ''
        Write-Host ("  Count still blank ({0}) -- read <{1}> from Teams and type it into col G:" -f $needCount.Count, $L.SendPlan) -ForegroundColor Yellow
        foreach ($s in $needCount) { Write-Host ('    ' + $s) -ForegroundColor Yellow }
    }
    $unm = @($res.Matches | Where-Object { $null -eq $_.Record })
    if ($unm.Count -gt 0) {
        Write-Host ''
        Write-Host ("  Scheduled jobs with NO page record ({0}):" -f $unm.Count) -ForegroundColor Yellow
        foreach ($m in $unm) { Write-Host ('    {0} scheduled {1} [{2}]' -f $m.Job, (Format-GiftMqStamp $m.Scheduled), $m.Status) -ForegroundColor Yellow }
    }
} catch {
    Write-Host ("[FAIL] {0}" -f $_.Exception.Message) -ForegroundColor Red
    $exitCode = 1
} finally {
    if ($null -ne $wb) { Close-Workbook $wb $false }
    if ($null -ne $excel) { Close-ExcelApp $excel }
    Restore-GiftMqQuickEdit
}
exit $exitCode
