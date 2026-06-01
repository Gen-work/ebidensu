# ============================================================
# DeliverMail.ps1 - Phase DeliverMail
#
# Sends one review-request mail per evidence Excel (grouped by Excel_NAME).
# For each group the script builds an Outlook *draft* (CreateItem + Display,
# NOT Send) so Misaki can eyeball it and click Send by hand. After she sends,
# she returns to this shell and presses Enter; only then is the group's
# isDelivered flag set to 1. A  -m "comment"  can be appended at the prompt
# (recorded per Excel_NAME in the DeliverComment column, like ReviewComment).
#
# Mail is never sent automatically -- the operator always clicks Send.
# Outlook is never Quit (it may be the operator's running session); only the
# COM reference is released at the end.
#
# All Japanese (subject template, body lines, reviewer name) arrives as
# parameters from VerifyConfig.psd1 (which carries a BOM) so this source
# stays pure ASCII per the project encoding policy.
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner = ([char]0x53B3),
    [string[]]$TargetIds = @(),

    [string]$From = '',
    [string]$ReviewerAddress = '',
    [string]$ReviewerDisplay = '',
    [string]$ReviewerShort = '',

    [string]$Phase = '',
    [string]$SubjectTemplate = '{0} review ({1})',
    [string[]]$BodyLines = @(),
    [string]$EvidenceFolder = '',
    [string]$CheckSheetFolder = '',
    [string]$CheckSheetFile = '',

    [string]$EvidenceDir = '',

    [switch]$Force,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'

try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$forceFlag  = [bool]$Force.IsPresent
$dryRunFlag = [bool]$DryRun.IsPresent

# -- Dot-source shared libs (all param()-free) ---------------
. (Join-Path $PSScriptRoot 'MappingStore.ps1')
. (Join-Path $PSScriptRoot 'ProgressLog.ps1')
. (Join-Path $PSScriptRoot 'WorkbookResolver.ps1')

# Splits raw input into an action ('' = sent/mark, 's' = skip, 'q' = quit)
# and an optional comment introduced by  -m "comment". Mirrors the Review flow.
function Parse-DeliverInput([string]$Raw) {
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

$To = $ReviewerAddress
if ([string]::IsNullOrWhiteSpace($To)) { $To = $ReviewerDisplay }
if ([string]::IsNullOrWhiteSpace($To)) {
    Write-Host '[ERROR] reviewer address/display not configured (Reviewer.Address in VerifyConfig.psd1).' -ForegroundColor Red
    exit 1
}

$targets = @(ConvertTo-TargetIdList $TargetIds)

$allRows = @(Import-Mapping $mappingPath)
if ($allRows.Count -eq 0) {
    Write-Host "[ERROR] mapping has no rows: $mappingPath" -ForegroundColor Red; exit 1
}
Ensure-MappingColumns $allRows | Out-Null   # guarantees isDelivered + DeliverComment

# Group by Excel_NAME (mapping order, deduped) honoring the target filter.
$names      = New-Object System.Collections.Generic.List[string]
$prefixByName = @{}
foreach ($r in $allRows) {
    if (-not (Test-TargetRow $r $targets)) { continue }
    $name = [string]$r.Excel_NAME
    if ([string]::IsNullOrWhiteSpace($name)) { continue }
    if (-not $names.Contains($name)) {
        $names.Add($name)
        $prefixByName[$name] = Get-RowProp $r 'Excel_Prefix'
    }
}

if ($names.Count -eq 0) {
    Write-Host '[INFO] no target rows.' -ForegroundColor Yellow
    return
}

$bodyFmt = ($BodyLines -join "`r`n")

Write-Host ''
Write-Host '===== DeliverMail =====' -ForegroundColor Green
Write-Host ("  WorkDir     : {0}" -f $WorkDir)
Write-Host ("  Mapping     : {0}" -f $mappingPath)
Write-Host ("  EvidenceDir : {0}" -f $EvidenceDir)
Write-Host ("  Workbooks   : {0}" -f $names.Count)
Write-Host ("  To          : {0}" -f $To)
if (-not [string]::IsNullOrWhiteSpace($From)) { Write-Host ("  From (set)  : {0}" -f $From) }
Write-Host ("  Force       : {0}" -f $forceFlag)
if ($targets.Count -gt 0) { Write-Host ("  TargetIds   : {0}" -f ($targets -join ', ')) }

if ($dryRunFlag) {
    foreach ($name in $names) {
        $subject = $SubjectTemplate -f $Phase, $name
        Write-Host ("  [DRY] {0}  subject: {1}" -f $name, $subject)
    }
    return
}

$outlook = $null
$cntDone = 0; $cntSkip = 0; $cntFail = 0

try {
    try {
        $outlook = New-Object -ComObject Outlook.Application
    } catch {
        Write-Host ("[ERROR] could not start Outlook: {0}" -f $_.Exception.Message) -ForegroundColor Red
        exit 1
    }

    for ($idx = 0; $idx -lt $names.Count; $idx++) {
        $name = $names[$idx]
        $groupRows = @($allRows | Where-Object { [string]$_.Excel_NAME -eq $name })

        $alreadyDone = $true
        foreach ($r in $groupRows) {
            if (-not (Test-SnapDone (Get-RowProp $r 'isDelivered'))) { $alreadyDone = $false; break }
        }
        if ($alreadyDone -and -not $forceFlag) {
            Write-Host ("[{0}/{1}] SKIP delivered: {2}" -f ($idx + 1), $names.Count, $name) -ForegroundColor DarkGray
            $cntSkip++
            continue
        }

        $fullStem     = Get-ExcelFullStem -Prefix ($prefixByName[$name]) -Name $name
        $evidenceFile = (Get-ExcelDestLeaf $fullStem)   # <stem>.xlsx
        $subject      = $SubjectTemplate -f $Phase, $name
        $body         = $bodyFmt -f $ReviewerShort, $Owner, $EvidenceFolder, $evidenceFile, $CheckSheetFolder, $CheckSheetFile

        # Non-fatal: warn if the evidence workbook is not where we expect.
        $onDisk = Find-WorkbookByExcelName -Dir $EvidenceDir -ExcelName $fullStem
        if ($null -eq $onDisk) {
            Write-Host ("  [WARN] evidence workbook not found under {0}: {1}" -f $EvidenceDir, $evidenceFile) -ForegroundColor Yellow
        }

        $mail = $null
        try {
            Write-Host ''
            Write-Host ("[{0}/{1}] DRAFT: {2}" -f ($idx + 1), $names.Count, $name) -ForegroundColor Cyan
            Write-Host ("  Subject : {0}" -f $subject)
            Write-Host ("  File    : {0}" -f $evidenceFile)

            $mail = $outlook.CreateItem(0)   # olMailItem
            $mail.To = $To
            $mail.Subject = $subject
            $mail.Body = $body
            if (-not [string]::IsNullOrWhiteSpace($From)) {
                try { $mail.SentOnBehalfOfName = $From } catch {}
            }
            $mail.Display($false)   # open draft window (non-modal); operator clicks Send

            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverMail' -JobName $name -Action 'draft' -Status 'info' -Message $subject

            Write-Host '  Draft opened in Outlook. Send it by hand, then return here.' -ForegroundColor DarkGray
            $raw    = Read-Host '  Enter=sent (mark delivered), s=skip, q=quit   ( add  -m "comment"  to record a note )'
            $choice = Parse-DeliverInput $raw

            if (-not [string]::IsNullOrWhiteSpace($choice.Comment)) {
                foreach ($r in $groupRows) {
                    if (-not ($r.PSObject.Properties.Name -contains 'DeliverComment')) {
                        $r | Add-Member -NotePropertyName 'DeliverComment' -NotePropertyValue '' -Force
                    }
                    $r.DeliverComment = $choice.Comment
                }
                Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
                Write-Host ("  [COMMENT] recorded: {0}" -f $choice.Comment) -ForegroundColor DarkCyan
            }

            if ($choice.Action -eq 'q') {
                Write-Host '  [QUIT] stopping at user request.' -ForegroundColor Yellow
                break
            }
            if ($choice.Action -eq 's') {
                Write-Host ("  [SKIP] not marked delivered: {0}" -f $name) -ForegroundColor DarkGray
                Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverMail' -JobName $name -Action 'send' -Status 'skip' -Message ''
                $cntSkip++
                continue
            }

            foreach ($r in $groupRows) {
                if (-not ($r.PSObject.Properties.Name -contains 'isDelivered')) {
                    $r | Add-Member -NotePropertyName 'isDelivered' -NotePropertyValue '0' -Force
                }
                $r.isDelivered = '1'
            }
            Export-MappingAtomic -Rows $allRows -Path $mappingPath | Out-Null
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverMail' -JobName $name -Action 'send' -Status 'ok' -Message $subject
            Write-Host ("  [OK] delivered: {0}" -f $name) -ForegroundColor Green
            $cntDone++
        } catch {
            Write-Host ("  [ERROR] {0}: {1}" -f $name, $_.Exception.Message) -ForegroundColor Red
            Write-ProgressEvent -WorkDir $WorkDir -Phase 'DeliverMail' -JobName $name -Action 'send' -Status 'fail' -Message $_.Exception.Message
            $cntFail++
        } finally {
            if ($mail) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($mail) } catch {} }
        }
    }
} finally {
    # Never Quit Outlook -- it may be the operator's live session. Just release.
    if ($outlook) { try { [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($outlook) } catch {} }
    [System.GC]::Collect()
    [System.GC]::WaitForPendingFinalizers()
}

Write-Host ''
Write-Host '===== DeliverMail Done =====' -ForegroundColor Green
Write-Host ("  Delivered : {0}" -f $cntDone)
Write-Host ("  Skipped   : {0}" -f $cntSkip)
Write-Host ("  Failed    : {0}" -f $cntFail) -ForegroundColor $(if ($cntFail -gt 0) { 'Yellow' } else { 'White' })
Write-Host ("  Mapping   : {0}" -f $mappingPath)
