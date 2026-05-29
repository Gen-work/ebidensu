#Requires -Version 5.1
param(
    [string]$WorkDir = '',
    [string]$Owner = '厳',
    [string[]]$TargetIds = @(),
    [switch]$Force
)
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'Common.ps1')
$cfg = Import-PowerShellDataFile (Join-Path $scriptDir 'VerifyConfig.psd1')

if (-not $WorkDir) { throw '-WorkDir is required' }
$mappingFile = Join-Path $WorkDir ($cfg.Paths.MappingPattern -f $Owner)
if (-not (Test-Path $mappingFile)) { throw "Mapping not found: $mappingFile" }

$rows = Import-Csv $mappingFile -Encoding UTF8BOM
if ($TargetIds.Count -gt 0) {
    $rows = @($rows | Where-Object {
        $_.Correl_ID_S -in $TargetIds -or $_.Correl_ID_M -in $TargetIds -or $_.JOB_NAME -in $TargetIds -or $_.Excel_NAME -in $TargetIds
    })
}

$pending = @($rows | Where-Object {
    ($Force -or -not $_.GFIX_log -or $_.GFIX_log -eq '0' -or $_.GFIX_log -eq '') -and
    ($_.isMiddle -eq '1' -or $_.isMiddle -eq 'true' -or $_.isMiddle -eq 'True')
})
if ($pending.Count -eq 0) {
    Write-Host '[GfixLogDownload] No pending rows.' -ForegroundColor Green
    exit 0
}

$logDir = Join-Path $WorkDir 'log'
Ensure-Dir $logDir
$downloadDir = Join-Path $env:USERPROFILE 'Downloads'

Write-Host "`n===== GFIX Step 4: GoAnywhere LOG Download =====" -ForegroundColor Green
Write-Host "Pending rows: $($pending.Count)" -ForegroundColor Cyan
Write-Host ''
Write-Host 'Prepare GoAnywhere page in Edge BEFORE starting:' -ForegroundColor Yellow
Write-Host '  - Filter BIZ code (example: JRV)'
Write-Host '  - Sort newer=up / older=down'
Write-Host '  - Keep focus on the list page'
Write-Host 'Then press Enter. (q to quit)' -ForegroundColor Magenta
$resp = Read-Host
if ($resp -eq 'q') { exit 0 }
Switch-ToEdge

foreach ($row in $pending) {
    $ifRaw = [string]$row.IF
    if ([string]::IsNullOrWhiteSpace($ifRaw)) { continue }

    $ifNorm = ($ifRaw -replace '-', '_').Trim()
    $correl = [string]$row.Correl_ID_S
    Write-Host "`n[$correl] IF search: $ifNorm" -ForegroundColor White

    Send-CtrlF
    Paste-Replace $ifNorm
    Send-Enter
    Send-Key '{ESC}' 300
    Send-ShiftTab 1

    Send-Enter
    Start-Sleep -Milliseconds 1200

    Send-CtrlF
    Paste-Replace 'ダウンロード'
    Send-Enter
    Send-Key '{ESC}' 200
    Send-Enter
    Start-Sleep -Milliseconds 1200

    $jobNum = [string]$row.JOB_NAME
    if ($jobNum -match '^\d{13}$') {
        $src = Join-Path $downloadDir ("{0}.log" -f $jobNum)
        $dst = Join-Path $logDir ("{0}.log" -f $jobNum)
        if (Test-Path -LiteralPath $src) {
            Move-Item -LiteralPath $src -Destination $dst -Force
            Write-Host "  moved: $jobNum.log" -ForegroundColor Green
        } else {
            Write-Host "  [WARN] not found in Downloads: $jobNum.log" -ForegroundColor Yellow
        }
    }

    $row.GFIX_log = '1'

    Send-Tab 2
    Send-Enter
    Start-Sleep -Milliseconds 800
}

$rows | Export-Csv $mappingFile -NoTypeInformation -Encoding UTF8BOM
Write-Host "`n[GfixLogDownload] Mapping saved + step finished." -ForegroundColor Green