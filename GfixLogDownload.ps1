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

$rows = @(Import-Csv $mappingFile -Encoding UTF8)
if ($TargetIds.Count -gt 0) {
    $rows = @($rows | Where-Object {
        $_.Correl_ID_S -in $TargetIds -or $_.Correl_ID_M -in $TargetIds -or $_.JOB_NAME -in $TargetIds -or $_.Excel_NAME -in $TargetIds
    })
}

$pending = @($rows | Where-Object {
    $Force -or -not $_.GFIX_log -or $_.GFIX_log -eq '0' -or $_.GFIX_log -eq ''
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
Write-Host '  *** Set max rows to 100 (default 20 is not enough!) ***' -ForegroundColor Red
Write-Host '  - Keep focus on the list page'
Write-Host 'Then press Enter. (q to quit)' -ForegroundColor Magenta
$resp = Read-Host
if ($resp -eq 'q') { exit 0 }
Switch-ToEdge
$hWnd = [WinAPI]::GetForegroundWindow()
[WinAPI]::ShowWindowAsync($hWnd, 3) | Out-Null  # SW_SHOWMAXIMIZED
Start-Sleep -Milliseconds 300

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
    $beforeDownload = Get-Date          # timestamp right before clicking Download
    Send-Enter
    Start-Sleep -Milliseconds 1800      # extra 600ms for download to land

    $jobNum = [string]$row.JOB_NAME

    # Find the .log file that appeared in Downloads after we clicked
    $newLog = Get-ChildItem -Path $downloadDir -Filter '*.log' -ErrorAction SilentlyContinue |
        Where-Object { $_.LastWriteTime -gt $beforeDownload -or $_.CreationTime -gt $beforeDownload } |
        Sort-Object LastWriteTime -Descending |
        Select-Object -First 1

    if ($null -ne $newLog) {
        # Prefix with Correl_ID_S so Get-GfixLogLines can find it, original name preserved
        $targetName = "{0}_{1}" -f $correl, $newLog.Name
        $dst = Join-Path $logDir $targetName
        Move-Item -LiteralPath $newLog.FullName -Destination $dst -Force
        Write-Host ("  moved: {0} -> log\{1}" -f $newLog.Name, $targetName) -ForegroundColor Green
        $row.GFIX_log = '1'
    } else {
        # Fallback: look for {correl}_*.log already in logDir (manually placed)
        $existing = @(Get-ChildItem -Path $logDir -Filter ("{0}_*.log" -f $correl) -ErrorAction SilentlyContinue)
        if ($existing.Count -gt 0) {
            Write-Host ("  already in log\: {0}" -f $existing[0].Name) -ForegroundColor DarkGray
            $row.GFIX_log = '1'
        } else {
            Write-Host ("  [WARN] no new .log in Downloads for [{0}]" -f $correl) -ForegroundColor Yellow
        }
    }

    Send-Tab 2
    Send-Enter
    Start-Sleep -Milliseconds 800
}

$rows | Export-Csv $mappingFile -NoTypeInformation -Encoding UTF8
Write-Host "`n[GfixLogDownload] Mapping saved + step finished." -ForegroundColor Green

# Return focus to this console window
$consoleHwnd = (Get-Process -Id $PID).MainWindowHandle
if ($consoleHwnd -ne [IntPtr]::Zero) {
    [WinAPI]::SetForegroundWindow($consoleHwnd) | Out-Null
    [WinAPI]::ShowWindowAsync($consoleHwnd, 9) | Out-Null   # SW_RESTORE
}