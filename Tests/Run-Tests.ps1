#Requires -Version 5.1
# ============================================================
#  Run-Tests.ps1
#
#  1) Parse-checks every .ps1 in the repo (syntax errors fail the run).
#  2) Runs each Tests\Test-*.ps1 and aggregates pass/fail.
#
#  Usage:  .\Tests\Run-Tests.ps1
# ============================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent

Write-Host ''
Write-Host '===== Parse check (all *.ps1) =====' -ForegroundColor Green
$parseErrors = 0
$psFiles = @(Get-ChildItem -LiteralPath $repoRoot -Filter '*.ps1' -File -Recurse)
foreach ($f in $psFiles) {
    $tokens = $null; $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs)
    if ($errs -and $errs.Count -gt 0) {
        $parseErrors += $errs.Count
        Write-Host ('  [PARSE-FAIL] {0}' -f $f.Name) -ForegroundColor Red
        foreach ($e in $errs) {
            Write-Host ('      line {0}: {1}' -f $e.Extent.StartLineNumber, $e.Message) -ForegroundColor Red
        }
    }
}
if ($parseErrors -eq 0) {
    Write-Host ('  OK: {0} files parsed clean' -f $psFiles.Count) -ForegroundColor Green
}

Write-Host ''
Write-Host '===== Unit tests =====' -ForegroundColor Green
$totalFail = 0
$testFiles = @(Get-ChildItem -LiteralPath $here -Filter 'Test-*.ps1' -File | Sort-Object Name)
foreach ($t in $testFiles) {
    & $t.FullName
    $rc = $LASTEXITCODE
    if ($null -eq $rc) { $rc = 0 }
    $totalFail += [int]$rc
}

Write-Host ''
Write-Host '===== Run-Tests summary =====' -ForegroundColor Green
Write-Host ('  parse errors : {0}' -f $parseErrors) -ForegroundColor $(if ($parseErrors -gt 0) { 'Red' } else { 'Green' })
Write-Host ('  test failures: {0}' -f $totalFail)   -ForegroundColor $(if ($totalFail   -gt 0) { 'Red' } else { 'Green' })

$rcAll = $parseErrors + $totalFail
if ($rcAll -gt 0) {
    Write-Host '===== RESULT: FAIL =====' -ForegroundColor Red
} else {
    Write-Host '===== RESULT: PASS =====' -ForegroundColor Green
}
exit $rcAll
