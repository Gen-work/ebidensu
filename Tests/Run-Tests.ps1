#Requires -Version 5.1
# ============================================================
#  Run-Tests.ps1
#
#  1) Parse-checks every .ps1 in the repo (syntax errors fail the run).
#  2) Runs each Tests\Test-*.ps1 and aggregates pass/fail.
#
#  Usage:  .\Tests\Run-Tests.ps1
#
#  Never stops to ask: EBI_NO_ASK=1 is set for the run, so every gate /
#  confirm / error question a suite reaches without an injected reader
#  takes its default answer (skip an error, yes to a confirm) instead of
#  waiting at Read-Host on a real console. A progress bar shows which
#  phase and suite is running.
# ============================================================
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
$prevNoAsk = $env:EBI_NO_ASK
$env:EBI_NO_ASK = '1'
$progressId = 4711

Write-Host ''
Write-Host '===== Parse check (all *.ps1) =====' -ForegroundColor Green
$parseErrors = 0
# Every failure of this run, collected here and printed again as ONE block at
# the end. The live output above it is long and mixes in lines that only look
# like failures ([note] exempt libraries, a fixture step's own [fail]/[refused]
# console lines inside the Runner suite); this block is the part to copy.
$Global:EbiTestFailures = New-Object System.Collections.ArrayList
$psFiles = @(Get-ChildItem -LiteralPath $repoRoot -Filter '*.ps1' -File -Recurse)
$parsed = 0
foreach ($f in $psFiles) {
    $parsed++
    if ($parsed -eq 1 -or ($parsed % 25) -eq 0 -or $parsed -eq $psFiles.Count) {
        Write-Progress -Id $progressId -Activity 'Run-Tests' -Status ('parse check {0}/{1}' -f $parsed, $psFiles.Count) -PercentComplete ([int](10 * $parsed / [Math]::Max(1, $psFiles.Count)))
    }
    $tokens = $null; $errs = $null
    [void][System.Management.Automation.Language.Parser]::ParseFile($f.FullName, [ref]$tokens, [ref]$errs)
    if ($errs -and $errs.Count -gt 0) {
        $parseErrors += $errs.Count
        Write-Host ('  [PARSE-FAIL] {0}' -f $f.Name) -ForegroundColor Red
        foreach ($e in $errs) {
            Write-Host ('      line {0}: {1}' -f $e.Extent.StartLineNumber, $e.Message) -ForegroundColor Red
            [void]$Global:EbiTestFailures.Add(('parse: {0} line {1}: {2}' -f $f.Name, $e.Extent.StartLineNumber, $e.Message))
        }
    }
}
if ($parseErrors -eq 0) {
    Write-Host ('  OK: {0} files parsed clean' -f $psFiles.Count) -ForegroundColor Green
}

# Where those files live. The refactor moves scripts out of the repo root into
# modules/ (steps), legacy/ (retired) and Tests/, so print the split: a file
# that lands in the wrong area is otherwise invisible in a single total.
$areas = [ordered]@{}
foreach ($f in $psFiles) {
    $rel  = $f.FullName.Substring($repoRoot.Length).TrimStart('\', '/')
    $sep  = $rel.IndexOfAny(@([char]'\', [char]'/'))
    $area = if ($sep -gt 0) { $rel.Substring(0, $sep) } else { '(root)' }
    if (-not $areas.Contains($area)) { $areas[$area] = 0 }
    $areas[$area] = $areas[$area] + 1
}
foreach ($area in $areas.Keys) {
    Write-Host ('    {0,-12} {1,4}' -f $area, $areas[$area]) -ForegroundColor DarkGray
}

Write-Host ''
Write-Host '===== Mask check (P2-08) =====' -ForegroundColor Green
# Fixtures are real page text: a sensitive item that lands in git is a
# history rewrite later, so the gate runs before the unit tests.
Write-Progress -Id $progressId -Activity 'Run-Tests' -Status 'mask check' -PercentComplete 10
. (Join-Path (Join-Path $repoRoot 'kernel') 'Mask.ps1')
$maskRes = Invoke-EbiMaskCheck -RepoRoot $repoRoot
foreach ($l in @(Format-EbiMaskReport -Result $maskRes -RepoRoot $repoRoot)) { Write-Host $l -ForegroundColor $(if ($maskRes['ok']) { 'Green' } else { 'Red' }) }
$maskFail = if ($maskRes['ok']) { 0 } else { 1 }
if ($maskFail) { [void]$Global:EbiTestFailures.Add('mask check: ' + $maskRes['message']) }

Write-Host ''
Write-Host '===== Unit tests =====' -ForegroundColor Green
$totalFail = 0
# Recursive: Tests/ gains subdirectories as the module tree grows, and a test
# that silently stops being discovered is worse than one that fails.
$testFiles = @(Get-ChildItem -LiteralPath $here -Filter 'Test-*.ps1' -File -Recurse | Sort-Object FullName)
$suiteNo = 0
foreach ($t in $testFiles) {
    $suiteNo++
    # Percent 10..100 over the suites (parse + mask took the first 10).
    Write-Progress -Id $progressId -Activity 'Run-Tests' -Status ('suite {0}/{1}: {2}   (failures so far: {3})' -f $suiteNo, $testFiles.Count, $t.Name, $Global:EbiTestFailures.Count) -PercentComplete ([int](10 + 90 * ($suiteNo - 1) / [Math]::Max(1, $testFiles.Count)))
    # One suite that dies (a terminating error outside any assertion) is ONE
    # failure, reported by file and line; the suites after it still run. On
    # PS 5.1 a provider error (Copy-Item into a missing directory) terminated
    # Test-P2 and, before this guard, took the rest of the run with it.
    # Two ways a suite can die: the error propagates out of the call (5.1's
    # provider errors do), or the script just ends without reaching its exit
    # (then $LASTEXITCODE is still the null set here, never the previous
    # suite's stale code). Both count as one failure with the file and line.
    $global:LASTEXITCODE = $null
    $errBefore = $Error.Count
    $rc = $null; $crash = $null
    try {
        & $t.FullName
        $rc = $LASTEXITCODE
    } catch { $crash = $_ }
    if ($null -eq $rc) {
        if ($null -eq $crash -and $Error.Count -gt $errBefore) { $crash = $Error[0] }
        $where = ''; $msg = 'the suite ended without reaching its exit'
        if ($null -ne $crash) {
            $msg = [string]$crash.Exception.Message
            try { if ($null -ne $crash.InvocationInfo -and $crash.InvocationInfo.ScriptLineNumber -gt 0 -and -not [string]::IsNullOrEmpty([string]$crash.InvocationInfo.ScriptName)) { $where = ' at ' + (Split-Path -Leaf ([string]$crash.InvocationInfo.ScriptName)) + ':' + $crash.InvocationInfo.ScriptLineNumber } } catch { }
        }
        $line = ('{0}: stopped with an error{1}: {2}' -f $t.Name, $where, $msg)
        Write-Host ('  [FAIL] ' + $line) -ForegroundColor Red
        Write-Host ('  ---- {0}: stopped, counted as 1 failure; the remaining suites still run ----' -f ($t.BaseName -replace '^Test-', '')) -ForegroundColor Red
        if ($null -ne $Global:EbiTestFailures) { [void]$Global:EbiTestFailures.Add($line) }
        $rc = 1
    }
    $totalFail += [int]$rc
}

Write-Progress -Id $progressId -Activity 'Run-Tests' -Completed
$env:EBI_NO_ASK = $prevNoAsk

Write-Host ''
Write-Host '===== Not passed (copy from here) =====' -ForegroundColor Green
if ($Global:EbiTestFailures.Count -eq 0) {
    Write-Host '  (none)' -ForegroundColor Green
} else {
    foreach ($line in $Global:EbiTestFailures) {
        Write-Host ('  [FAIL] {0}' -f $line) -ForegroundColor Red
    }
}
Write-Host ('  {0} not passed. Only these count: [note] lines are libraries still exempt from the step contract, and [info]/[fail]/[skip]/[refused] lines inside the Runner suite are fixture steps failing on purpose.' -f $Global:EbiTestFailures.Count) -ForegroundColor DarkGray

Write-Host ''
Write-Host '===== Run-Tests summary =====' -ForegroundColor Green
Write-Host ('  parse errors : {0}' -f $parseErrors) -ForegroundColor $(if ($parseErrors -gt 0) { 'Red' } else { 'Green' })
Write-Host ('  test failures: {0}' -f $totalFail)   -ForegroundColor $(if ($totalFail   -gt 0) { 'Red' } else { 'Green' })
Write-Host ('  mask hits    : {0}' -f @($maskRes['hits']).Count) -ForegroundColor $(if ($maskFail -gt 0) { 'Red' } else { 'Green' })

$rcAll = $parseErrors + $totalFail + $maskFail
if ($rcAll -gt 0) {
    Write-Host '===== RESULT: FAIL =====' -ForegroundColor Red
} else {
    Write-Host '===== RESULT: PASS =====' -ForegroundColor Green
}
exit $rcAll
