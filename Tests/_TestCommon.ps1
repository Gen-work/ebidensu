# Tiny assert helpers shared by the Test-*.ps1 files. No param() block.
$script:TestPass = 0
$script:TestFail = 0
$script:TestName = 'tests'

function Reset-Tests {
    param([string]$Name = 'tests')
    $script:TestPass = 0
    $script:TestFail = 0
    $script:TestName = $Name
    Write-Host ''
    Write-Host ('===== {0} =====' -f $Name) -ForegroundColor Green
}

function Register-TestFailure {
    # Run-Tests.ps1 creates $Global:EbiTestFailures before the suites run and
    # prints it as one block at the very end, so every [FAIL] of a whole run
    # can be copied from one place. A test file run on its own has no such
    # list and just prints its [FAIL] lines as before.
    param([string]$Line)
    $list = Get-Variable -Name EbiTestFailures -Scope Global -ValueOnly -ErrorAction SilentlyContinue
    if ($null -ne $list) { [void]$list.Add(('{0}: {1}' -f $script:TestName, $Line)) }
}

function Assert-True {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) {
        $script:TestPass++; Write-Host ('  [PASS] {0}' -f $Msg) -ForegroundColor DarkGreen
    } else {
        $script:TestFail++; Write-Host ('  [FAIL] {0}' -f $Msg) -ForegroundColor Red
        Register-TestFailure $Msg
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Msg)
    if ([string]$Expected -eq [string]$Actual) {
        $script:TestPass++; Write-Host ('  [PASS] {0}' -f $Msg) -ForegroundColor DarkGreen
    } else {
        $script:TestFail++
        $line = ("{0} (expected '{1}', got '{2}')" -f $Msg, $Expected, $Actual)
        Write-Host ('  [FAIL] ' + $line) -ForegroundColor Red
        Register-TestFailure $line
    }
}

function Complete-Tests {
    $color = if ($script:TestFail -gt 0) { 'Red' } else { 'Green' }
    Write-Host ('  ---- {0}: {1} passed, {2} failed ----' -f $script:TestName, $script:TestPass, $script:TestFail) -ForegroundColor $color
    return $script:TestFail
}
