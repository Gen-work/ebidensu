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

function Assert-True {
    param([bool]$Cond, [string]$Msg)
    if ($Cond) {
        $script:TestPass++; Write-Host ('  [PASS] {0}' -f $Msg) -ForegroundColor DarkGreen
    } else {
        $script:TestFail++; Write-Host ('  [FAIL] {0}' -f $Msg) -ForegroundColor Red
    }
}

function Assert-Equal {
    param($Expected, $Actual, [string]$Msg)
    if ([string]$Expected -eq [string]$Actual) {
        $script:TestPass++; Write-Host ('  [PASS] {0}' -f $Msg) -ForegroundColor DarkGreen
    } else {
        $script:TestFail++
        Write-Host ("  [FAIL] {0} (expected '{1}', got '{2}')" -f $Msg, $Expected, $Actual) -ForegroundColor Red
    }
}

function Complete-Tests {
    $color = if ($script:TestFail -gt 0) { 'Red' } else { 'Green' }
    Write-Host ('  ---- {0}: {1} passed, {2} failed ----' -f $script:TestName, $script:TestPass, $script:TestFail) -ForegroundColor $color
    return $script:TestFail
}
