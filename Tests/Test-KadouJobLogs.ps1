#Requires -Version 5.1
# Test-KadouJobLogs.ps1 -- verify.job_logs (kadou.collect, the 10/26
# JJODWDBx operation check) on the GoAnywhere job logs of the reference
# evidence (Tests/fixtures/jod-kadou: 07/29 DBM and 08/03 DBF, host and user
# names masked): which logs are one transfer's, the four checks, the run
# chosen for the evidence.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Profile.ps1')

Reset-Tests 'KadouJobLogs'

$fx = Join-Path (Join-Path $here 'fixtures') 'jod-kadou'
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-kadou-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
$log = New-Object PSObject
$log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
$log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
$ctx = @{ WorkDir = $tmp; RunId = 'test'; Profile = @{}; Log = $log; DryRun = $false; Session = @{}; Item = $null; KeyColumns = @('CORREL') }

$r = . Import-EbiStep -Registry $reg -Use 'verify.job_logs'
if (-not $r['ok']) { throw ('cannot load verify.job_logs: ' + $r['message']) }
$entry = $r['Entry']
function Invoke-JobLogs {
    param([hashtable]$With)
    $res = Resolve-EbiStepInputs -Manifest $entry['Manifest'] -With $With -Session $ctx['Session']
    if (-not $res['ok']) { throw ('verify.job_logs inputs: ' + $res['message']) }
    return (& $entry['Invoke'] $res['In'] $ctx)
}

try {
    $all = @(Get-ChildItem -LiteralPath $fx -Filter '*.log' | Sort-Object Name | ForEach-Object { $_.FullName })
    Assert-Equal 6 $all.Count 'fixtures: six reference job logs (three transfers on 07/29, three on 08/03)'

    $o = Invoke-JobLogs @{ paths = $all; key = 'JIDSU86S'; recvDir = '/appl/IDS/IDSVer1/gfix/recv'; shell = 'IDSLB283run.sh'; ss = 'U'; day = '2026-08-03' }
    Assert-True ($o['ok'] -and $o['code'] -eq 'ok' -and $o['runCount'] -eq 1) ('job_logs: the IDS transfer of 08/03 is one run, ok: ' + $o['reason'])
    Assert-Equal 'JIDSU86S.260803.10422857' $o['chosenFile'] 'job_logs: the chosen run is that file'
    Assert-Equal '1000003451754' $o['runs'][0]['jobNumber'] 'job_logs: the job number read off the log'
    Assert-Equal '31878' $o['runs'][0]['bytes'] 'job_logs: the size (commas dropped)'
    Assert-True (([string]$o['chosenPath']).EndsWith('1000003451754.log')) 'job_logs: chosenPath is that log file'

    $o = Invoke-JobLogs @{ paths = $all; key = 'JIGPS86S'; recvDir = '/appl/IGP2/IGP2Ver1/gfix/recv/'; shell = 'IGPLC083run.sh'; ss = 'S' }
    Assert-Equal 'ok' $o['code'] ('job_logs: IGP2 IF5133 with shell IGPLC083 / S (a trailing / on recvDir is fine): ' + $o['reason'])

    $o = Invoke-JobLogs @{ paths = $all; key = 'JIGPS86S'; shell = 'IGPLB283run.sh'; ss = 'U' }
    Assert-True ($o['code'] -eq 'ng' -and $o['reason'] -like '*not IGPLB283run.sh*' -and $o['reason'] -like '*not U*') ('job_logs: the wrong shell and SS are both said: ' + $o['reason'])
    Assert-True (-not $o['hasChosen']) 'job_logs: an ng run is never chosen for the evidence'

    $o = Invoke-JobLogs @{ paths = $all; key = 'JIDSU86S'; day = '2026-10-26' }
    Assert-True ($o['code'] -eq 'unknown' -and $o['runCount'] -eq 0 -and $o['reason'] -like '*2026-10-26*') 'job_logs: no run that day -> unknown, said'
    $o = Invoke-JobLogs @{ paths = $all; key = 'JIDSU8BS' }
    Assert-True ($o['code'] -eq 'ok' -and $o['chosenStart'] -eq '2026-07-29 10:53:07') 'job_logs: the 07/29 DBM transfer (no day filter)'

    # a failed run: return code 1, no delete line
    $bad = (Get-Content -LiteralPath (Join-Path $fx '1000003451754.log') -Raw -Encoding UTF8)
    $bad = $bad.Replace("return code '0'", "return code '1'") -replace "(?m)^.*was deleted\.\r?\n", ''
    $bad = $bad.Replace('JIDSU86S.260803.10422857', 'JIDSU86S.260803.12000011').Replace('2026-08-03 10:42:3', '2026-08-03 12:00:0')
    $badPath = Join-Path $tmp 'bad.log'
    [System.IO.File]::WriteAllText($badPath, $bad, (New-Object System.Text.UTF8Encoding($false)))
    $o = Invoke-JobLogs @{ paths = @($all + $badPath); key = 'JIDSU86S'; recvDir = '/appl/IDS/IDSVer1/gfix/recv'; shell = 'IDSLB283run.sh'; ss = 'U'; day = '2026-08-03' }
    Assert-True ($o['code'] -eq 'ng' -and $o['runCount'] -eq 2 -and $o['okCount'] -eq 1) ('job_logs: one good and one failed run -> ng for the transfer: ' + $o['reason'])
    $badRun = @($o['runs'] | Where-Object { $_['code'] -eq 'ng' })[0]
    Assert-True ((@($badRun['reasons']) -join '|') -like '*return code 1*' -and (@($badRun['reasons']) -join '|') -like '*was deleted*') ('job_logs: both failed checks are named: ' + (@($badRun['reasons']) -join '; '))
    Assert-Equal 'JIDSU86S.260803.10422857' $o['chosenFile'] 'job_logs: the good run is still the one chosen for the evidence'
    Assert-Equal 2 @($o['summary']).Count 'job_logs: one summary line per run'

    # two good runs: earliest, unless the receive folder listing only has the later one
    $late = (Get-Content -LiteralPath (Join-Path $fx '1000003451754.log') -Raw -Encoding UTF8).Replace('JIDSU86S.260803.10422857', 'JIDSU86S.260803.12000011').Replace('2026-08-03 10:42:3', '2026-08-03 12:00:0')
    $latePath = Join-Path $tmp 'late.log'
    [System.IO.File]::WriteAllText($latePath, $late, (New-Object System.Text.UTF8Encoding($false)))
    $o = Invoke-JobLogs @{ paths = @($latePath) + $all; key = 'JIDSU86S'; day = '2026-08-03' }
    Assert-True ($o['code'] -eq 'ok' -and $o['runCount'] -eq 2 -and $o['chosenFile'] -eq 'JIDSU86S.260803.10422857') 'job_logs: two ok runs -> the earliest is chosen (runs sorted by start, not by path order)'
    $o = Invoke-JobLogs @{ paths = @($latePath) + $all; key = 'JIDSU86S'; day = '2026-08-03'; preferText = "JIDSU86S.260803.12000011`t2026/08/03 12:00:05`t31.13 KB" }
    Assert-True ($o['chosenFile'] -eq 'JIDSU86S.260803.12000011' -and $o['listed']) 'job_logs: preferText -> the earliest ok run whose file the receive folder lists'
    $o = Invoke-JobLogs @{ paths = $all; key = 'JIDSU86S'; day = '2026-08-03'; preferText = 'nothing here' }
    Assert-True ($o['hasChosen'] -and -not $o['listed'] -and $o['reason'] -like '*NOT in the receive folder*') 'job_logs: no listed file -> the earliest ok run, flagged'

    Assert-True (-not (VerifyJobLogs-IsKey -FileName 'JIDSU86S.260803.1' -Key 'JIDSU86')) 'job_logs: a key is a whole name before the dot (JIDSU86 is not JIDSU86S)'
    Assert-True (VerifyJobLogs-IsKey -FileName 'jidsu86s' -Key 'JIDSU86S') 'job_logs: an unstamped file name equal to the key counts, case folded'

    $o = Invoke-JobLogs @{ paths = @((Join-Path $tmp 'missing.log')); key = 'JIDSU86S' }
    Assert-Equal 'file_not_found' $o['failure'] 'job_logs: a listed path that is missing -> file_not_found'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
