#Requires -Version 5.1
# Test-Ledger.ps1 -- kernel/Ledger.ps1 (P1-04): ledger keys, read/append,
# run.json, unfinished-run discovery. The resume behaviour itself is
# exercised end to end in Test-Runner.ps1.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Ledger.ps1')

Reset-Tests 'Ledger'

Assert-Equal 'item:A1 / JOB_A|shot' (Get-EbiLedgerKey -Item 'A1 / JOB_A' -Step 'shot') 'key: (item, step)'
Assert-Equal 'group:JOB_A|open' (Get-EbiLedgerKey -Item 'A1 / JOB_A' -Group 'JOB_A' -Step 'open') 'key: (group, step) when a group is given'
Assert-True ((Get-EbiLedgerFile -WorkDir 'W' -RunId 'r1').EndsWith(('run' + [IO.Path]::DirectorySeparatorChar + 'r1' + [IO.Path]::DirectorySeparatorChar + 'ledger.jsonl'))) 'file: run/<runId>/ledger.jsonl'
Assert-True ((Get-EbiRunFile -WorkDir 'W' -RunId 'r1').EndsWith('run.json')) 'file: run/<runId>/run.json'

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-ledger-' + [guid]::NewGuid().ToString('N'))
try {
    $file = Get-EbiLedgerFile -WorkDir $tmp -RunId 'r1'
    $r = Read-EbiLedger -Path $file
    Assert-True ($r['ok'] -and $r['count'] -eq 0 -and $r['done'].Count -eq 0) 'read: no ledger yet is ok and empty'

    $rec = New-EbiLedgerRecord -RunId 'r1' -Item 'A1' -Step 'shot' -Status 'ok' -Outputs @{ path = 'a.png'; width = 10 }
    Assert-Equal 'A1' $rec['item'] 'record: item'
    Assert-True (-not $rec.Contains('group')) 'record: no group key on an item record'
    Assert-True ([datetime]::TryParse([string]$rec['ts'], [ref]([datetime]::MinValue))) 'record: ts is a timestamp'
    Assert-True ((Add-EbiLedgerRecord -Path $file -Record $rec)['ok']) 'append: first record'
    [void](Add-EbiLedgerRecord -Path $file -Record (New-EbiLedgerRecord -RunId 'r1' -Group 'G1' -Step 'open' -Status 'ok' -Outputs @{ url = 'u' }))
    [void](Add-EbiLedgerRecord -Path $file -Record (New-EbiLedgerRecord -RunId 'r1' -Item 'A1' -Step 'opt' -Status 'skipped' -Outputs @{ a = $null; skipped = $true }))
    [void](Add-EbiLedgerRecord -Path $file -Record (New-EbiLedgerRecord -RunId 'r1' -Item 'A1' -Step 'shot' -Status 'ok' -Outputs @{ path = 'b.png'; width = 20 }))
    $r = Read-EbiLedger -Path $file
    Assert-Equal 4 $r['count'] 'read: four records'
    Assert-Equal 3 $r['done'].Count 'read: three keys (shot twice)'
    Assert-Equal 'b.png' $r['done']['item:A1|shot']['outputs']['path'] 'read: the LAST record for a key wins'
    Assert-Equal 'u' $r['done']['group:G1|open']['outputs']['url'] 'read: a group record is keyed by group'
    Assert-True ($r['done']['item:A1|opt'].Contains('outputs') -and $r['done']['item:A1|opt']['outputs']['skipped'] -eq $true) 'read: a skipped record keeps its null-field outputs'
    Assert-True (-not $rec.Contains('outputs') -or ($rec['outputs'] -is [hashtable])) 'record: outputs is a hashtable'
    [System.IO.File]::AppendAllText($file, 'garbage' + [Environment]::NewLine + '{"runId":"r1","item":"A2","step":"x","status":"ok","outputs":{}}' + [Environment]::NewLine, (New-Object System.Text.UTF8Encoding($false)))
    $r = Read-EbiLedger -Path $file
    Assert-True ($r['ok'] -and $r['done'].Contains('item:A2|x')) 'read: a bad middle line does not lose the good lines'
    Assert-Equal 1 @($r['badLines']).Count 'read: ... and is reported'

    $run = @{ runId = 'r1'; startedAt = '2026-09-28T09:00:00'; operator = 'misaki'; workDir = $tmp; timeWindow = @{ from = 'a'; to = 'b' } }
    $wf = @{ id = 'wf.x'; version = '1.2.3'; profile = 'host-open'; _path = 'C:\w\x.json' }
    $w = Write-EbiRunFile -WorkDir $tmp -RunId 'r1' -Run $run -Workflow $wf -RunArgs @{ only = @('A1'); limit = 0 }
    Assert-True ($w['ok']) 'run.json: written'
    $d = (Read-EbiRunFile -WorkDir $tmp -RunId 'r1')['value']
    Assert-Equal 'misaki' $d['operator'] 'run.json: operator'
    Assert-Equal 'b' $d['timeWindow']['to'] 'run.json: timeWindow'
    Assert-Equal 'wf.x' $d['workflow']['id'] 'run.json: workflow id'
    Assert-Equal '1.2.3' $d['workflow']['version'] 'run.json: workflow version'
    Assert-Equal 'host-open' $d['profile'] 'run.json: profile name'
    Assert-Equal 'A1' $d['args']['only'][0] 'run.json: args'
    Assert-Equal 'False' ([string]$d['finished']) 'run.json: unfinished at start'
    Assert-Equal 1 @(Find-EbiUnfinishedRuns -WorkDir $tmp).Count 'unfinished: found'
    Assert-Equal 1 @(Find-EbiUnfinishedRuns -WorkDir $tmp -WorkflowId 'wf.x').Count 'unfinished: found by workflow id'
    Assert-Equal 0 @(Find-EbiUnfinishedRuns -WorkDir $tmp -WorkflowId 'wf.y').Count 'unfinished: another workflow id finds nothing'
    [void](Write-EbiRunFile -WorkDir $tmp -RunId 'r1' -Run $run -Workflow $wf -RunArgs @{} -Finished $true -Result @{ ok = $true })
    Assert-Equal 0 @(Find-EbiUnfinishedRuns -WorkDir $tmp).Count 'unfinished: a finished run is not listed'
    Assert-Equal 'True' ([string](Read-EbiRunFile -WorkDir $tmp -RunId 'r1')['value']['result']['ok']) 'run.json: the result is stored at the end'
    Assert-Equal 0 @(Find-EbiUnfinishedRuns -WorkDir (Join-Path $tmp 'nowhere')).Count 'unfinished: no run dir -> nothing, no throw'
} finally {
    if (Test-Path -LiteralPath $tmp) { Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue }
}

$rc = Complete-Tests
exit $rc
