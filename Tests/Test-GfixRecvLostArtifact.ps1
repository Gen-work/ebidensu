#Requires -Version 5.1
# Test-GfixRecvLostArtifact.ps1 -- 10/20: the worklist said track = ok for
# MJODWSPQ but capture\gfix\MJODWSPQ\track.json was not there, so logs
# stopped on a FAILED panel and nothing told the operator what to redo.
# Now logs / evidence set the upstream verdict back to pending and skip the
# row (not_met -> skip), so status / menu / the next track run pick it up.
#
# The steps under test are copied out of the SHIPPED workflows (the first
# steps of logs.each / evidence.each), so this checks what the office PC
# runs, not a hand-written copy. Real modules, real gfix-recv profile, a
# temp work dir with a two-row gfixrecv.csv.

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
$repo = Split-Path $here -Parent
. (Join-Path $repo 'kernel/Runner.ps1')
. (Join-Path $repo 'kernel/Profile.ps1')
. (Join-Path $repo 'kernel/Json.ps1')

Reset-Tests 'GfixRecvLostArtifact'

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-lost-' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null

function Write-LostCsv {
    param([string]$Track, [string]$Logs)
    $lines = @(
        'Excel_NAME,JOB,USER,RCVFILE,GFIX_DATE,GFIX_TIME,FileCount,Owner,track,logs,evidence,deliver,note',
        ('MJODWSPQ,MJODJSPQ,JODUSR,JJPCRS13,2026-10-20,11:02:00,1,,' + $Track + ',' + $Logs + ',,,'),
        ('QJODWSPP,QJODJSPP,JODUSR,JJPCRS13,2026-10-20,11:45:00,1,,' + $Track + ',' + $Logs + ',,,')
    )
    $enc = New-Object System.Text.UTF8Encoding($true)
    [System.IO.File]::WriteAllText((Join-Path $tmp 'gfixrecv.csv'), (($lines -join "`r`n") + "`r`n"), $enc)
}

function Read-LostCsvRow {
    param([string]$Key)
    $rows = @(Import-Csv -LiteralPath (Join-Path $tmp 'gfixrecv.csv') -Encoding UTF8)
    return @($rows | Where-Object { $_.Excel_NAME -eq $Key })[0]
}

function New-LostWorkflow {
    # setup = the shipped table.load; each = the shipped steps before the
    # first one that needs a window / Excel, then a marker that proves the
    # row went on past them.
    param([string]$Flow, [string[]]$Ids, [string]$Field)
    $src = Read-EbiJson -Path (Join-Path $repo ('workflows/gfixRecv.' + $Flow + '.json'))
    $wf = $src['value']
    $each = @(@($wf['each']) | Where-Object { $Ids -contains [string]$_['id'] })
    $each += @{ id = 'reached'; use = 'file.write_json'; with = @{ path = 'reached/{{item.keySafe}}.json'; data = @{ key = '{{item.key}}' } } }
    $mini = @{
        schema = 1; id = ('lost.' + $Flow); title = 'lost artifact'; version = '0.0.0'; profile = 'gfix-recv'
        vars = $wf['vars']; source = $wf['source']; onError = $wf['onError']
        setup = @(@($wf['setup']) | Where-Object { [string]$_['id'] -eq 'load' })
        each = $each
    }
    $p = Join-Path $tmp ('lost.' + $Flow + '.json')
    [void](Write-EbiJson -Path $p -Value $mini)
    return $p
}

function Write-Sidecar {
    param([string]$Key, [string]$Name, [hashtable]$Data)
    $d = Join-Path $tmp ('capture/gfix/' + $Key)
    New-Item -ItemType Directory -Path $d -Force | Out-Null
    [void](Write-EbiJson -Path (Join-Path $d $Name) -Value $Data)
}

try {
    $pr = Read-EbiProfile -Dir (Join-Path $repo 'profiles/gfix-recv') -WorkDir $tmp
    Assert-True $pr['ok'] ('profile loads: ' + $pr['message'])
    $profile = $pr['value']
    $modules = Join-Path $repo 'modules'
    $quiet = { param($q) return 's' }

    # ---- logs: track = ok but track.json missing -> track pending, row skipped
    Write-LostCsv -Track 'ok' -Logs ''
    Write-Sidecar -Key 'QJODWSPP' -Name 'track.json' -Data @{ key = 'QJODWSPP'; verdict = 'ok'; recvJobs = @('1') }
    $wf = New-LostWorkflow -Flow 'logs' -Ids @('track', 'trackLost', 'trackHere', 'trackOk')
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $tmp -ModulesRoot $modules -RunId 'lost-logs' -Profile $profile -AskHandler $quiet -AllDays
    $mj = @($res['items'] | Where-Object { $_['key'] -eq 'MJODWSPQ' })[0]
    $qj = @($res['items'] | Where-Object { $_['key'] -eq 'QJODWSPP' })[0]
    Assert-Equal 'skip' $mj['status'] ('logs: the row whose track.json is gone is skipped, not failed: ' + $mj['failure'] + ' ' + $mj['message'])
    Assert-True ([string]$mj['message'] -like '*track.json is not there*set back to pending*') ('logs: the skip says what to redo: ' + $mj['message'])
    Assert-Equal '' ([string](Read-LostCsvRow 'MJODWSPQ').track) 'logs: track is set back to pending (blank) in gfixrecv.csv'
    Assert-True (-not (Test-Path (Join-Path $tmp 'reached/MJODWSPQ.json'))) 'logs: nothing after the check ran for that row'
    Assert-Equal 'ok' $qj['status'] ('logs: a row with its track.json goes on: ' + $qj['message'])
    Assert-Equal 'ok' ([string](Read-LostCsvRow 'QJODWSPP').track) 'logs: that row keeps track = ok'
    Assert-True (Test-Path (Join-Path $tmp 'reached/QJODWSPP.json')) 'logs: that row reached the next step'

    # ---- evidence: logs.json missing -> logs pending; track.json missing -> track pending
    Remove-Item -LiteralPath (Join-Path $tmp 'reached') -Recurse -Force -ErrorAction SilentlyContinue
    Write-LostCsv -Track 'ok' -Logs 'ok'
    Write-Sidecar -Key 'MJODWSPQ' -Name 'track.json' -Data @{ key = 'MJODWSPQ'; verdict = 'ok' }
    $wf = New-LostWorkflow -Flow 'evidence' -Ids @('track', 'trackLost', 'trackHere', 'trackOk', 'logs', 'logsLost', 'logsHere')
    $res = Invoke-EbiWorkflow -Path $wf -WorkDir $tmp -ModulesRoot $modules -RunId 'lost-evidence' -Profile $profile -AskHandler $quiet -AllDays
    $mj = @($res['items'] | Where-Object { $_['key'] -eq 'MJODWSPQ' })[0]
    Assert-Equal 'skip' $mj['status'] ('evidence: no logs.json -> skipped: ' + $mj['message'])
    Assert-Equal 'ok' ([string](Read-LostCsvRow 'MJODWSPQ').track) 'evidence: track.json is there -> track stays ok'
    Assert-Equal '' ([string](Read-LostCsvRow 'MJODWSPQ').logs) 'evidence: logs is set back to pending'
    Assert-True (-not (Test-Path (Join-Path $tmp 'reached/MJODWSPQ.json'))) 'evidence: the book is never reached for that row'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
