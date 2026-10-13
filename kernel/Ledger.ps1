#Requires -Version 5.1
# ============================================================
#  kernel/Ledger.ps1
#
#  The run ledger and run.json (P1-04). Dot-source only (no param()
#  block, ASCII source, no class). All file I/O through kernel/Json.ps1.
#
#  run/<runId>/ledger.jsonl  -- one record per completed (item, step) or
#  (group, step), WITH the step's outputs (STEP-CONTRACT.md 6.1), so a
#  resume can replay them into the template scope instead of re-running
#  the step. Only "each" is ledgered: setup and teardown re-run on every
#  resume (6.2). A when-skipped step is a record too (status 'skipped',
#  outputs all null + skipped=true). A FAILED step is never a record --
#  the next run must try it again.
#
#  run/<runId>/run.json      -- the run.* scope (runId / startedAt /
#  date / dateSlash / mmdd / toolDir / operator / workDir / timeWindow) plus workflow id/version, the CLI
#  arguments and the final result, written at start and at end
#  (P0-R16). --resume restores run.* from it, so nothing is asked twice.
#
#  Record shape:
#    { "runId", "item" | "group", "step", "status", "outputs", "ts" }
#  Ledger key: 'item:<key display>|<step id>' or 'group:<group>|<step id>'.
#  Both are decided here so the writer and the reader cannot drift.
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')

function Get-EbiRunDir {
    param([string]$WorkDir, [string]$RunId)
    return (Join-Path (Join-Path $WorkDir 'run') $RunId)
}

function Get-EbiLedgerFile {
    param([string]$WorkDir, [string]$RunId)
    return (Join-Path (Get-EbiRunDir -WorkDir $WorkDir -RunId $RunId) 'ledger.jsonl')
}

function Get-EbiRunFile {
    param([string]$WorkDir, [string]$RunId)
    return (Join-Path (Get-EbiRunDir -WorkDir $WorkDir -RunId $RunId) 'run.json')
}

function Get-EbiLedgerKey {
    # (item, step) or (group, step) -> the one key both sides use. A
    # once:group / once:groupEnd step passes Group and no Item.
    param([string]$Item = '', [string]$Group = '', [string]$Step)
    if ($Group -ne '') { return ('group:' + $Group + '|' + $Step) }
    return ('item:' + $Item + '|' + $Step)
}

function New-EbiLedgerRecord {
    param([string]$RunId, [string]$Item = '', [string]$Group = '', [string]$Step, [string]$Status, $Outputs)
    $rec = [ordered]@{ runId = $RunId }
    if ($Group -ne '') { $rec['group'] = $Group } else { $rec['item'] = $Item }
    $rec['step']    = $Step
    $rec['status']  = $Status
    $rec['outputs'] = $(if ($null -eq $Outputs) { @{} } else { $Outputs })
    $rec['ts']      = (Get-Date).ToString('o')
    return $rec
}

function Add-EbiLedgerRecord {
    # Append one record. Returns @{ ok; message } (Add-EbiJsonLine's).
    param([string]$Path, $Record)
    return (Add-EbiJsonLine -Path $Path -Value $Record)
}

function Read-EbiLedger {
    <#
      ledger.jsonl -> @{ ok; done; count; badLines; message }. done is a
      hashtable ledger key -> record (the LAST record for a key wins, so a
      step re-run after an ask/retry on a later attempt replaces the
      earlier line). A missing file is ok with an empty map (a first run
      has no ledger). A malformed middle line is reported in badLines.
    #>
    param([string]$Path)
    $done = @{}
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        return @{ ok = $true; done = $done; count = 0; badLines = @(); message = '' }
    }
    $read = Read-EbiJsonLines -Path $Path
    if (-not $read['ok']) { return @{ ok = $false; done = $done; count = 0; badLines = @(); message = $read['message'] } }
    $n = 0
    foreach ($rec in @($read['value'])) {
        if (-not ($rec -is [System.Collections.IDictionary]) -or -not $rec.Contains('step')) { continue }
        $item  = if ($rec.Contains('item')  -and $null -ne $rec['item'])  { [string]$rec['item'] }  else { '' }
        $group = if ($rec.Contains('group') -and $null -ne $rec['group']) { [string]$rec['group'] } else { '' }
        $done[(Get-EbiLedgerKey -Item $item -Group $group -Step ([string]$rec['step']))] = $rec
        $n++
    }
    return @{ ok = $true; done = $done; count = $n; badLines = @($read['badLines']); message = $read['message'] }
}

function Get-EbiLastCommitIndex {
    <#
      PURE. Resume safety for in-memory resources (2026-10-12). A step that
      edits an open workbook only becomes durable when a later step SAVES
      it (outputs.saved = true: excel.tidy, excel.close save). Returns the
      index in Calls (the workflow's "each" list) of the last step of this
      item whose ledger record says saved = true, or -1. On resume, a
      workbook-editing step AFTER that index must run again: the edit it
      made was never saved, the re-opened book does not have it, and
      replaying its outputs would leave a hole (an evidence book without
      its Excel snap / log lines).
    #>
    param($Calls, $Done, [string]$Item, [string]$Group = '')
    $last = -1
    $i = 0
    foreach ($c in @($Calls)) {
        if ($c -is [System.Collections.IDictionary]) {
            $once = if ($c.Contains('once')) { [string]$c['once'] } else { '' }
            $k = if ($once -eq 'group' -or $once -eq 'groupEnd') { Get-EbiLedgerKey -Group $Group -Step ([string]$c['id']) } else { Get-EbiLedgerKey -Item $Item -Step ([string]$c['id']) }
            if ($null -ne $Done -and $Done.Contains($k)) {
                $rec = $Done[$k]
                $o = if ($rec.Contains('outputs')) { $rec['outputs'] } else { $null }
                if ($o -is [System.Collections.IDictionary] -and $o.Contains('saved') -and $o['saved'] -eq $true) { $last = $i }
            }
        }
        $i++
    }
    return $last
}

function Get-EbiCallIndex {
    # PURE. Position of the call with this id in Calls, or -1.
    param($Calls, [string]$Id)
    $i = 0
    foreach ($c in @($Calls)) { if ($c -is [System.Collections.IDictionary] -and [string]$c['id'] -eq $Id) { return $i }; $i++ }
    return -1
}

function Write-EbiRunFile {
    # run.json: the run.* scope + what was run and how. Atomic (Write-EbiJson).
    param([string]$WorkDir, [string]$RunId, [hashtable]$Run, [hashtable]$Workflow, [hashtable]$RunArgs, [bool]$Finished = $false, $Result = $null)
    $doc = @{
        runId      = $RunId
        startedAt  = $Run['startedAt']
        date       = $Run['date']
        dateSlash  = $Run['dateSlash']
        mmdd       = $Run['mmdd']
        toolDir    = $Run['toolDir']
        operator   = $Run['operator']
        workDir    = $Run['workDir']
        timeWindow = $Run['timeWindow']
        workflow   = @{ id = [string]$Workflow['id']; version = $(if ($Workflow.Contains('version')) { [string]$Workflow['version'] } else { '' }); path = $(if ($Workflow.Contains('_path')) { [string]$Workflow['_path'] } else { '' }) }
        profile    = $(if ($Workflow.Contains('profile')) { [string]$Workflow['profile'] } else { '' })
        args       = $RunArgs
        finished   = $Finished
        updatedAt  = (Get-Date).ToString('o')
    }
    if ($null -ne $Result) { $doc['result'] = $Result }
    return (Write-EbiJson -Path (Get-EbiRunFile -WorkDir $WorkDir -RunId $RunId) -Value $doc)
}

function Read-EbiRunFile {
    # @{ ok; value; message; exists } (Read-EbiJson's).
    param([string]$WorkDir, [string]$RunId)
    return (Read-EbiJson -Path (Get-EbiRunFile -WorkDir $WorkDir -RunId $RunId))
}

function Set-EbiRunsAbandoned {
    <#
      Mark run.json documents finished (abandoned = true, abandonedAt) so
      Find-EbiUnfinishedRuns stops offering them -- the operator's "d"
      (forget) at the UNFINISHED RUN prompt: a run from another day that
      will never be resumed. Nothing else in run/<runId>/ is touched.
      Returns @{ ok; forgotten = @(runIds); message }.
    #>
    param([string]$WorkDir, $Runs)
    $done = New-Object System.Collections.ArrayList
    foreach ($doc in @($Runs)) {
        if (-not ($doc -is [System.Collections.IDictionary]) -or [string]::IsNullOrWhiteSpace([string]$doc['runId'])) { continue }
        $doc['finished'] = $true
        $doc['abandoned'] = $true
        $doc['abandonedAt'] = (Get-Date).ToString('o')
        $doc['updatedAt'] = (Get-Date).ToString('o')
        $w = Write-EbiJson -Path (Get-EbiRunFile -WorkDir $WorkDir -RunId ([string]$doc['runId'])) -Value $doc
        if (-not $w['ok']) { return @{ ok = $false; forgotten = $done.ToArray(); message = $w['message'] } }
        [void]$done.Add([string]$doc['runId'])
    }
    return @{ ok = $true; forgotten = $done.ToArray(); message = '' }
}

function Find-EbiUnfinishedRuns {
    <#
      Every run/<runId>/run.json under WorkDir with finished = false,
      newest first; optionally only those of one workflow id. For
      "ebi run --resume" without a run id (P1-10), and for the "there is
      an unfinished run, say --resume or start anew" prompt.
      Returns an array of the run.json hashtables.
    #>
    param([string]$WorkDir, [string]$WorkflowId = '')
    $out = New-Object System.Collections.ArrayList
    $runRoot = Join-Path $WorkDir 'run'
    if (-not (Test-Path -LiteralPath $runRoot)) { return $out.ToArray() }
    foreach ($dir in @(Get-ChildItem -LiteralPath $runRoot -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending)) {
        $r = Read-EbiJson -Path (Join-Path $dir.FullName 'run.json')
        if (-not $r['ok'] -or -not ($r['value'] -is [hashtable])) { continue }
        $doc = $r['value']
        if ($doc.Contains('finished') -and [bool]$doc['finished']) { continue }
        if ($WorkflowId -ne '') {
            $wid = if ($doc.Contains('workflow') -and ($doc['workflow'] -is [hashtable]) -and $doc['workflow'].Contains('id')) { [string]$doc['workflow']['id'] } else { '' }
            if ($wid -ne $WorkflowId) { continue }
        }
        [void]$out.Add($doc)
    }
    return $out.ToArray()
}
