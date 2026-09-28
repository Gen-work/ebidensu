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
#  operator / workDir / timeWindow) plus workflow id/version, the CLI
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

function Write-EbiRunFile {
    # run.json: the run.* scope + what was run and how. Atomic (Write-EbiJson).
    param([string]$WorkDir, [string]$RunId, [hashtable]$Run, [hashtable]$Workflow, [hashtable]$RunArgs, [bool]$Finished = $false, $Result = $null)
    $doc = @{
        runId      = $RunId
        startedAt  = $Run['startedAt']
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
