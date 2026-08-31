# ============================================================
#  kernel/Trace.ps1
#
#  Append-only, machine-readable trace events. Dot-source only.
#
#  One JSON object per line, written to the run's own directory:
#
#      <WorkDir>/run/<RunId>/trace.jsonl
#
#  The path is fixed by spec/VOCABULARY.md 3.3 and is a sibling of
#  run/<RunId>/ledger.jsonl -- everything a single run produced stays
#  together, and "ebi trace <runId>" / "ebi bundle" can select one run
#  without reading every event ever written. Each event ALSO carries the
#  run id as a field, so a merged or concatenated file is still separable.
#
#  Fields:
#      ts         ISO 8601 round-trip timestamp. Named to match the ledger
#                 record shape in spec/STEP-CONTRACT.md 6.1, which is the
#                 file sitting next to this one in the same run directory.
#      runId      which run produced this event
#      phase      workflow / stage name
#      key        the item being processed (see spec/VOCABULARY.md 1)
#      tags       arbitrary workflow-specific dimensions; the tracer is
#                 never taught their names
#      action     what was attempted
#      status     ok | fail | skip | info | start
#      message    one human-readable sentence
#      data       OPTIONAL structured payload, omitted when not supplied
#
#  "data" is how a structured value reaches the trace WITHOUT being
#  flattened into "message": step warnings (spec/STEP-CONTRACT.md 3.1
#  requires the runner to write them here), a failed step's partial
#  outputs, timings, artifact paths. Anything put here must be
#  JSON-serializable -- handles and COM objects belong in $Ctx.Session
#  and never appear in a trace (spec/STEP-CONTRACT.md 3.4, which also says
#  the trace is where a step's outputs land). It is serialized
#  at depth 10; anything nested deeper than that is truncated by
#  ConvertTo-Json, so do not put a whole object graph in here.
#
#  UTF-8 is written without a BOM. In Windows PowerShell 5.1,
#  Set-Content -Encoding UTF8 would add a BOM and can corrupt JSONL when
#  used repeatedly, so writes use UTF8Encoding($false) directly.
# ============================================================

function Get-TraceRunId {
    # A blank RunId is a caller bug, but losing the trace is worse than
    # bucketing it: the bucket's name says what happened. Normalized in one
    # place so the directory and the event's own runId field never disagree.
    param([string]$RunId)
    if ([string]::IsNullOrWhiteSpace($RunId)) { return 'unknown' }
    return $RunId
}

function Get-TraceRunDir {
    param([string]$WorkDir, [string]$RunId)
    return (Join-Path (Join-Path $WorkDir 'run') (Get-TraceRunId $RunId))
}

function Get-TraceFile {
    param([string]$WorkDir, [string]$RunId)
    return (Join-Path (Get-TraceRunDir $WorkDir $RunId) 'trace.jsonl')
}

function Write-TraceEvent {
    param(
        [string]$WorkDir,
        [string]$RunId   = '',
        [string]$Phase   = '',
        [string]$Key     = '',
        [System.Collections.IDictionary]$Tags = @{},
        [string]$Action  = '',
        [string]$Status  = '',
        [string]$Message = '',
        [object]$Data    = $null
    )

    # Tracing must never make the operation being traced fail.
    try {
        if ([string]::IsNullOrWhiteSpace($WorkDir)) { return }

        $file = Get-TraceFile $WorkDir $RunId
        $dir  = Split-Path -Path $file -Parent
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }

        $evt = [ordered]@{
            ts        = (Get-Date).ToString('o')
            runId     = (Get-TraceRunId $RunId)
            phase     = $Phase
            key       = $Key
            tags      = if ($null -eq $Tags) { @{} } else { $Tags }
            action    = $Action
            status    = $Status
            message   = $Message
        }
        # Absent rather than null, so a reader can tell "no payload" from
        # "payload that happened to be null".
        if ($null -ne $Data) { $evt['data'] = $Data }

        $line = $evt | ConvertTo-Json -Compress -Depth 10
        $encoding = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText($file, $line + [Environment]::NewLine, $encoding)
    } catch {
        Write-Host ('  [trace WARN] {0}' -f $_.Exception.Message) -ForegroundColor DarkYellow
    }
}

function Read-TraceEvents {
    param(
        [string]$WorkDir,
        [string]$RunId = '',
        [int]$Tail = 0
    )

    if ([string]::IsNullOrWhiteSpace($WorkDir)) { return @() }

    $file = Get-TraceFile $WorkDir $RunId
    if (-not (Test-Path -LiteralPath $file)) { return @() }

    $lines = @(Get-Content -LiteralPath $file -Encoding UTF8 -ErrorAction SilentlyContinue)
    if ($Tail -gt 0 -and $lines.Count -gt $Tail) {
        $lines = @($lines | Select-Object -Last $Tail)
    }

    $events = [System.Collections.Generic.List[object]]::new()
    foreach ($line in $lines) {
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        try {
            $events.Add(($line | ConvertFrom-Json))
        } catch {
            # A partial final line can be observed while another process writes.
        }
    }
    return $events.ToArray()
}
