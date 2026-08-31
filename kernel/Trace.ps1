# ============================================================
#  kernel/Trace.ps1
#
#  Append-only, machine-readable trace events. Dot-source only.
#
#  Events are written one JSON object per line to <WorkDir>/trace.jsonl.
#  "key" identifies the item being processed; "tags" carries any
#  workflow-specific dimensions without teaching the tracer their names.
#
#  UTF-8 is written without a BOM. In Windows PowerShell 5.1,
#  Set-Content -Encoding UTF8 would add a BOM and can corrupt JSONL when
#  used repeatedly, so writes use UTF8Encoding($false) directly.
# ============================================================

function Get-TraceFile {
    param([string]$WorkDir)
    return (Join-Path $WorkDir 'trace.jsonl')
}

function Write-TraceEvent {
    param(
        [string]$WorkDir,
        [string]$Phase   = '',
        [string]$Key     = '',
        [System.Collections.IDictionary]$Tags = @{},
        [string]$Action  = '',
        [string]$Status  = '',
        [string]$Message = ''
    )

    # Tracing must never make the operation being traced fail.
    try {
        if ([string]::IsNullOrWhiteSpace($WorkDir)) { return }
        if (-not (Test-Path -LiteralPath $WorkDir)) {
            New-Item -ItemType Directory -Path $WorkDir -Force | Out-Null
        }

        $event = [ordered]@{
            timestamp = (Get-Date).ToString('o')
            phase     = $Phase
            key       = $Key
            tags      = if ($null -eq $Tags) { @{} } else { $Tags }
            action    = $Action
            status    = $Status
            message   = $Message
        }
        $line = $event | ConvertTo-Json -Compress -Depth 10
        $encoding = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText(
            (Get-TraceFile $WorkDir),
            $line + [Environment]::NewLine,
            $encoding
        )
    } catch {
        Write-Host ('  [trace WARN] {0}' -f $_.Exception.Message) -ForegroundColor DarkYellow
    }
}

function Read-TraceEvents {
    param(
        [string]$WorkDir,
        [int]$Tail = 0
    )

    $file = Get-TraceFile $WorkDir
    if (-not (Test-Path -LiteralPath $file)) { return @() }

    $lines = @(Get-Content -LiteralPath $file -Encoding UTF8 -ErrorAction SilentlyContinue)
    if ($Tail -gt 0 -and $lines.Count -gt $Tail) {
        $lines = @($lines[($lines.Count - $Tail)..($lines.Count - 1)])
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
