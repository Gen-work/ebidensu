# ============================================================
#  ProgressLog.ps1
#
#  Append-only progress event log so the operator can watch what the
#  tools are doing WITHOUT opening (and locking) mapping_<Owner>.csv.
#  Dot-source only (no param() block).
#
#  Each phase appends one JSON line per row it touches to:
#      <WorkDir>\status\progress.jsonl
#  Fields: timestamp, phase, correl_id_s, job_name, action, status, message.
#
#  Watch-MappingProgress.ps1 tails this file plus a read-only copy of the
#  mapping, so nothing here ever holds a write lock on the CSV.
#
#  Encoding: the .jsonl is written UTF-8 *without* BOM (PS 5.1's
#  Set-Content -Encoding UTF8 would prepend a BOM on every append, which
#  corrupts a line-delimited file -- so we use UTF8Encoding($false)).
# ============================================================

function Get-ProgressDir {
    param([string]$WorkDir)
    return (Join-Path $WorkDir 'status')
}

function Get-ProgressFile {
    param([string]$WorkDir)
    return (Join-Path (Get-ProgressDir $WorkDir) 'progress.jsonl')
}

function Write-ProgressEvent {
    param(
        [string]$WorkDir,
        [string]$Phase,
        [string]$CorrelIdS = '',
        [string]$JobName   = '',
        [string]$Action    = '',
        [string]$Status    = '',   # ok | fail | skip | info | start
        [string]$Message   = ''
    )
    # Progress logging must never break the main phase.
    try {
        if ([string]::IsNullOrWhiteSpace($WorkDir)) { return }
        $dir = Get-ProgressDir $WorkDir
        if (-not (Test-Path -LiteralPath $dir)) {
            New-Item -ItemType Directory -Path $dir -Force | Out-Null
        }
        $obj = [ordered]@{
            timestamp   = (Get-Date).ToString('o')
            phase       = $Phase
            correl_id_s = $CorrelIdS
            job_name    = $JobName
            action      = $Action
            status      = $Status
            message     = $Message
        }
        $line = ($obj | ConvertTo-Json -Compress -Depth 4)
        $enc  = New-Object System.Text.UTF8Encoding($false)
        [System.IO.File]::AppendAllText((Get-ProgressFile $WorkDir), $line + [Environment]::NewLine, $enc)
    } catch {
        Write-Host ('  [progress-log WARN] {0}' -f $_.Exception.Message) -ForegroundColor DarkYellow
    }
}

function Read-ProgressEvents {
    param([string]$WorkDir, [int]$Tail = 0)
    $file = Get-ProgressFile $WorkDir
    if (-not (Test-Path -LiteralPath $file)) { return @() }
    $lines = @(Get-Content -LiteralPath $file -Encoding UTF8 -ErrorAction SilentlyContinue)
    if ($Tail -gt 0 -and $lines.Count -gt $Tail) {
        $lines = @($lines[($lines.Count - $Tail)..($lines.Count - 1)])
    }
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($ln in $lines) {
        if ([string]::IsNullOrWhiteSpace($ln)) { continue }
        try { $out.Add(($ln | ConvertFrom-Json)) } catch {}
    }
    return $out.ToArray()
}
