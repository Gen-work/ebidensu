# ============================================================
#  Watch-MappingProgress.ps1
#
#  Read-only progress monitor. Lets the operator watch what the tools are
#  doing WITHOUT opening (and locking) mapping_<Owner>.csv in Excel/SAKURA.
#
#  How it stays lock-free (the operator's idea, spec section 4.6):
#    - It copies mapping_<Owner>.csv to status\mapping_snapshot_<Owner>.csv
#      and reads the COPY, so it never holds a handle on the live mapping.
#    - It tails status\progress.jsonl, which phases only ever append to.
#  Refresh happens on an interval; press Enter for an immediate refresh,
#  or run with -Once for a single snapshot.
#
#  Usage:
#    .\Watch-MappingProgress.ps1 -WorkDir C:\work\proj
#    .\Watch-MappingProgress.ps1 -WorkDir C:\work\proj -IntervalSec 3 -Tail 20
#    .\Watch-MappingProgress.ps1 -WorkDir C:\work\proj -Once
# ============================================================

param(
    [string]$WorkDir = '',
    [string]$Owner   = '',
    [int]$IntervalSec = 5,
    [int]$Tail        = 15,
    [switch]$Once
)

$ErrorActionPreference = 'Stop'
try {
    [Console]::OutputEncoding = [System.Text.UTF8Encoding]::new()
    $OutputEncoding = [System.Text.UTF8Encoding]::new()
} catch {}

$scriptDir = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $scriptDir 'ProgressLog.ps1')

if ([string]::IsNullOrWhiteSpace($WorkDir)) { $WorkDir = Read-Host 'WorkDir path' }
if (-not (Test-Path -LiteralPath $WorkDir)) {
    Write-Host "[ERROR] WorkDir not found: $WorkDir" -ForegroundColor Red; exit 1
}

$cfg = $null
try { $cfg = Import-PowerShellDataFile (Join-Path $scriptDir 'VerifyConfig.psd1') } catch {}

$mappingPath = Join-Path $WorkDir ('mapping_{0}.csv' -f $Owner)
$statusDir   = Join-Path $WorkDir 'status'
if (-not (Test-Path -LiteralPath $statusDir)) { New-Item -ItemType Directory -Path $statusDir -Force | Out-Null }
$snapshotPath = Join-Path $statusDir ('mapping_snapshot_{0}.csv' -f $Owner)

# Phases to summarise: snap-style fields + the bitmask phases.
function Get-WatchPhases {
    $phases = [System.Collections.Generic.List[object]]::new()
    if ($null -ne $cfg -and $cfg.ContainsKey('PhaseOrder')) {
        foreach ($p in $cfg.PhaseOrder) {
            $field = [string]$p.Field
            if ([string]::IsNullOrWhiteSpace($field)) { continue }
            $bit = 0
            if ($p.ContainsKey('BitValue')) { $bit = [int]$p.BitValue }
            $phases.Add([pscustomobject]@{ Key = [string]$p.Key; Field = $field; Bit = $bit })
        }
    }
    if ($phases.Count -eq 0) {
        foreach ($fn in @('GIFT_HM_snap','GIFT_MQ_snap','GIFT_Jenkins_snap','GFIX_HM_snap','GFIX_Jenkins_snap','GFIX_log','DF_snap')) {
            $phases.Add([pscustomobject]@{ Key = $fn; Field = $fn; Bit = 0 })
        }
    }
    return $phases
}
$watchPhases = Get-WatchPhases

function Read-MappingSnapshot {
    # Copy then read the copy, so we never hold the live mapping open.
    if (-not (Test-Path -LiteralPath $mappingPath)) { return @() }
    try {
        Copy-Item -LiteralPath $mappingPath -Destination $snapshotPath -Force
        return @(Import-Csv -LiteralPath $snapshotPath -Encoding UTF8)
    } catch {
        return @()
    }
}

function Test-FieldDone {
    param([string]$Value, [int]$Bit)
    if ($Bit -gt 0) {
        $v = 0; try { $v = [int]$Value } catch { $v = 0 }
        return (($v -band $Bit) -eq $Bit)
    }
    return ((-not [string]::IsNullOrEmpty($Value)) -and ($Value -ne '0'))
}

function Show-Once {
    $rows = Read-MappingSnapshot
    $total = $rows.Count

    Clear-Host
    Write-Host ('===== Mapping Progress  ({0}) =====' -f (Get-Date).ToString('HH:mm:ss')) -ForegroundColor Green
    Write-Host ("  WorkDir : {0}" -f $WorkDir)
    Write-Host ("  Mapping : mapping_{0}.csv   rows={1}" -f $Owner, $total) -ForegroundColor Cyan
    if ($total -eq 0) {
        Write-Host '  (no mapping rows yet)' -ForegroundColor DarkGray
    } else {
        Write-Host ''
        Write-Host ('  {0,-22} {1,8}   {2}' -f 'phase', 'done/all', 'bar') -ForegroundColor DarkGray
        foreach ($ph in $watchPhases) {
            $done = 0
            foreach ($r in $rows) {
                $val = ''
                if ($r.PSObject.Properties.Name -contains $ph.Field) { $val = [string]$r.($ph.Field) }
                if (Test-FieldDone $val $ph.Bit) { $done++ }
            }
            $pct = if ($total -gt 0) { [int](($done * 100) / $total) } else { 0 }
            $barLen = [int]($pct / 5)
            $bar = ('#' * $barLen).PadRight(20, '.')
            $color = if ($done -eq $total) { 'Green' } elseif ($done -eq 0) { 'DarkGray' } else { 'Yellow' }
            Write-Host ('  {0,-22} {1,4}/{2,-3}   {3} {4,3}%' -f $ph.Key, $done, $total, $bar, $pct) -ForegroundColor $color
        }
    }

    # recent progress events
    $events = @(Read-ProgressEvents -WorkDir $WorkDir -Tail $Tail)
    Write-Host ''
    Write-Host ('  ----- recent events (last {0}) -----' -f $Tail) -ForegroundColor DarkGray
    if ($events.Count -eq 0) {
        Write-Host '  (no progress.jsonl events yet)' -ForegroundColor DarkGray
    } else {
        foreach ($e in $events) {
            $ts = ''
            try { $ts = ([datetime]$e.timestamp).ToString('HH:mm:ss') } catch { $ts = [string]$e.timestamp }
            $st = [string]$e.status
            $color = switch ($st) { 'ok' { 'Green' } 'fail' { 'Red' } 'skip' { 'DarkGray' } default { 'Gray' } }
            Write-Host ('  {0}  {1,-16} {2,-9} {3,-6} {4}' -f $ts, $e.phase, $e.correl_id_s, $st, $e.message) -ForegroundColor $color
        }
    }
    Write-Host ''
    if (-not $Once) { Write-Host ('  (refresh every {0}s; Ctrl+C to stop)' -f $IntervalSec) -ForegroundColor DarkGray }
}

if ($Once) {
    Show-Once
    exit 0
}

while ($true) {
    Show-Once
    Start-Sleep -Seconds ([Math]::Max(1, $IntervalSec))
}
