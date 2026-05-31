# ============================================================
#  GfixLog.ps1
#
#  PURE GFIX receive-log matching -- NO Excel COM, NO mapping I/O.
#  Kept separate so the matcher can be unit-tested with plain log
#  fixtures (Tests\Test-GfixLog.ps1) before the Excel paste step is
#  ever wired in. Dot-source only (no param() block).
#
#  Match target (per spec section 9):
#    Command line containing:
#      /appl/<TO_CODE>/<TO_CODE>Ver1/gfix/recv/<Correl_ID_S> <SS_CODE>
#    e.g. for TO=IDS, Correl_ID_S=JIDSF48S, SS_CODE=F :
#      /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F
#    appearing on a line like:
#      2026-05-29 10:59:29 INFO Command: '/appl/IDS/shell/IDSLB053run.sh
#        /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F'
#
#  SS_CODE: 5th char of Correl_ID_S. .NET indexing is 0-based, so this is
#  Substring(4,1). e.g. JIDSF48S -> index 4 = 'F'.
# ============================================================

function Get-GfixSsCode {
    param([string]$CorrelIdS)
    if ([string]::IsNullOrEmpty($CorrelIdS) -or $CorrelIdS.Length -lt 5) { return '' }
    return $CorrelIdS.Substring(4, 1)
}

function Get-GfixExpectedPath {
    param([string]$ToCode, [string]$CorrelIdS)
    return ('/appl/{0}/{0}Ver1/gfix/recv/{1}' -f $ToCode, $CorrelIdS)
}

# The substring we look for inside a Command: line. Includes the trailing
# " <SS_CODE>" when the SS code is known, so we do not match a sibling
# correl id that is a prefix of another.
function Get-GfixExpectedCommandFragment {
    param([string]$ToCode, [string]$CorrelIdS)
    $path = Get-GfixExpectedPath $ToCode $CorrelIdS
    $ss   = Get-GfixSsCode $CorrelIdS
    if ([string]::IsNullOrEmpty($ss)) { return $path }
    return ('{0} {1}' -f $path, $ss)
}

# Parse a leading 'yyyy-MM-dd HH:mm:ss' timestamp. Returns [datetime] or $null.
function Get-GfixLogTimestamp {
    param([string]$Line)
    if ([string]::IsNullOrEmpty($Line)) { return $null }
    $m = [regex]::Match($Line, '(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2}:\d{2})')
    if (-not $m.Success) { return $null }
    $dt = [datetime]::MinValue
    if ([datetime]::TryParse($m.Groups[1].Value, [ref]$dt)) { return $dt }
    return $null
}

function Test-GfixCommandLine {
    param([string]$Line, [string]$Fragment)
    if ([string]::IsNullOrEmpty($Line)) { return $false }
    if ($Line -notmatch 'Command:') { return $false }
    return $Line.Contains($Fragment)
}

# Scans -LogDir for the receive log of one correl id.
#   Returns a PSCustomObject:
#     CorrelIdS, Fragment,
#     Candidates : @( @{File;Timestamp;CommandLine} ... )
#     Chosen     : @{File;Timestamp;CommandLine;Lines}  (newest, or $null)
#     Warning    : non-empty when >1 candidate matched
#     Error      : non-empty when 0 candidates / bad input (caller fails the row)
#
# File preference: '<Correl_ID_S>_*.log' first (named by GfixLogDownload),
# else any '*.log'. The whole chosen file's lines are returned so the Excel
# step can paste the entire log (not just the Command line).
function Find-GfixLogForCorrel {
    param(
        [string]$LogDir,
        [string]$ToCode,
        [string]$CorrelIdS
    )
    $result = [ordered]@{
        CorrelIdS  = $CorrelIdS
        Fragment   = ''
        Candidates = @()
        Chosen     = $null
        Warning    = ''
        Error      = ''
    }
    if ([string]::IsNullOrWhiteSpace($CorrelIdS)) {
        $result.Error = 'empty Correl_ID_S'
        return [pscustomobject]$result
    }
    $result.Fragment = Get-GfixExpectedCommandFragment $ToCode $CorrelIdS
    if (-not (Test-Path -LiteralPath $LogDir)) {
        $result.Error = "log dir not found: $LogDir"
        return [pscustomobject]$result
    }

    $files = @(Get-ChildItem -LiteralPath $LogDir -Filter ('{0}_*.log' -f $CorrelIdS) -File -ErrorAction SilentlyContinue)
    if ($files.Count -eq 0) {
        $files = @(Get-ChildItem -LiteralPath $LogDir -Filter '*.log' -File -ErrorAction SilentlyContinue)
    }

    $cands = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $files) {
        $lines = $null
        try { $lines = @(Get-Content -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop) }
        catch { continue }
        foreach ($ln in $lines) {
            if (Test-GfixCommandLine $ln $result.Fragment) {
                $cands.Add([pscustomobject]@{
                    File        = $f.FullName
                    Timestamp   = (Get-GfixLogTimestamp $ln)
                    CommandLine = $ln
                    Lines       = $lines
                })
                break
            }
        }
    }

    $result.Candidates = $cands.ToArray()
    if ($cands.Count -eq 0) {
        $result.Error = ('no log matches command fragment: {0}' -f $result.Fragment)
        return [pscustomobject]$result
    }

    $sorted = @($cands | Sort-Object -Property @{
        Expression = { if ($_.Timestamp) { $_.Timestamp } else { [datetime]::MinValue } }
    } -Descending)
    $result.Chosen = $sorted[0]

    if ($cands.Count -gt 1) {
        $names = (@($cands | ForEach-Object { Split-Path -Leaf $_.File }) -join ', ')
        $result.Warning = ('{0} logs matched; chose newest ({1}). candidates: {2}' -f `
            $cands.Count, (Split-Path -Leaf $result.Chosen.File), $names)
    }
    return [pscustomobject]$result
}
