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
#  SS_CODE default: 5th char of Correl_ID_S. .NET indexing is 0-based,
#  so this is Substring(4,1). e.g. JIDSF48S -> index 4 = 'F'.
#  Exception: J<biz>LxxS receive jobs use SS_CODE 'J' instead of 'L'.
#  Optional mapping-provided SS_CODE overrides both rules.
# ============================================================

function Get-GfixSsCode {
    param(
        [string]$CorrelIdS,
        [string]$SsCodeOverride = ''
    )
    if (-not [string]::IsNullOrWhiteSpace($SsCodeOverride)) { return $SsCodeOverride.Trim() }
    if ([string]::IsNullOrEmpty($CorrelIdS) -or $CorrelIdS.Length -lt 5) { return '' }
    if ($CorrelIdS -match '^J[A-Za-z0-9]{3}L[A-Za-z0-9]{2}S$') { return 'J' }
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
    param([string]$ToCode, [string]$CorrelIdS, [string]$SsCode = '')
    $path = Get-GfixExpectedPath $ToCode $CorrelIdS
    $ss   = Get-GfixSsCode $CorrelIdS $SsCode
    if ([string]::IsNullOrEmpty($ss)) { return $path }
    return ('{0} {1}' -f $path, $ss)
}

# Command paths may append the transfer timestamp to Correl_ID_S, for example
# JIDSU86S.260729.10515511. ReplaceEvidence must accept that spelling as the
# same receive command while retaining the strict trailing SS-code boundary.
function Get-GfixExpectedCommandPattern {
    param([string]$ToCode, [string]$CorrelIdS, [string]$SsCode = '')
    $base = $CorrelIdS
    if ($base -match '^(?<base>.+)\.\d{6}\.\d{8}$') { $base = [string]$Matches['base'] }
    $path = Get-GfixExpectedPath $ToCode $base
    $ss = Get-GfixSsCode $base $SsCode
    $pattern = [regex]::Escape($path) + '(?:\.\d{6}\.\d{8})?'
    if (-not [string]::IsNullOrEmpty($ss)) { $pattern += '\s+' + [regex]::Escape($ss) + '(?=\s|[''"]|$)' }
    return $pattern
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

# The correl token the Command: line itself points at, INCLUDING the transfer
# batch stamp when the command carries one:
#   '... /appl/IDS/IDSVer1/gfix/recv/JIDSU86S.260729.10515511 F'
#     -> 'JIDSU86S.260729.10515511'
#   '... /appl/IDS/IDSVer1/gfix/recv/JIDSU86S F'   -> 'JIDSU86S'
# Returns '' when the line carries no recv path for this TO_CODE. Pure string
# work -- this is what tells two candidate logs apart, since the FILE name is
# only our own download naming and says nothing about which run it recorded.
function Get-GfixCommandCorrelToken {
    param([string]$Line, [string]$ToCode)
    if ([string]::IsNullOrEmpty($Line) -or [string]::IsNullOrWhiteSpace($ToCode)) { return '' }
    $prefix = ('/appl/{0}/{0}Ver1/gfix/recv/' -f $ToCode)
    $m = [regex]::Match($Line, ([regex]::Escape($prefix) + '([A-Za-z0-9]+(?:\.\d{6}\.\d{8})?)'))
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

# The transfer batch stamp of a correl id ('.YYMMDD.NNNNNNNN'), or '' when the
# id is the plain spelling.
function Get-GfixCorrelStamp {
    param([string]$CorrelId)
    if ([string]::IsNullOrWhiteSpace($CorrelId)) { return '' }
    if ($CorrelId.Trim() -match '\.(?<stamp>\d{6}\.\d{8})$') { return [string]$Matches['stamp'] }
    return ''
}

# Identity of the receive RUN a candidate log recorded. Two files carrying the
# same command token and the same log timestamp are the same run downloaded
# twice under two different names -- which is exactly what happens when a
# folder holds both '<correl>_x.log' (legacy naming) and
# '<correl>.<stamp>_x.log' (timestamped naming) for one job. That is a
# duplicate, NOT the "several candidates, chose newest" ambiguity worth
# warning about.
function Get-GfixLogRunKey {
    param([string]$CorrelToken, $Timestamp, [string]$CommandLine)
    $tok = [string]$CorrelToken
    if ([string]::IsNullOrWhiteSpace($tok)) { $tok = ([string]$CommandLine).Trim() }
    $ts = if ($null -ne $Timestamp) { ([datetime]$Timestamp).ToString('yyyy-MM-dd HH:mm:ss') } else { '' }
    return ('{0}|{1}' -f $tok, $ts)
}

# ---------------------------------------------------------------------------
# Select-GfixLogCandidate
#   Narrows the logs whose Command: line matched down to the ONE receive run
#   this correl means, and says whether anything genuinely ambiguous is left.
#
#   1. EXACT RUN WINS. When Correl_ID_S carries a batch stamp and some
#      candidate's command token carries that same stamp, only those
#      candidates survive -- the mapping row named a specific transfer, so a
#      different run of the same job is not a competing answer.
#   2. DUPLICATE SPELLINGS COLLAPSE. Remaining candidates are grouped by run
#      identity (Get-GfixLogRunKey); each group keeps one representative,
#      preferring the file whose NAME carries the batch stamp (the current
#      download convention) so reruns stay reproducible.
#   3. ONLY REAL AMBIGUITY WARNS. A warning is produced when more than one
#      distinct RUN is left -- genuine retries of the same job, which is the
#      case that deserves the operator's attention.
#
#   Candidates: @( @{ File; Timestamp; CommandLine; CorrelToken; ... } ... )
#   Returns @{ Chosen; Runs; Duplicates; Warning } -- Runs is the deduped set
#   newest-first, Duplicates the count collapsed in step 2.
# ---------------------------------------------------------------------------
function Select-GfixLogCandidate {
    param([object[]]$Candidates, [string]$CorrelIdS)

    $res = @{ Chosen = $null; Runs = @(); Duplicates = 0; Warning = '' }
    $all = @($Candidates | Where-Object { $null -ne $_ })
    if ($all.Count -eq 0) { return $res }

    # 1. exact-run preference
    $wantStamp = Get-GfixCorrelStamp $CorrelIdS
    if (-not [string]::IsNullOrWhiteSpace($wantStamp)) {
        $exact = @($all | Where-Object { (Get-GfixCorrelStamp ([string]$_.CorrelToken)) -eq $wantStamp })
        if ($exact.Count -gt 0) { $all = $exact }
    }

    # 2. collapse duplicate spellings of one run
    $groups = [ordered]@{}
    foreach ($c in $all) {
        $key = Get-GfixLogRunKey -CorrelToken ([string]$c.CorrelToken) -Timestamp $c.Timestamp -CommandLine ([string]$c.CommandLine)
        if (-not $groups.Contains($key)) {
            $groups[$key] = $c
            continue
        }
        # Keep the stamped file name; it is the naming the downloader emits
        # now, so a rerun picks the same file every time. The stamp is matched
        # ANYWHERE in the name -- downloads carry a suffix after it
        # ('JIDSU86S.260729.10515511_a.log'), so an end-anchored test finds
        # nothing.
        $stampRx = '\.\d{6}\.\d{8}'
        $kept = $groups[$key]
        $keptStamped = ([System.IO.Path]::GetFileName([string]$kept.File)) -match $stampRx
        $thisStamped = ([System.IO.Path]::GetFileName([string]$c.File)) -match $stampRx
        if ($thisStamped -and -not $keptStamped) { $groups[$key] = $c }
        $res.Duplicates++
    }

    $runs = @($groups.Values | Sort-Object -Property @{
        Expression = { if ($_.Timestamp) { $_.Timestamp } else { [datetime]::MinValue } }
    } -Descending)
    $res.Runs = $runs
    $res.Chosen = $runs[0]

    # 3. warn only on genuinely different runs
    if ($runs.Count -gt 1) {
        $names = (@($runs | ForEach-Object { [System.IO.Path]::GetFileName([string]$_.File) }) -join ', ')
        $res.Warning = ('{0} different receive runs matched; chose newest ({1}). candidates: {2}' -f `
            $runs.Count, [System.IO.Path]::GetFileName([string]$res.Chosen.File), $names)
    }
    return $res
}

function Test-GfixCommandLine {
    param([string]$Line, [string]$Fragment, [string]$Pattern = '')
    if ([string]::IsNullOrEmpty($Line)) { return $false }
    if ($Line -notmatch 'Command:') { return $false }
    if (-not [string]::IsNullOrWhiteSpace($Pattern)) { return [regex]::IsMatch($Line, $Pattern) }
    return $Line.Contains($Fragment)
}

# Scans -LogDir for the receive log of one correl id.
#   Returns a PSCustomObject:
#     CorrelIdS, Fragment,
#     Candidates : @( @{File;Timestamp;CommandLine;CorrelToken} ... ) -- every
#                  file whose Command: line matched, before deduplication
#     Runs       : the same set after duplicate spellings of ONE run were
#                  collapsed (Select-GfixLogCandidate), newest first
#     Duplicates : how many candidates were collapsed as same-run duplicates
#     Chosen     : @{File;Timestamp;CommandLine;Lines}  (newest run, or $null)
#     Warning    : non-empty ONLY when more than one distinct RUN matched
#     Error      : non-empty when 0 candidates / bad input (caller fails the row)
#
# Every '*.log' is inspected because both '<Correl_ID_S>_*.log' and
# '<Correl_ID_S>.<timestamp>_*.log' are valid names. That is also why the
# duplicate collapse exists: a folder holding BOTH spellings of one download
# used to look like two competing candidates and warned on every run, even
# though it always went on to pick the right file. Run identity comes from the
# log's own Command: line, never from our download naming. The whole chosen
# file's lines are returned so the Excel step can paste the entire log.
function Find-GfixLogForCorrel {
    param(
        [string]$LogDir,
        [string]$ToCode,
        [string]$CorrelIdS,
        [string]$SsCode = ''
    )
    $result = [ordered]@{
        CorrelIdS  = $CorrelIdS
        Fragment   = ''
        Candidates = @()
        # Candidates after duplicate spellings of one run were collapsed,
        # newest first. Runs.Count > 1 is the only genuine ambiguity.
        Runs       = @()
        Duplicates = 0
        Chosen     = $null
        Warning    = ''
        Error      = ''
    }
    if ([string]::IsNullOrWhiteSpace($CorrelIdS)) {
        $result.Error = 'empty Correl_ID_S'
        return [pscustomobject]$result
    }
    $result.Fragment = Get-GfixExpectedCommandFragment $ToCode $CorrelIdS $SsCode
    $commandPattern = Get-GfixExpectedCommandPattern $ToCode $CorrelIdS $SsCode
    if (-not (Test-Path -LiteralPath $LogDir)) {
        $result.Error = "log dir not found: $LogDir"
        return [pscustomobject]$result
    }

    # Do not restrict this to '<correl>_*.log': timestamped downloads commonly
    # use '<correl>.<timestamp>_*.log', and a legacy exact-name log may coexist.
    $files = @(Get-ChildItem -LiteralPath $LogDir -Filter '*.log' -File -ErrorAction SilentlyContinue)

    $cands = [System.Collections.Generic.List[object]]::new()
    foreach ($f in $files) {
        $lines = $null
        try { $lines = @(Get-Content -LiteralPath $f.FullName -Encoding UTF8 -ErrorAction Stop) }
        catch { continue }
        foreach ($ln in $lines) {
            if (Test-GfixCommandLine $ln $result.Fragment $commandPattern) {
                $cands.Add([pscustomobject]@{
                    File        = $f.FullName
                    Timestamp   = (Get-GfixLogTimestamp $ln)
                    CommandLine = $ln
                    CorrelToken = (Get-GfixCommandCorrelToken $ln $ToCode)
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

    # Which candidate, and is anything really ambiguous? The plain and
    # batch-stamped spellings of ONE download are collapsed here instead of
    # being reported as competing candidates.
    $pick = Select-GfixLogCandidate -Candidates $cands.ToArray() -CorrelIdS $CorrelIdS
    $result.Chosen     = $pick.Chosen
    $result.Runs       = $pick.Runs
    $result.Duplicates = [int]$pick.Duplicates
    $result.Warning    = [string]$pick.Warning
    return [pscustomobject]$result
}
