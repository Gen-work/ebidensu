# modules/verify/verify.crosscheck.ps1
# One fact read from several sources: all agree -> ok; any disagreement ->
# unknown, with who disagreed and what each said (P2-10, Plan.md 6.3,
# INTERVIEW iron rule two). No majority vote, no "the plausible one", no
# 3/9 guessing: the v2.20 bug rewrote a correct 00:00:01 into 00:00:07 by
# picking, and this step exists so that never happens again. A single
# source is ok with a warning, because nothing was cross-checked.

. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')     # ConvertTo-EbiHalfWidth
. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')   # ConvertTo-EbiDateTime

$Manifest = @{
  id         = 'verify.crosscheck'
  group      = 'verify'
  summary    = 'Compare the same fact from several sources; any disagreement is unknown'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    readings     = @{ type='list';   required=$true; desc='[ { source, value }, ... ]' }
    compare      = @{ type='string'; default='equal'; enum=@('equal', 'numericEqual', 'timeWithinSec') }
    toleranceSec = @{ type='int';    default=0; desc='with timeWithinSec: how far apart two times may be' }
    normalize    = @{ type='list';   default=@(); desc='any of trim, fullwidth, thousands (applied before equal / numericEqual)' }
  }
  outputs    = @{
    code          = @{ type='string'; desc='ok | unknown' }
    disagreements = @{ type='list';   desc='[ { a, aValue, b, bValue }, ... ] every pair that differs' }
    sources       = @{ type='int' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.crosscheck'; with = @{ readings = @(@{ source = 'page'; value = '{{steps.row.out.record.endTime}}' }, @{ source = 'log'; value = '{{steps.logrec.out.record.endTime}}' }); compare = 'timeWithinSec'; toleranceSec = 60 } }
  notes      = 'unknown is a verdict for human.gate, never a failure. Values that cannot be read as a number / time under numericEqual / timeWithinSec count as a disagreement, not as a pass.'
}

function VerifyCrosscheck-Normalize {
    # PURE. Apply the requested normalizations to one text.
    param([string]$Value, $Normalize)
    $v = if ($null -eq $Value) { '' } else { $Value }
    $n = @(@($Normalize) | ForEach-Object { [string]$_ })
    if ($n -contains 'trim') { $v = [regex]::Replace($v.Trim(), '\s+', ' ') }
    if ($n -contains 'fullwidth') { $v = ConvertTo-EbiHalfWidth -Value $v }
    if ($n -contains 'thousands') { $v = $v.Replace(',', '') }
    return $v
}

function VerifyCrosscheck-Same {
    # PURE. Two normalized texts under a compare mode -> $true / $false.
    param([string]$A, [string]$B, [string]$Compare, [int]$ToleranceSec)
    switch ($Compare) {
        'equal' { return [string]::Equals($A, $B, [System.StringComparison]::Ordinal) }
        'numericEqual' {
            $x = 0.0; $y = 0.0
            $ci = [System.Globalization.CultureInfo]::InvariantCulture; $st = [System.Globalization.NumberStyles]::Float
            if (-not [double]::TryParse($A.Replace(',', ''), $st, $ci, [ref]$x) -or -not [double]::TryParse($B.Replace(',', ''), $st, $ci, [ref]$y)) { return $false }
            return ($x -eq $y)
        }
        'timeWithinSec' {
            $ta = ConvertTo-EbiDateTime -Text $A; $tb = ConvertTo-EbiDateTime -Text $B
            if (-not $ta['ok'] -or -not $tb['ok']) { return $false }
            return ([Math]::Abs(($ta['value'] - $tb['value']).TotalSeconds) -le $ToleranceSec)
        }
    }
    return $false
}

function VerifyCrosscheck-Run {
    # PURE. readings -> @{ code; disagreements; sources }
    param($Readings, [string]$Compare, [int]$ToleranceSec, $Normalize)
    $items = New-Object System.Collections.ArrayList
    foreach ($r in @($Readings)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $src = if ($r.Contains('source')) { [string]$r['source'] } else { ('#' + ($items.Count + 1)) }
        $val = if ($r.Contains('value') -and $null -ne $r['value']) { [string]$r['value'] } else { '' }
        [void]$items.Add(@{ source = $src; raw = $val; value = (VerifyCrosscheck-Normalize -Value $val -Normalize $Normalize) })
    }
    $dis = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $items.Count; $i++) {
        for ($j = $i + 1; $j -lt $items.Count; $j++) {
            if (-not (VerifyCrosscheck-Same -A $items[$i]['value'] -B $items[$j]['value'] -Compare $Compare -ToleranceSec $ToleranceSec)) {
                [void]$dis.Add(@{ a = $items[$i]['source']; aValue = $items[$i]['raw']; b = $items[$j]['source']; bValue = $items[$j]['raw'] })
            }
        }
    }
    return @{ code = $(if ($dis.Count -eq 0) { 'ok' } else { 'unknown' }); disagreements = $dis.ToArray(); sources = $items.Count }
}

function Invoke-Step {
    param($In, $Ctx)
    $readings = @($In['readings'])
    if ($readings.Count -eq 0) { return @{ ok = $false; failure = 'input_invalid'; message = 'readings is empty'; code = ''; disagreements = @(); sources = 0 } }
    $r = VerifyCrosscheck-Run -Readings $readings -Compare ([string]$In['compare']) -ToleranceSec ([int]$In['toleranceSec']) -Normalize $In['normalize']
    $warnings = @()
    if ($r['sources'] -eq 1) { $warnings = @( @{ code = 'single_source'; message = 'only one source: nothing was cross-checked'; data = @{} } ) }
    return @{ ok = $true; code = $r['code']; disagreements = $r['disagreements']; sources = $r['sources']; warnings = $warnings }
}
