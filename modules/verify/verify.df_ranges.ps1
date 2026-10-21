# modules/verify/verify.df_ranges.ps1
# How many lines each DF picture must show so it covers the same records
# as the HOST pictures on the send-data sheet (10/21, the operator):
#   - HOST line n = DF line n - offset (1: HOST counts a first line the CSV
#     does not have), on both pictures;
#   - a file pair has a START picture (from 000001) and, when long, an END
#     picture that always runs to the HOST's last line;
#   - the DF file may run up to maxExtra lines past the HOST's last line
#     - offset (a trailing line, or a last line OCR missed under a red
#     box): the END picture runs on to the file's end;
#   - further off (the send data had empty lines the receive side
#     dropped), the numbers cannot be lined up: the default pictures are kept (16 lines
#     from the top, the END view's 17), no alignment.
# HOST pictures are taken in sheet order; a picture whose numbers start at
# 000001 opens the next file pair. Pure.

$Manifest = @{
  id         = 'verify.df_ranges'
  group      = 'verify'
  summary    = 'Per DF pair: the lines its first / END picture must show to match the HOST pictures'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    pairs    = @{ type='list'; required=$true; desc='the DF file pairs in capture order (track.json pairs: lines = the file''s line count)' }
    ranges   = @{ type='list'; default=@(); desc='screen.read_line_numbers ranges of the HOST pictures, sheet order' }
    offset   = @{ type='int';  default=1; desc='HOST line n = DF line n - offset' }
    maxLines = @{ type='int';  default=40; desc='a picture asked for more lines than this is not trusted (default kept)' }
    maxExtra = @{ type='int';  default=2; desc='DF lines past the HOST''s last line still lined up (a trailing line the HOST does not show, or a last line OCR missed under a red box); more = not lined up' }
  }
  outputs    = @{
    plans   = @{ type='list';   desc='@{ mode; first; end; reason } per pair: mode align (first / end = lines the first / END picture shows; end 0 = one picture) | free (no alignment: empty lines dropped) | default (nothing to align to)' }
    aligned = @{ type='int';    desc='pairs whose pictures follow the HOST ranges' }
    alert   = @{ type='string'; desc='one line for the gate when a pair could not be lined up for a reason a person should check, else empty' }
    reason  = @{ type='string'; desc='one line per pair, for the run log / gate' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.df_ranges'; with = @{ pairs = '{{steps.track.out.data.pairs}}'; ranges = '{{steps.hostLines.out.ranges}}' } }
  notes      = 'Pure. A pair with no START picture, an unreadable picture or more than two pictures keeps the default pictures, with an alert.'
}

function VerifyDfRanges-Sets {
    # PURE. Readable ranges -> sets @(@{ start; finish }) (finish $null when
    # the pair has one picture) + the count of unreadable pictures + whether
    # some set had more than two pictures.
    param($Ranges)
    $sets = New-Object System.Collections.ArrayList
    $bad = 0; $extra = $false
    $cur = $null
    foreach ($r in @($Ranges)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        if (-not [bool]$r['ok']) { $bad++; continue }
        if ([int]$r['first'] -eq 1) {
            if ($null -ne $cur) { [void]$sets.Add($cur) }
            $cur = @{ start = $r; finish = $null; more = $false }
        } elseif ($null -ne $cur) {
            if ($null -eq $cur['finish']) { $cur['finish'] = $r } else { $cur['more'] = $true; $extra = $true; $cur['finish'] = $r }
        } else { $bad++ }
    }
    if ($null -ne $cur) { [void]$sets.Add($cur) }
    return @{ sets = $sets.ToArray(); bad = $bad; extra = $extra }
}

function VerifyDfRanges-Plan {
    # PURE. One pair (line count N) + its HOST set -> @{ mode; first; end; reason }.
    param([int]$Lines, $Set, [int]$Offset = 1, [int]$MaxLines = 40, [int]$MaxExtra = 2)
    if ($null -eq $Set) { return @{ mode = 'default'; first = 0; end = 0; reason = 'no HOST picture from 000001 for this pair' } }
    if ([bool]$Set['more']) { return @{ mode = 'default'; first = 0; end = 0; reason = 'more than two HOST pictures for this pair (only a START and an END picture are lined up)' } }
    $s = $Set['start']; $f = $Set['finish']
    $first = [int]$s['last'] - $Offset
    if ($null -eq $f) {
        if ($first -eq $Lines) { return @{ mode = 'align'; first = $first; end = 0; reason = ('one HOST picture 000001-{0:000000} = the whole file ({1} lines)' -f $s['last'], $Lines) } }
        if ($first -gt $Lines) { return @{ mode = 'free'; first = 0; end = 0; reason = ('HOST {0:000000} lines but the file has {1}: not lined up (empty lines?)' -f $s['last'], $Lines) } }
        return @{ mode = 'default'; first = 0; end = 0; reason = ('only a START picture (to {0:000000}) for a {1}-line file: no END picture to line up' -f $s['last'], $Lines) }
    }
    $hostLast = [int]$f['last']
    # extra: DF lines after the HOST's last line. 0 = the usual. 1..MaxExtra
    # (10/21 KJODWWBV: HOST 007550-007568, the file 7568 lines): the file
    # carries a line the HOST does not show, or OCR missed the last line(s)
    # under a red box -- either way the END picture runs on to the file's
    # last line. Anything else: empty lines dropped, not lined up.
    $extra = $Lines - ($hostLast - $Offset)
    if ($extra -lt 0 -or $extra -gt $MaxExtra) {
        return @{ mode = 'free'; first = 0; end = 0; reason = ('HOST ends at {0:000000} but the file has {1} lines (not {2}): empty lines dropped on the receive side, not lined up' -f $hostLast, $Lines, ($hostLast - $Offset)) }
    }
    $end = [int]$f['last'] - [int]$f['first'] + 1 + $extra
    if ($first -lt 1 -or $end -lt 1 -or $first -gt $MaxLines -or $end -gt $MaxLines) {
        return @{ mode = 'default'; first = 0; end = 0; reason = ('HOST pictures 000001-{0:000000} / {1:000000}-{2:000000} would need {3} / {4} DF lines (limit {5})' -f $s['last'], $f['first'], $f['last'], $first, $end, $MaxLines) }
    }
    return @{ mode = 'align'; first = $first; end = $end; reason = ('HOST 000001-{0:000000} / {1:000000}-{2:000000} -> DF 00001-{3:00000} / {4:00000}-{5:00000}{6}' -f $s['last'], $f['first'], $f['last'], $first, ([int]$f['first'] - $Offset), $Lines, $(if ($extra -gt 0) { (' (the file runs ' + $extra + ' line(s) past the HOST''s last: the END picture shows them too)') } else { '' })) }
}

function Invoke-Step {
    param($In, $Ctx)
    $pairs = @(@($In['pairs']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $g = VerifyDfRanges-Sets -Ranges $In['ranges']
    $sets = @($g['sets'])
    $plans = New-Object System.Collections.ArrayList
    $why = New-Object System.Collections.ArrayList
    $alerts = New-Object System.Collections.ArrayList
    $aligned = 0
    $any = (@(@($In['ranges']) | Where-Object { $_ -is [System.Collections.IDictionary] }).Count -gt 0)
    for ($i = 0; $i -lt $pairs.Count; $i++) {
        $set = if ($i -lt $sets.Count) { $sets[$i] } else { $null }
        $p = VerifyDfRanges-Plan -Lines ([int]$pairs[$i]['lines']) -Set $set -Offset ([int]$In['offset']) -MaxLines ([int]$In['maxLines']) -MaxExtra ([int]$In['maxExtra'])
        [void]$plans.Add($p)
        if ($p['mode'] -eq 'align') { $aligned++ }
        $name = [string]$pairs[$i]['rightName']
        [void]$why.Add(('{0}: {1} -- {2}' -f $(if ($name) { $name } else { 'pair ' + ($i + 1) }), $p['mode'], $p['reason']))
        if ($p['mode'] -eq 'default' -and $any) { [void]$alerts.Add(('DF pictures of {0} NOT lined up with the HOST pictures: {1} -- check them by hand' -f $(if ($name) { $name } else { 'pair ' + ($i + 1) }), $p['reason'])) }
    }
    if ($any -and [int]$g['bad'] -gt 0) { [void]$alerts.Add(('{0} HOST picture(s) on the send-data sheet could not be read (no line numbers)' -f $g['bad'])) }
    if ($sets.Count -gt $pairs.Count) { [void]$alerts.Add(('{0} HOST picture set(s) but {1} DF pair(s)' -f $sets.Count, $pairs.Count)) }
    foreach ($w in $why) { $Ctx.Log.Info($w) }
    return @{ ok = $true; plans = $plans.ToArray(); aligned = $aligned; alert = ($alerts -join ' / '); reason = ($why -join ' / ') }
}
