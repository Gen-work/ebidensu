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
# 000001 opens the next file pair. A picture with no line numbers is a
# HOST view of an empty data set (10/26 K/LJODWAF1: one job, two I/F --
# STEP ..20 with 678 records, STEP ..30 empty): it takes the 0-line pair.
# The pairs come from track in file-name order, which need not be the
# send-data sheet's order (there 0006 = the empty one came first): each
# HOST slot takes the pair whose line count fits it, so the DF pictures
# are stacked -- and labelled by excel.label_sections -- in the send-data
# sheet's order. Pure.

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
    pairs    = @{ type='list'; required=$true; desc='the DF file pairs as track listed them (track.json pairs: lines = the file''s line count)' }
    shots    = @{ type='list'; default=@(); desc='track''s DF shots, one per pair in the same order (track.json dfShots); handed back in the new order for screen.launch_capture reuse' }
    ranges   = @{ type='list'; default=@(); desc='screen.read_line_numbers ranges of the HOST pictures, sheet order' }
    offset   = @{ type='int';  default=1; desc='HOST line n = DF line n - offset' }
    maxLines = @{ type='int';  default=40; desc='a picture asked for more lines than this is not trusted (default kept)' }
    maxExtra = @{ type='int';  default=2; desc='DF lines past the HOST''s last line still lined up (a trailing line the HOST does not show, or a last line OCR missed under a red box); more = not lined up' }
  }
  outputs    = @{
    pairs   = @{ type='list';   desc='the pairs in the send-data sheet''s order (pairs no HOST slot took come last, as listed)' }
    shots   = @{ type='list';   desc='the shots input in that same order' }
    order   = @{ type='list';   desc='for each output pair, its index in the input (0-based)' }
    moved   = @{ type='bool';   desc='the order differs from the input' }
    plans   = @{ type='list';   desc='@{ mode; first; end; reason } per output pair: mode align (first / end = lines the first / END picture shows; end 0 = one picture) | empty (a 0-line file under a HOST picture with no line numbers) | free (no alignment: empty lines dropped) | default (nothing to align to)' }
    aligned = @{ type='int';    desc='pairs whose pictures follow the HOST ranges' }
    alert   = @{ type='string'; desc='one line for the gate when a pair could not be lined up for a reason a person should check, else empty' }
    reason  = @{ type='string'; desc='one line per pair, for the run log / gate' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.df_ranges'; with = @{ pairs = '{{steps.track.out.data.pairs}}'; ranges = '{{steps.hostLines.out.ranges}}'; shots = '{{steps.track.out.data.dfShots}}' } }
  notes      = 'Pure. A pair with no START picture, an unreadable picture (unless the pair is an empty file) or more than two pictures keeps the default pictures, with an alert. Feed pairs / shots out of this step to screen.launch_capture so the plans line up with them.'
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
        # a START picture runs from 000001; OCR may miss its top line or two
        # (10/21 SJODWDBC read from 000002), so 2..3 opens a pair too when no
        # pair is waiting for its END picture
        $opens = ([int]$r['first'] -eq 1) -or ([int]$r['first'] -le 3 -and ($null -eq $cur -or $null -ne $cur['finish']))
        if ($opens) {
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
        # the whole file in one shot; the file may carry up to MaxExtra lines
        # the HOST does not show (10/21 RJODWDBC: HOST 000001-000002, 2 lines)
        $x1 = $Lines - $first
        if ($x1 -ge 0 -and $x1 -le $MaxExtra) { return @{ mode = 'align'; first = $Lines; end = 0; reason = ('one HOST picture 000001-{0:000000} = the whole file ({1} lines{2})' -f $s['last'], $Lines, $(if ($x1 -gt 0) { ', ' + $x1 + ' past the HOST''s last' } else { '' })) } }
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

function VerifyDfRanges-Slots {
    # PURE. Ranges in sheet order -> slots @(@{ kind = 'set'; set } |
    # @{ kind = 'blank' }): the readable pictures grouped as in
    # VerifyDfRanges-Sets, and each picture with no line numbers a slot of
    # its own (it closes the pair before it). Also extra (some set had more
    # than two pictures) and stray (readable pictures that opened nothing).
    param($Ranges)
    $slots = New-Object System.Collections.ArrayList
    $stray = 0; $extra = $false
    $runs = New-Object System.Collections.ArrayList
    $run = New-Object System.Collections.ArrayList
    foreach ($r in @($Ranges)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        if ([bool]$r['ok']) { [void]$run.Add($r); continue }
        [void]$runs.Add(@{ kind = 'run'; ranges = $run.ToArray() }); $run = New-Object System.Collections.ArrayList
        [void]$runs.Add(@{ kind = 'blank' })
    }
    [void]$runs.Add(@{ kind = 'run'; ranges = $run.ToArray() })
    foreach ($x in $runs) {
        if ([string]$x['kind'] -eq 'blank') { [void]$slots.Add(@{ kind = 'blank' }); continue }
        if (@($x['ranges']).Count -eq 0) { continue }
        $g = VerifyDfRanges-Sets -Ranges $x['ranges']
        foreach ($st in @($g['sets'])) { [void]$slots.Add(@{ kind = 'set'; set = $st }) }
        $stray += [int]$g['bad']
        if ([bool]$g['extra']) { $extra = $true }
    }
    return @{ slots = $slots.ToArray(); stray = $stray; extra = $extra }
}

function VerifyDfRanges-Assign {
    # PURE. Slots + pair line counts -> @{ order = pair index per slot (-1 =
    # no pair), rest = pair indexes no slot took, in input order }.
    # 1. a set slot takes the first free pair whose line count fits its
    #    HOST last line (- offset, up to maxExtra more);
    # 2. a blank slot takes the first free 0-line pair;
    # 3. slots still empty take the free pairs left, in order (what they
    #    cannot line up the plan says);
    # so pairs that fit keep the sheet's order and nothing is dropped.
    # sure: every pair with lines was placed under a readable HOST set by
    # its count -- else (10/26 SJODWAF1: all three HOST pictures unreadable,
    # the data ones too) the caller keeps the old order, since a blank slot
    # may as well be an unread data picture.
    param($Slots, [int[]]$Lines, [int]$Offset = 1, [int]$MaxExtra = 2)
    $slotsA = @($Slots)
    $order = New-Object int[] $slotsA.Count
    for ($k = 0; $k -lt $order.Length; $k++) { $order[$k] = -1 }
    $used = New-Object bool[] $Lines.Length
    $byCount = 0
    for ($k = 0; $k -lt $slotsA.Count; $k++) {
        $sl = $slotsA[$k]
        for ($i = 0; $i -lt $Lines.Length; $i++) {
            if ($used[$i]) { continue }
            $fits = $false
            if ([string]$sl['kind'] -eq 'blank') { $fits = ($Lines[$i] -eq 0) }
            else {
                $st = $sl['set']
                $hostLast = if ($null -ne $st['finish']) { [int]$st['finish']['last'] } else { [int]$st['start']['last'] }
                $x = $Lines[$i] - ($hostLast - $Offset)
                $fits = ($x -ge 0 -and $x -le $MaxExtra)
            }
            if ($fits) { $order[$k] = $i; $used[$i] = $true; if ($Lines[$i] -gt 0) { $byCount++ }; break }
        }
    }
    for ($k = 0; $k -lt $slotsA.Count; $k++) {
        if ($order[$k] -ge 0) { continue }
        for ($i = 0; $i -lt $Lines.Length; $i++) { if (-not $used[$i]) { $order[$k] = $i; $used[$i] = $true; break } }
    }
    $rest = @(for ($i = 0; $i -lt $Lines.Length; $i++) { if (-not $used[$i]) { $i } })
    $sure = ($byCount -eq @($Lines | Where-Object { $_ -gt 0 }).Count)
    return @{ order = @($order); rest = $rest; sure = $sure }
}

function Invoke-Step {
    param($In, $Ctx)
    $pairs = @(@($In['pairs']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $shotsIn = @($In['shots'])
    $g = VerifyDfRanges-Slots -Ranges $In['ranges']
    $slots = @($g['slots'])
    $lines = [int[]]@($pairs | ForEach-Object { [int]$_['lines'] })
    if ($null -eq $lines) { $lines = [int[]]@() }
    $a = VerifyDfRanges-Assign -Slots $slots -Lines $lines -Offset ([int]$In['offset']) -MaxExtra ([int]$In['maxExtra'])
    if (-not [bool]$a['sure']) {
        # not sure which HOST pictures go with which file: the pairs keep
        # track's order and take the readable sets one by one (as before 10/26)
        $old = VerifyDfRanges-Sets -Ranges $In['ranges']
        $sets = @($old['sets'])
        $slots = @($sets | ForEach-Object { @{ kind = 'set'; set = $_ } })
        $ord = New-Object int[] $slots.Count
        for ($k = 0; $k -lt $slots.Count; $k++) { $ord[$k] = $(if ($k -lt $pairs.Count) { $k } else { -1 }) }
        $a = @{ order = @($ord); rest = @(for ($i = $slots.Count; $i -lt $pairs.Count; $i++) { $i }); sure = $false }
        $g = @{ slots = $slots; stray = [int]$old['bad'] }
    }
    $plans = New-Object System.Collections.ArrayList
    $outPairs = New-Object System.Collections.ArrayList
    $outShots = New-Object System.Collections.ArrayList
    $outOrder = New-Object System.Collections.ArrayList
    $why = New-Object System.Collections.ArrayList
    $alerts = New-Object System.Collections.ArrayList
    $aligned = 0
    $any = (@(@($In['ranges']) | Where-Object { $_ -is [System.Collections.IDictionary] }).Count -gt 0)
    $blankLeft = 0; $setLeft = 0
    $take = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $slots.Count; $k++) {
        $i = [int]$a['order'][$k]
        if ($i -lt 0) { if ([string]$slots[$k]['kind'] -eq 'blank') { $blankLeft++ } else { $setLeft++ }; continue }
        [void]$take.Add(@{ i = $i; slot = $slots[$k] })
    }
    foreach ($i in @($a['rest'])) { [void]$take.Add(@{ i = [int]$i; slot = $null }) }
    foreach ($t in $take) {
        $i = [int]$t['i']; $sl = $t['slot']
        $n = [int]$pairs[$i]['lines']
        if ($null -ne $sl -and [string]$sl['kind'] -eq 'blank') {
            $p = if ($n -eq 0) { @{ mode = 'empty'; first = 0; end = 0; reason = 'an empty file; its HOST picture shows no line numbers' } }
                 else { @{ mode = 'default'; first = 0; end = 0; reason = ('its HOST picture could not be read (no line numbers) for a ' + $n + '-line file') } }
        } else {
            $set = if ($null -ne $sl) { $sl['set'] } else { $null }
            $p = VerifyDfRanges-Plan -Lines $n -Set $set -Offset ([int]$In['offset']) -MaxLines ([int]$In['maxLines']) -MaxExtra ([int]$In['maxExtra'])
        }
        [void]$plans.Add($p); [void]$outPairs.Add($pairs[$i]); [void]$outOrder.Add($i)
        if ($shotsIn.Count -gt 0) { [void]$outShots.Add($(if ($i -lt $shotsIn.Count) { $shotsIn[$i] } else { $null })) }
        if ($p['mode'] -eq 'align') { $aligned++ }
        $name = [string]$pairs[$i]['rightName']
        $shown = if ($name) { $name } else { 'pair ' + ($i + 1) }
        [void]$why.Add(('{0}: {1} -- {2}' -f $shown, $p['mode'], $p['reason']))
        if ($p['mode'] -eq 'default' -and $any) { [void]$alerts.Add(('DF pictures of {0} NOT lined up with the HOST pictures: {1} -- check them by hand' -f $shown, $p['reason'])) }
    }
    $bad = [int]$g['stray'] + $blankLeft
    if ($any -and $bad -gt 0) { [void]$alerts.Add(('{0} HOST picture(s) on the send-data sheet could not be read (no line numbers)' -f $bad)) }
    if ($setLeft -gt 0) { [void]$alerts.Add(('{0} HOST picture set(s) but {1} DF pair(s)' -f ($slots.Count - $blankLeft), $pairs.Count)) }
    $moved = $false
    for ($k = 0; $k -lt $outOrder.Count; $k++) { if ([int]$outOrder[$k] -ne $k) { $moved = $true } }
    if ($moved) { $Ctx.Log.Info(('DF pairs put in the send-data sheet''s order: ' + (@($outPairs | ForEach-Object { [string]$_['rightName'] }) -join ', '))) }
    foreach ($w in $why) { $Ctx.Log.Info($w) }
    # a pair track left no shot for keeps a null there, so the index still
    # matches the pair (launch_capture reuses nothing for it)
    return @{ ok = $true; pairs = $outPairs.ToArray(); shots = $outShots.ToArray(); order = $outOrder.ToArray(); moved = $moved; plans = $plans.ToArray(); aligned = $aligned; alert = ($alerts -join ' / '); reason = ($why -join ' / ') }
}
