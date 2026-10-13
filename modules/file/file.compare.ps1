# modules/file/file.compare.ps1
# Compare pairs of text files line by line, blind to CRLF vs LF (the
# question a diff tool's "same content" answers, asked without the tool).
# Each file is decoded in its own single encoding (UTF-8 if it is valid
# UTF-8 throughout, else Shift_JIS) and the TEXT is compared: the GIFT side
# writes Shift_JIS, the GFIX side UTF-8, and DF calls them the same content.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.compare'
  group      = 'file'
  summary    = 'Compare pairs of text files line by line; ok when every pair is identical'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    pairs     = @{ type='list';   required=$true; desc='@{ left; right } paths (verify.pair_files output)' }
    leftField = @{ type='string'; default='left' }
    rightField = @{ type='string'; default='right' }
    repairByContent = @{ type='bool'; default=$false; desc='when pairs differ, re-pair the differing files by content: a left file goes with the ONE differing right file identical to it. Only when every differing left finds exactly one such right and no right is claimed twice; anything else keeps the pairs and the ng. Several files of one job arrive in an order the file names need not follow' }
  }
  outputs    = @{
    code      = @{ type='string'; desc='ok (all identical) | ng (a pair differs) | unknown (no pairs, or a file missing)' }
    results   = @{ type='list';   desc='@{ left; right; identical; firstDiff; leftLines; rightLines; leftEncoding; rightEncoding } per pair' }
    identical = @{ type='int';    desc='pairs that are identical' }
    reason    = @{ type='string' }
    pairs     = @{ type='list';   desc='the pairs as compared: the input pairs, re-paired when repairByContent moved a right file (every other field of a pair is kept, so the DF step can take this list)' }
    repaired  = @{ type='int';    desc='pairs whose right file was changed by repairByContent (0 = the input pairing)' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'file.compare'; with = @{ pairs = '{{steps.pair.out.pairs}}' } }
  notes      = 'A difference is a verdict (ng), not a failure: the workflow decides (human.gate). firstDiff is 1-based. repairByContent never invents a match: a re-pairing is reported (repaired, a repaired_by_content warning, the reason) and only happens when the content proves it.'
}

function FileCompare-Repair {
    <#
      PURE. Given the per-pair texts (neutral, one entry per pair) and which
      pairs differ, find the one re-pairing of the DIFFERING pairs that
      makes every one identical: left i takes the differing right j whose
      text equals its own. Returns @{ ok; map = @{ i -> j } } -- ok only
      when every differing left has exactly one such j and no j is taken
      twice (two identical rights would be a guess; two candidates for one
      left too). Pairs already identical are never touched.
    #>
    param([string[]]$LeftTexts, [string[]]$RightTexts, [int[]]$Differing)
    $map = @{}
    $taken = @{}
    foreach ($i in @($Differing)) {
        $hits = @(foreach ($j in @($Differing)) { if ((Compare-EbiTextLines -Left $LeftTexts[$i] -Right $RightTexts[$j])['identical']) { $j } })
        if ($hits.Count -ne 1) { return @{ ok = $false; map = @{}; why = $(if ($hits.Count -eq 0) { 'no differing right file has the content of left ' + $i } else { 'left ' + $i + ' matches ' + $hits.Count + ' right files' }) } }
        $j = [int]$hits[0]
        if ($taken.Contains($j)) { return @{ ok = $false; map = @{}; why = ('right ' + $j + ' matches two left files') } }
        $taken[$j] = $true
        $map[$i] = $j
    }
    return @{ ok = $true; map = $map; why = '' }
}

function FileCompare-Rewire {
    <#
      PURE. Apply a repair map to the pair list: pair i keeps every field
      but takes the right-side fields (RightField, rightName and anything
      starting with 'right') of pair map[i], and notes the right file it
      was listed with (repairedFrom, the name). Unmapped pairs are copied
      as they are. Returns the new list.
    #>
    param([object[]]$Pairs, [hashtable]$Map, [string]$RightField)
    $out = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Pairs.Count; $i++) {
        $copy = @{}
        foreach ($k in $Pairs[$i].Keys) { $copy[[string]$k] = $Pairs[$i][$k] }
        if ($Map.Contains($i)) {
            $src = $Pairs[[int]$Map[$i]]
            foreach ($k in @($src.Keys)) {
                $ks = [string]$k
                if ($ks -eq $RightField -or $ks.StartsWith('right', [System.StringComparison]::OrdinalIgnoreCase)) { $copy[$ks] = $src[$k] }
            }
            $copy['repairedFrom'] = $(if ($Pairs[$i].Contains('rightName') -and -not [string]::IsNullOrWhiteSpace([string]$Pairs[$i]['rightName'])) { [string]$Pairs[$i]['rightName'] } else { Split-Path ([string]$Pairs[$i][$RightField]) -Leaf })
        }
        [void]$out.Add($copy)
    }
    return $out.ToArray()
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $lf = [string]$In['leftField']; $rf = [string]$In['rightField']
    $pairs = @(@($In['pairs']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    if ($pairs.Count -eq 0) { return @{ ok = $true; code = 'unknown'; results = @(); identical = 0; reason = 'no pairs to compare'; pairs = @(); repaired = 0 } }
    $repaired = 0
    $repairNote = ''
    $warn = New-Object System.Collections.ArrayList
    $pass = 0
    while ($true) {
        $pass++
        $res = New-Object System.Collections.ArrayList
        $same = 0; $code = 'ok'; $why = New-Object System.Collections.ArrayList
        $lt = New-Object System.Collections.ArrayList; $rt = New-Object System.Collections.ArrayList
        $diffIdx = New-Object System.Collections.ArrayList
        $k0 = -1
        foreach ($p in $pairs) {
            $k0++
            [void]$lt.Add(''); [void]$rt.Add('')
            $l = Resolve-EbiWorkPath -PathValue ([string]$p[$lf]) -WorkDir $work
            $r = Resolve-EbiWorkPath -PathValue ([string]$p[$rf]) -WorkDir $work
            if (-not (Test-Path -LiteralPath $l -PathType Leaf) -or -not (Test-Path -LiteralPath $r -PathType Leaf)) {
                if ($Ctx['DryRun']) { [void]$res.Add(@{ left = $l; right = $r; identical = $false; firstDiff = 0; leftLines = 0; rightLines = 0 }); continue }
                $code = 'unknown'; [void]$why.Add(('missing: ' + $(if (-not (Test-Path -LiteralPath $l)) { $l } else { $r })))
                [void]$res.Add(@{ left = $l; right = $r; identical = $false; firstDiff = 0; leftLines = 0; rightLines = 0 }); continue
            }
            # each file in its own single encoding (GIFT Shift_JIS, GFIX UTF-8):
            # the TEXT is compared, as the diff tool's "same content" does
            $da = ConvertFrom-EbiTextBytes -Bytes ([System.IO.File]::ReadAllBytes($l))
            $db = ConvertFrom-EbiTextBytes -Bytes ([System.IO.File]::ReadAllBytes($r))
            $lt[$k0] = ConvertTo-EbiJisNeutral -Text $da['text']; $rt[$k0] = ConvertTo-EbiJisNeutral -Text $db['text']
            $c = Compare-EbiTextLines -Left $lt[$k0] -Right $rt[$k0]
            if (-not [bool]$c['identical']) { [void]$diffIdx.Add($k0) }
            $encNote = if ($da['encoding'] -ne $db['encoding']) { (' [' + $da['encoding'] + ' vs ' + $db['encoding'] + ']') } else { '' }
            if (-not [bool]$c['identical']) {
            # say exactly where: the line, the column, the two characters
            $la = @(Get-EbiTextLines -Text $da['text']); $lb = @(Get-EbiTextLines -Text $db['text'])
            $k = [int]$c['firstDiff'] - 1
            $fd = Get-EbiFirstCharDiff -Left $(if ($k -lt $la.Count) { [string]$la[$k] } else { '' }) -Right $(if ($k -lt $lb.Count) { [string]$lb[$k] } else { '' })
            if ([int]$fd['col'] -gt 0) { $encNote = $encNote + (' col ' + $fd['col'] + ': ' + $(if ($fd['left'] -ne '') { $fd['left'] } else { 'end' }) + ' vs ' + $(if ($fd['right'] -ne '') { $fd['right'] } else { 'end' })) }
            }
            if ([bool]$c['identical']) { $same++; if ($encNote -ne '') { [void]$why.Add(('{0} vs {1}: same text{2}' -f (Split-Path $l -Leaf), (Split-Path $r -Leaf), $encNote)) } } else { if ($code -eq 'ok') { $code = 'ng' }; [void]$why.Add(('{0} vs {1}: first difference at line {2} ({3} vs {4} lines){5}' -f (Split-Path $l -Leaf), (Split-Path $r -Leaf), $c['firstDiff'], $c['leftLines'], $c['rightLines'], $encNote)) }
            [void]$res.Add(@{ left = $l; right = $r; identical = [bool]$c['identical']; firstDiff = [int]$c['firstDiff']; leftLines = [int]$c['leftLines']; rightLines = [int]$c['rightLines']; leftEncoding = $da['encoding']; rightEncoding = $db['encoding'] })
        }
        if ($Ctx['DryRun']) { return @{ ok = $true; code = 'ok'; results = $res.ToArray(); identical = 0; reason = 'dry run'; pairs = $pairs; repaired = 0 } }
        # Several files, every pair differing, but each left file identical
        # to ANOTHER right file: the arrival order the pairing followed
        # disagreed with the real order. Re-pair by content, once, and
        # compare again -- the proof is the second pass, not this guess.
        if ($pass -eq 1 -and [bool]$In['repairByContent'] -and $code -eq 'ng' -and $diffIdx.Count -ge 2) {
            $fix = FileCompare-Repair -LeftTexts ([string[]]$lt.ToArray()) -RightTexts ([string[]]$rt.ToArray()) -Differing ([int[]]$diffIdx.ToArray())
            if ($fix['ok']) {
                $moves = @(foreach ($i in @($fix['map'].Keys | Sort-Object)) { (Split-Path ([string]$pairs[$i][$lf]) -Leaf) + ' <- ' + (Split-Path ([string]$pairs[[int]$fix['map'][$i]][$rf]) -Leaf) + ' (was ' + (Split-Path ([string]$pairs[$i][$rf]) -Leaf) + ')' })
                $pairs = @(FileCompare-Rewire -Pairs $pairs -Map $fix['map'] -RightField $rf)
                $repaired = $fix['map'].Count
                $repairNote = ('re-paired ' + $repaired + ' pair(s) by content, the arrival order disagreed: ' + ($moves -join '; '))
                [void]$warn.Add(@{ code = 'repaired_by_content'; message = $repairNote; data = @{ moves = $moves } })
                $Ctx.Log.Info($repairNote)
                continue
            }
            $Ctx.Log.Info(('no re-pairing by content: ' + $fix['why']))
        }
        break
    }
    $reason = if ($code -eq 'ok') { ('' + $same + ' pair(s) identical' + $(if ($why.Count -gt 0) { ' -- ' + ($why -join '; ') } else { '' })) } elseif ($why.Count -gt 0) { $why -join '; ' } else { ('' + $same + ' pair(s) identical') }
    if ($repairNote -ne '') { $reason = $repairNote + ' -- ' + $reason }
    $out = @{ ok = $true; code = $code; results = $res.ToArray(); identical = $same; reason = $reason; pairs = $pairs; repaired = $repaired }
    if ($warn.Count -gt 0) { $out['warnings'] = $warn.ToArray() }
    return $out
}
