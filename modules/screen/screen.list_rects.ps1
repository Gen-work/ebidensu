# modules/screen/screen.list_rects.ps1
# Where are the target entries of a list in a screenshot of that list
# scrolled to its END? The page text gives the entries in order; the
# picture gives the text bands (screen.find_ink_rows). Counting both from
# the bottom pairs them up: the last band is the last entry. One rect per
# run of consecutive targets, the way the operator boxes three files in a
# row with one rectangle. Without bands it falls back to a fixed pitch
# from a fixed last-row position.
#
# A past day re-run (K-46): the targets are older entries, above the ~30
# rows an end-scrolled picture shows. With allowAbove that is not a
# failure but 'above' = true plus 'anchor' = the topmost target, so the
# workflow can Ctrl+F it into view and call this step again with
# anchorName + anchorBands (the orange active-match band): the bands are
# then counted from the hit instead of from the bottom.

. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')

$Manifest = @{
  id         = 'screen.list_rects'
  group      = 'screen'
  summary    = 'Locate target list entries in an end-scrolled screenshot; one rect per run'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    names          = @{ type='list'; required=$true; desc='every entry of the list in page order (parsed page text)' }
    targets        = @{ type='list'; required=$true; desc='the entries to box' }
    bands          = @{ type='list'; default=@(); desc='screen.find_ink_rows bands of the entry column' }
    x              = @{ type='int';  required=$true; desc='box left edge, px' }
    width          = @{ type='int';  required=$true }
    height         = @{ type='int';  default=25; desc='box height for one row, px' }
    lastRowCenterY = @{ type='int';  default=0; desc='fallback when bands is empty: centre of the last row, px' }
    rowPitch       = @{ type='int';  default=20; desc='fallback row pitch, px' }
    allowAbove     = @{ type='bool'; default=$false; desc='bands from the bottom: targets above the picture -> ok with above = true instead of not_found' }
    anchorName     = @{ type='string'; default=''; desc='the entry the picture is anchored on (the Ctrl+F hit); with anchorBands' }
    anchorBands    = @{ type='list'; default=@(); desc='screen.find_ink_rows bands of the hit highlight (ink orange); the first is the anchor row' }
  }
  outputs    = @{
    rects   = @{ type='list'; desc='@{ x; y; w; h; names } in image px' }
    missing = @{ type='list'; desc='targets not among names, or above the top of the picture' }
    source  = @{ type='string'; desc='bands | anchor | pitch' }
    above   = @{ type='bool'; desc='bands from the bottom: some target is above the top of the picture' }
    anchor  = @{ type='string'; desc='the topmost target in page order (what to Ctrl+F when above)' }
  }
  failures   = @(
    @{ id = 'not_found'; transient = $false }
  )
  example    = @{ use = 'screen.list_rects'; with = @{ names = '{{steps.rec.out.names}}'; targets = '{{steps.mine.out.plucked}}'; bands = '{{steps.ink.out.bands}}'; x = 445; width = 892 } }
  notes      = 'not_found when no target can be placed at all (with allowAbove: unless the reason is targets above the picture). Bands are matched from the bottom, or from the anchor band when anchorName is set, only as far as there are bands; an entry outside the picture is reported missing, not boxed at a guessed place. Anchor mode: the anchor row keeps its place even when its own link text lost the blue under the highlight.'
}

function ScreenListRects-AnchorBand {
    # PURE. Index of the band the hit sits on: the band whose centre is
    # within half a pitch of the hit, else the hit itself inserted as a
    # band (the highlighted name may have lost its blue ink).
    # -> @{ bands; index }
    param([object[]]$Bands, [double]$Y, [double]$Pitch)
    $tol = [Math]::Max(4.0, $Pitch / 2.0)
    $best = -1; $bd = [double]::MaxValue
    for ($i = 0; $i -lt $Bands.Count; $i++) {
        $d = [Math]::Abs([double]$Bands[$i]['center'] - $Y)
        if ($d -lt $bd) { $bd = $d; $best = $i }
    }
    if ($best -ge 0 -and $bd -le $tol) { return @{ bands = $Bands; index = $best } }
    $at = 0
    while ($at -lt $Bands.Count -and [double]$Bands[$at]['center'] -lt $Y) { $at++ }
    $hit = @{ top = [int][Math]::Floor($Y); bottom = [int][Math]::Ceiling($Y); height = 1; center = $Y }
    $out = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $Bands.Count; $i++) {
        if ($i -eq $at) { [void]$out.Add($hit) }
        [void]$out.Add($Bands[$i])
    }
    if ($at -ge $Bands.Count) { [void]$out.Add($hit) }
    return @{ bands = $out.ToArray(); index = $at }
}

function Invoke-Step {
    param($In, $Ctx)
    $names = @(@($In['names']) | ForEach-Object { [string]$_ })
    $targets = @(@($In['targets']) | ForEach-Object { [string]$_ })
    $bands = @(@($In['bands']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $hits = @(@($In['anchorBands']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $anchorName = [string]$In['anchorName']
    $h = [int]$In['height']
    $rects = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList
    $source = 'pitch'
    $above = $false
    $why = @()
    $idx = @()
    foreach ($t in $targets) { $i = [array]::IndexOf($names, $t); if ($i -ge 0) { $idx += $i } }
    $idx = @($idx | Sort-Object -Unique)
    $anchor = if ($idx.Count -gt 0) { $names[[int]$idx[0]] } else { '' }
    if ($anchorName -ne '') {
        $ai = [array]::IndexOf($names, $anchorName)
        if ($ai -lt 0) { return @{ ok = $false; failure = 'not_found'; message = ('the anchor ' + $anchorName + ' is not in the page text'); rects = @(); missing = $targets; source = 'anchor'; above = $false; anchor = $anchor } }
        if ($hits.Count -eq 0) {
            if ($Ctx['DryRun']) { return @{ ok = $true; rects = @(); missing = $targets; source = 'anchor'; above = $false; anchor = $anchor } }
            return @{ ok = $false; failure = 'not_found'; message = ('no Ctrl+F highlight in the picture (searched ' + $anchorName + '; the find bar found nothing, or the highlight colour is not the orange the ink scan looks for)'); rects = @(); missing = $targets; source = 'anchor'; above = $false; anchor = $anchor }
        }
        $ab = ScreenListRects-AnchorBand -Bands $bands -Y ([double]$hits[0]['center']) -Pitch ([double]$In['rowPitch'])
        $bands = @($ab['bands'])
        $source = 'anchor'
        $off = [int]$ab['index'] - $ai
    } elseif ($bands.Count -gt 0) {
        $source = 'bands'
        $off = $bands.Count - $names.Count
    }
    if ($source -ne 'pitch') {
        foreach ($t in $targets) { if ([array]::IndexOf($names, $t) -lt 0) { [void]$missing.Add($t); $why += ($t + ' is not in the page text (' + $names.Count + ' entries; copied before the list finished loading?)') } }
        $n = $names.Count; $m = $bands.Count
        $k = 0
        while ($k -lt $idx.Count) {
            $a = [int]$idx[$k]; $b = $a
            while (($k + 1) -lt $idx.Count -and [int]$idx[$k + 1] -eq ($b + 1)) { $k++; $b = [int]$idx[$k] }
            $ba = $a + $off; $bb = $b + $off
            if ($ba -lt 0 -or $bb -ge $m) {
                foreach ($j in $a..$b) { [void]$missing.Add($names[$j]) }
                if ($source -eq 'bands') { $above = $true; $why += ('{0} is row {1} from the bottom, the picture has {2} text row(s)' -f $names[$a], ($n - $a), $m) }
                else { $why += ('{0} is {1} row(s) from the Ctrl+F hit {2}, outside the picture ({3} text row(s))' -f $names[$a], ($a - $ai), $anchorName, $m) }
                $k++; continue
            }
            $ca = [double]$bands[$ba]['center']; $cb = [double]$bands[$bb]['center']
            $top = [int][Math]::Round($ca - $h / 2.0, [System.MidpointRounding]::AwayFromZero); $bottom = [int][Math]::Round($cb + $h / 2.0, [System.MidpointRounding]::AwayFromZero)
            [void]$rects.Add(@{ x = [int]$In['x']; y = $top; w = [int]$In['width']; h = ($bottom - $top); names = @($names[$a..$b]) })
            $k++
        }
    } else {
        $g = Get-EbiListEntryRects -Names $names -Targets $targets -LastRowCenterY ([double]$In['lastRowCenterY']) -RowPitch ([double]$In['rowPitch']) -X ([int]$In['x']) -Width ([int]$In['width']) -Height $h
        foreach ($r in @($g['rects'])) { [void]$rects.Add(@{ x = [int]$r['x']; y = [int]$r['y']; w = [int]$r['width']; h = [int]$r['height']; names = @($r['names']) }) }
        foreach ($x in @($g['missing'])) { [void]$missing.Add($x) }
    }
    if ($above -and [bool]$In['allowAbove']) {
        # the caller Ctrl+Fs 'anchor' and asks again; nothing to warn about yet
        return @{ ok = $true; rects = $rects.ToArray(); missing = $missing.ToArray(); source = $source; above = $true; anchor = $anchor }
    }
    if ($rects.Count -eq 0) {
        if ($Ctx['DryRun']) { return @{ ok = $true; rects = @(); missing = $missing.ToArray(); source = $source; above = $above; anchor = $anchor } }
        return @{ ok = $false; failure = 'not_found'; message = ('none of ' + ($targets -join ', ') + ' could be placed' + $(if ($why.Count -gt 0) { ': ' + ($why -join '; ') } else { '' })); rects = @(); missing = $missing.ToArray(); source = $source; above = $above; anchor = $anchor }
    }
    $out = @{ ok = $true; rects = $rects.ToArray(); missing = $missing.ToArray(); source = $source; above = $above; anchor = $anchor }
    $warn = New-Object System.Collections.ArrayList
    if ($missing.Count -gt 0) { [void]$warn.Add(@{ code = 'target_missing'; message = ('not placed: ' + ($missing -join ', ')) }) }
    if ($source -eq 'bands' -and $bands.Count -ge 2) {
        # the bottom band should be a list row: one pitch below the one above it
        $gap = [double]$bands[$bands.Count - 1]['center'] - [double]$bands[$bands.Count - 2]['center']
        $p = [double]$In['rowPitch']
        if ($p -gt 0 -and [Math]::Abs($gap - $p) -gt ($p * 0.35)) { [void]$warn.Add(@{ code = 'pitch_mismatch'; message = ('the last two ink bands are ' + $gap + ' px apart (row pitch ' + $p + '): something below the list may be counted as a row -- check the box') }) }
    }
    if ($warn.Count -gt 0) { $out['warnings'] = $warn.ToArray() }
    return $out
}
