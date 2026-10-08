# modules/screen/screen.list_rects.ps1
# Where are the target entries of a list in a screenshot of that list
# scrolled to its END? The page text gives the entries in order; the
# picture gives the text bands (screen.find_ink_rows). Counting both from
# the bottom pairs them up: the last band is the last entry. One rect per
# run of consecutive targets, the way the operator boxes three files in a
# row with one rectangle. Without bands it falls back to a fixed pitch
# from a fixed last-row position.

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
  }
  outputs    = @{
    rects   = @{ type='list'; desc='@{ x; y; w; h; names } in image px' }
    missing = @{ type='list'; desc='targets not among names, or above the top of the picture' }
    source  = @{ type='string'; desc='bands | pitch' }
  }
  failures   = @(
    @{ id = 'not_found'; transient = $false }
  )
  example    = @{ use = 'screen.list_rects'; with = @{ names = '{{steps.rec.out.names}}'; targets = '{{steps.mine.out.plucked}}'; bands = '{{steps.ink.out.bands}}'; x = 445; width = 892 } }
  notes      = 'not_found when no target can be placed at all. Bands are matched from the bottom only as far as there are bands; an entry scrolled off the top is reported missing, not boxed at a guessed place.'
}

function Invoke-Step {
    param($In, $Ctx)
    $names = @(@($In['names']) | ForEach-Object { [string]$_ })
    $targets = @(@($In['targets']) | ForEach-Object { [string]$_ })
    $bands = @(@($In['bands']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $h = [int]$In['height']
    $rects = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList
    $source = 'pitch'
    if ($bands.Count -gt 0) {
        $source = 'bands'
        $idx = @()
        foreach ($t in $targets) { $i = [array]::IndexOf($names, $t); if ($i -lt 0) { [void]$missing.Add($t) } else { $idx += $i } }
        $idx = @($idx | Sort-Object -Unique)
        $n = $names.Count; $m = $bands.Count
        $k = 0
        while ($k -lt $idx.Count) {
            $a = [int]$idx[$k]; $b = $a
            while (($k + 1) -lt $idx.Count -and [int]$idx[$k + 1] -eq ($b + 1)) { $k++; $b = [int]$idx[$k] }
            $ba = $m - ($n - $a); $bb = $m - ($n - $b)
            if ($ba -lt 0) { foreach ($j in $a..$b) { [void]$missing.Add($names[$j]) }; $k++; continue }
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
    if ($rects.Count -eq 0) {
        if ($Ctx['DryRun']) { return @{ ok = $true; rects = @(); missing = $missing.ToArray(); source = $source } }
        return @{ ok = $false; failure = 'not_found'; message = ('none of ' + ($targets -join ', ') + ' could be placed'); rects = @(); missing = $missing.ToArray(); source = $source }
    }
    $out = @{ ok = $true; rects = $rects.ToArray(); missing = $missing.ToArray(); source = $source }
    if ($missing.Count -gt 0) { $out['warnings'] = @(@{ code = 'target_missing'; message = ('not placed: ' + ($missing -join ', ')) }) }
    return $out
}
