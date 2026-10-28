# modules/screen/screen.stitch_scroll.ps1
# A horizontally long snap of a text file (10/28): open the file in a text
# viewer, capture its window, scroll one page right, capture again, ...
# until the widest line's end is on a capture (file.line_ends needCols) or
# the view cannot scroll further; then merge the captures into one long
# PNG (outside Excel). The captures themselves are kept too, for
# excel.place_tiles (the same pages put on the sheet and grouped, inside
# Excel) -- the evidence books did both by hand.
#
# Where each capture starts (kernel/Stitch.ps1 Resolve-EbiTileShift): the
# window's horizontal scroll bar position when it can be read (exact, in
# columns); else the pixels -- the same picture twice = the view did not
# move (the end); a long stretch repeated = a scroll cut short at the end
# of the range; otherwise one page (a page scroll moves exactly the columns
# the text area shows: measured on the J4 GFIX data pictures, 265 a page,
# joined without a gap or a repeat).
#
# How to scroll (advance, tried in order until the view moves):
#   message  WM_HSCROLL page-right to the scroll bar's window (no mouse)
#   click    a click on the scroll bar's trough, right of its thumb
#   keys     keysAdvance (SendKeys syntax)

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')    # Save-EbiScreenRegionPng, Get-EbiPngSize, Get-EbiGrayRegion
. (Join-Path $PSScriptRoot '..\..\kernel\Stitch.ps1')

$Manifest = @{
  id         = 'screen.stitch_scroll'
  group      = 'screen'
  summary    = 'Open a text file in a viewer, capture it page by page to the right until the line end, merge into one long PNG'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    exe         = @{ type='path';   required=$true; desc='the text viewer' }
    file        = @{ type='path';   required=$true; desc='the text file (also what the window title must show)' }
    args        = @{ type='list';   default=@(); desc='arguments before the file (e.g. a read-only switch)' }
    windowTitle = @{ type='string'; default=''; desc='the viewer window title contains this; empty = the file name' }
    needCols    = @{ type='int';    default=0; desc='columns the captures must reach (file.line_ends needCols: the widest line + its line end mark); 0 = until the view cannot scroll further' }
    saveDir     = @{ type='path';   required=$true }
    name        = @{ type='string'; default=''; desc='file name stem of the captures (<name>__01.png .. and <name>__wide.png); empty = the file name' }
    stitch      = @{ type='bool';   default=$true; desc='also merge the captures into <name>__wide.png' }
    frame       = @{ type='string'; default='each'; enum=@('each', 'first'); desc='the merged PNG: each = every capture whole (window frame and all) side by side; first = one frame (the first capture''s title / menu / toolbar), the text bands joined column to column' }
    x           = @{ type='int';    default=40 }
    y           = @{ type='int';    default=40 }
    width       = @{ type='int';    default=0; desc='visible window width; 0 = leave as the viewer opens it (the geometry defaults were measured at 1906)' }
    height      = @{ type='int';    default=0; desc='visible window height; 0 = fitLines decides (else leave as the viewer opens it)' }
    fitLines    = @{ type='int';    default=0; desc='height = 0: size the window to show this many lines and the [EOF] row, no more (file.line_ends lines); capped by maxHeight' }
    rowPx       = @{ type='int';    default=14; desc='px a text row (J4 GFIX data pictures: 14)' }
    chromePx    = @{ type='int';    default=110; desc='px of the window that is not text rows: title, menu, toolbar, ruler (92) + the horizontal scroll bar and border (18); 166 = 3 lines + [EOF]' }
    maxHeight   = @{ type='int';    default=900 }
    keysFirst   = @{ type='string'; default='^{HOME}'; desc='keys before the first capture (the top-left of the file)' }
    advance     = @{ type='list';   default=@('message', 'click'); desc='how to scroll a page right, tried in order until the view moves: message | click | keys' }
    keysAdvance = @{ type='string'; default=''; desc='advance = keys: the keys of one step right' }
    clickRight  = @{ type='int';    default=40; desc='advance = click when the scroll bar cannot be found: px from the capture''s right edge' }
    clickBottom = @{ type='int';    default=8; desc='... and from its bottom edge' }
    textLeft    = @{ type='int';    default=30; desc='px from the capture''s left edge to column 0 (past the line-number strip)' }
    textRight   = @{ type='int';    default=19; desc='px from the right edge back to the text area''s end (the vertical scroll bar)' }
    bandTop     = @{ type='int';    default=76; desc='first row stitched (the ruler); the rows above come from the first capture only' }
    bandBottom  = @{ type='int';    default=-18; desc='row after the stitched band; <= 0 from the bottom (the horizontal scroll bar is left out)' }
    charPx      = @{ type='int';    default=7; desc='px a half-width column' }
    maxTiles    = @{ type='int';    default=40 }
    close       = @{ type='bool';   default=$true; desc='close the viewer after (only a window this step opened)' }
    windowWaitSec = @{ type='int';  default=15 }
    settleMs    = @{ type='int';    default=700; desc='wait after the window appears / after each scroll' }
  }
  outputs    = @{
    tiles       = @{ type='list';   desc='the captures, left to right' }
    starts      = @{ type='list';   desc='the first column of each capture' }
    how         = @{ type='list';   desc='per capture after the first: scroll | match | page (how its start was known)' }
    stitched    = @{ type='path';   desc='the merged PNG; empty when stitch = false' }
    width       = @{ type='int';    desc='its width, px' }
    height      = @{ type='int' }
    layout      = @{ type='map';    desc='@{ textLeft; textRight; bandTop; bandBottom; charPx; pageCols; tileW; tileH; starts } for excel.place_tiles' }
    coveredCols = @{ type='int';    desc='columns 0 .. coveredCols-1 are on a capture' }
    needCols    = @{ type='int' }
    reachedEnd  = @{ type='bool';   desc='needCols > 0: the line end is on the last capture; else the view could not scroll further' }
    atScrollEnd = @{ type='bool' }
  }
  failures   = @(
    @{ id = 'file_not_found';  transient = $false }
    @{ id = 'no_window';       transient = $true  }
    @{ id = 'foreground_lost'; transient = $true  }
    @{ id = 'save_failed';     transient = $true  }
    @{ id = 'geometry';        transient = $false }
    @{ id = 'not_at_line_end'; transient = $false }
  )
  example    = @{ use = 'screen.stitch_scroll'; with = @{ exe = '{{vars.viewer}}'; file = '{{vars.dataFile}}'; args = @('-R'); needCols = '{{steps.eol.out.needCols}}'; saveDir = 'capture/wide' } }
  notes      = 'not_at_line_end: the captures and the merged PNG are kept and named in the message (maxTiles reached, or the view stopped short of needCols: the viewer wraps lines, or its line width setting is narrower than the file). The geometry defaults are the J4 GFIX data pictures (a 1906 px wide window); a capture must come out pixel for pixel, so run at 100 % display scaling.'
}

function ScreenStitchScroll-Plan {
    # PURE. Captures a needCols of columns takes at pageCols a page (1 when unknown).
    param([int]$NeedCols, [int]$PageCols, [int]$MaxTiles)
    if ($NeedCols -le 0 -or $PageCols -le 0) { return 1 }
    return [Math]::Min([Math]::Max(1, $MaxTiles), [int][Math]::Ceiling($NeedCols / [double]$PageCols))
}

function ScreenStitchScroll-FitHeight {
    # PURE. Window height that shows Lines text rows + the [EOF] row: 0 when
    # Lines <= 0 (leave the height alone); at most MaxHeight.
    param([int]$Lines, [int]$RowPx, [int]$ChromePx, [int]$MaxHeight)
    if ($Lines -le 0 -or $RowPx -le 0) { return 0 }
    $h = $ChromePx + $RowPx * ($Lines + 1)
    if ($MaxHeight -gt 0 -and $h -gt $MaxHeight) { $h = $MaxHeight }
    return $h
}

function ScreenStitchScroll-PickWindow {
    <#
      PURE. Which top window shows the file? -> @{ handle; ours }.
      A window that was not there before the start and whose title has
      Title > one that was already there with Title (the viewer reused its
      window for a file already open: captured, but not ours to close).
    #>
    param($Windows, $Before, [string]$Title)
    $old = @(@($Before) | ForEach-Object { [string]$_ })
    $cand = @(@($Windows) | Where-Object { ([string]$_['title']).IndexOf($Title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })
    foreach ($w in $cand) { if ($old -notcontains [string]$w['handle']) { return @{ handle = $w['handle']; ours = $true } } }
    if ($cand.Count -gt 0) { return @{ handle = $cand[0]['handle']; ours = $false } }
    return @{ handle = [IntPtr]::Zero; ours = $false }
}

function ScreenStitchScroll-TroughPoint {
    <#
      PURE. Screen point of a click on the horizontal scroll bar's trough,
      just left of its right arrow (a square as tall as the bar): from the
      bar's rectangle when known, else Right / Bottom px off the capture's
      bottom-right corner. -> @{ x; y }
    #>
    param($BarRect, $Visible, [int]$Right = 40, [int]$Bottom = 8)
    if ($BarRect -is [System.Collections.IDictionary] -and [int]$BarRect['W'] -gt 0 -and [int]$BarRect['H'] -gt 0) {
        $h = [int]$BarRect['H']
        return @{ x = ([int]$BarRect['X'] + [int]$BarRect['W'] - $h - 4); y = ([int]$BarRect['Y'] + [int][Math]::Floor($h / 2)) }
    }
    return @{ x = ([int]$Visible['X'] + [int]$Visible['W'] - $Right); y = ([int]$Visible['Y'] + [int]$Visible['H'] - $Bottom) }
}

function ScreenStitchScroll-Shot {
    # Capture the visible window to Dest. @{ ok; failure; message; rect }.
    param([IntPtr]$HWnd, [string]$Dest)
    $v = Get-EbiWindowVisibleRect -HWnd $HWnd
    if (-not $v['ok']) { return @{ ok = $false; failure = 'no_window'; message = 'the window has no rectangle (closed?)'; rect = $null } }
    $r = Save-EbiScreenRegionPng -X $v['X'] -Y $v['Y'] -W $v['W'] -H $v['H'] -Dest $Dest
    return @{ ok = [bool]$r['ok']; failure = [string]$r['failure']; message = [string]$r['message']; rect = $v }
}

function ScreenStitchScroll-Result {
    # The output hashtable (every declared key), from what is known so far.
    param($Tiles, $Starts, $How, [string]$Stitched, [int]$W, [int]$H, $Layout, [int]$Covered, [int]$Need, [bool]$Reached, [bool]$AtEnd)
    return @{ tiles = @($Tiles); starts = @($Starts); how = @($How); stitched = $Stitched; width = $W; height = $H; layout = $Layout;
              coveredCols = $Covered; needCols = $Need; reachedEnd = $Reached; atScrollEnd = $AtEnd }
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $exe = Resolve-EbiWorkPath -PathValue ([string]$In['exe']) -WorkDir $work
    $file = Resolve-EbiWorkPath -PathValue ([string]$In['file']) -WorkDir $work
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['saveDir']) -WorkDir $work
    $stem = if ([string]$In['name'] -ne '') { [string]$In['name'] } else { [System.IO.Path]::GetFileName($file) }
    $title = if ([string]$In['windowTitle'] -ne '') { [string]$In['windowTitle'] } else { [System.IO.Path]::GetFileName($file) }
    $need = [Math]::Max(0, [int]$In['needCols'])
    $wide = if ([bool]$In['stitch']) { Join-Path $dir ($stem + '__wide.png') } else { '' }
    $empty = @{ textLeft = [int]$In['textLeft']; textRight = [int]$In['textRight']; bandTop = [int]$In['bandTop']; bandBottom = [int]$In['bandBottom']; charPx = [int]$In['charPx']; pageCols = 0; tileW = 0; tileH = 0; starts = @() }
    if ($Ctx['DryRun']) {
        $pc = if ([int]$In['width'] -gt 0) { [int][Math]::Floor(([int]$In['width'] - [int]$In['textLeft'] - [int]$In['textRight']) / [double][Math]::Max(1, [int]$In['charPx'])) } else { 0 }
        $n = ScreenStitchScroll-Plan -NeedCols $need -PageCols $pc -MaxTiles ([int]$In['maxTiles'])
        $tiles = @(for ($k = 1; $k -le $n; $k++) { Join-Path $dir ('{0}__{1:D2}.png' -f $stem, $k) })
        $Ctx.Log.Info(('would open {0} in {1}, capture it page by page to the right{2} into {3}{4}' -f $file, $exe, $(if ($need -gt 0) { (' up to column {0} ({1} captures at {2} columns a page)' -f $need, $n, $(if ($pc -gt 0) { $pc } else { '?' })) } else { ' until it cannot scroll further' }), $dir, $(if ($wide) { ' and merge them into ' + $wide } else { '' })))
        $r = ScreenStitchScroll-Result -Tiles $tiles -Starts @() -How @() -Stitched $wide -W 0 -H 0 -Layout $empty -Covered 0 -Need $need -Reached $false -AtEnd $false
        $r['ok'] = $true; return $r
    }
    $fail = { param($f, $m) $o = ScreenStitchScroll-Result -Tiles @() -Starts @() -How @() -Stitched '' -W 0 -H 0 -Layout $empty -Covered 0 -Need $need -Reached $false -AtEnd $false; $o['ok'] = $false; $o['failure'] = $f; $o['message'] = $m; return $o }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { return (& $fail 'file_not_found' ('viewer not found: ' + $exe)) }
    if (-not (Test-Path -LiteralPath $file -PathType Leaf)) { return (& $fail 'file_not_found' ('text file not found: ' + $file)) }

    $before = @(Get-EbiTopWindows | Where-Object { ([string]$_['title']).IndexOf($title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { [string]$_['handle'] })
    $argList = @(@($In['args']) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { [string]$_ }) + @('"' + $file.Replace('"', '') + '"')
    $proc = $null
    try { $proc = Start-Process -FilePath $exe -ArgumentList $argList -PassThru } catch { return (& $fail 'file_not_found' $_.Exception.Message) }
    $pick = @{ handle = [IntPtr]::Zero; ours = $false }
    $deadline = (Get-Date).AddSeconds([Math]::Max(2, [int]$In['windowWaitSec']))
    while ((Get-Date) -lt $deadline) {
        $pick = ScreenStitchScroll-PickWindow -Windows @(Get-EbiTopWindows) -Before $before -Title $title
        if ($pick['handle'] -ne [IntPtr]::Zero -and ($pick['ours'] -or $proc.HasExited)) { break }
        Start-Sleep -Milliseconds 250
    }
    $h = [IntPtr]$pick['handle']
    if ($h -eq [IntPtr]::Zero) {
        try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
        return (& $fail 'no_window' ('no window titled "*{0}*" appeared within {1}s' -f $title, $In['windowWaitSec']))
    }
    $tiles = New-Object System.Collections.ArrayList
    $starts = New-Object System.Collections.ArrayList
    $how = New-Object System.Collections.ArrayList
    $warnings = New-Object System.Collections.ArrayList
    $layout = $empty; $covered = 0; $atEnd = $false; $outW = 0; $outH = 0   # not $W / $H: names are case-blind and $h is the window
    try {
        Start-Sleep -Milliseconds ([int]$In['settleMs'])
        $wantH = if ([int]$In['height'] -gt 0) { [int]$In['height'] } else { ScreenStitchScroll-FitHeight -Lines ([int]$In['fitLines']) -RowPx ([int]$In['rowPx']) -ChromePx ([int]$In['chromePx']) -MaxHeight ([int]$In['maxHeight']) }
        if ([int]$In['width'] -gt 0 -or $wantH -gt 0) {
            # one of the two given: the other stays as the viewer opened it
            $cur = Get-EbiWindowVisibleRect -HWnd $h
            $ww = if ([int]$In['width'] -gt 0) { [int]$In['width'] } else { [int]$cur['W'] }
            $hh = if ($wantH -gt 0) { $wantH } else { [int]$cur['H'] }
            if ($ww -gt 0 -and $hh -gt 0) { [void](Set-EbiWindowVisibleSize -HWnd $h -X ([int]$In['x']) -Y ([int]$In['y']) -W $ww -H $hh) }
        }
        $fg = Set-EbiForeground -HWnd $h
        if (-not $fg['ok']) { return (& $fail 'foreground_lost' $fg['message']) }
        if ([string]$In['keysFirst'] -ne '') { Send-EbiKeys -Keys ([string]$In['keysFirst']) -WaitMs ([int]$In['settleMs']) }
        $p0 = Join-Path $dir ('{0}__01.png' -f $stem)
        $s = ScreenStitchScroll-Shot -HWnd $h -Dest $p0
        if (-not $s['ok']) { return (& $fail $(if ($s['failure'] -eq 'no_window') { 'no_window' } else { 'save_failed' }) $s['message']) }
        $sz = Get-EbiPngSize -Path $p0
        $g = Get-EbiStitchGeometry -TileW ([int]$sz['width']) -TileH ([int]$sz['height']) -TextLeft ([int]$In['textLeft']) -TextRight ([int]$In['textRight']) -BandTop ([int]$In['bandTop']) -BandBottom ([int]$In['bandBottom']) -CharPx ([int]$In['charPx'])
        if (-not $g['ok']) { return (& $fail 'geometry' ($g['message'] + ' -- set textLeft / textRight / bandTop / bandBottom / charPx for this window')) }
        $sc = Get-EbiHScroll -HWnd $h
        $pageCols = [int]$g['pageCols']
        if ($sc['ok'] -and [int]$sc['page'] -gt 0 -and [int]$sc['page'] -ne $pageCols) {
            [void]$warnings.Add(@{ code = 'page_mismatch'; message = ('the scroll bar says {0} columns a page, the capture''s text area {1} ({2} px / {3} px a column): check textLeft / textRight / charPx' -f $sc['page'], $pageCols, $g['textW'], $g['charPx']) })
        }
        [void]$tiles.Add($p0); [void]$starts.Add($(if ($sc['ok']) { [int]$sc['pos'] } else { 0 }))
        $covered = [int]$starts[0] + $pageCols
        $methods = @(@($In['advance']) | ForEach-Object { [string]$_ } | Where-Object { @('message', 'click', 'keys') -contains $_ })
        if ($methods.Count -eq 0) { $methods = @('message', 'click') }
        while (($need -le 0 -or $covered -lt $need) -and $tiles.Count -lt [int]$In['maxTiles']) {
            if ($sc['ok'] -and [bool]$sc['atEnd']) { $atEnd = $true; break }
            $prev = [string]$tiles[$tiles.Count - 1]
            $next = Join-Path $dir ('{0}__{1:D2}.png' -f $stem, ($tiles.Count + 1))
            $moved = $null; $sc2 = $sc
            foreach ($m in $methods) {
                $fg = Set-EbiForeground -HWnd $h
                if (-not $fg['ok']) { return (& $fail 'foreground_lost' $fg['message']) }
                if ($m -eq 'message') { if (-not (Invoke-EbiHScrollPage -Scroll $sc -SettleMs ([int]$In['settleMs']))) { continue } }
                elseif ($m -eq 'click') {
                    $v = Get-EbiWindowVisibleRect -HWnd $h
                    $pt = ScreenStitchScroll-TroughPoint -BarRect $(if ($sc['ok']) { $sc['rect'] } else { $null }) -Visible $v -Right ([int]$In['clickRight']) -Bottom ([int]$In['clickBottom'])
                    Invoke-EbiClick -X ([int]$pt['x']) -Y ([int]$pt['y']) -SettleMs ([int]$In['settleMs'])
                } else {
                    if ([string]$In['keysAdvance'] -eq '') { continue }
                    Send-EbiKeys -Keys ([string]$In['keysAdvance']) -WaitMs ([int]$In['settleMs'])
                }
                $s = ScreenStitchScroll-Shot -HWnd $h -Dest $next
                if (-not $s['ok']) { return (& $fail $(if ($s['failure'] -eq 'no_window') { 'no_window' } else { 'save_failed' }) $s['message']) }
                $sc2 = Get-EbiHScroll -HWnd $h
                $match = $null
                if (-not ($sc['ok'] -and $sc2['ok'])) { $match = Find-EbiTileOverlap -PathA $prev -PathB $next -Geometry $g -MinOverlapPx (40 * [int]$g['charPx']) }
                $shift = Resolve-EbiTileShift -ScrollA $sc -ScrollB $sc2 -Match $match -PageCols $pageCols -CharPx ([int]$g['charPx'])
                if (-not [bool]$shift['same']) { $moved = $shift; break }
                $Ctx.Log.Debug(('advance {0}: the view did not move' -f $m))
            }
            if ($null -eq $moved) {
                $atEnd = $true
                if (Test-Path -LiteralPath $next) { Remove-Item -LiteralPath $next -Force -ErrorAction SilentlyContinue }
                break
            }
            [void]$tiles.Add($next); [void]$starts.Add([int]$starts[$starts.Count - 1] + [int]$moved['cols']); [void]$how.Add([string]$moved['how'])
            $covered = [int]$starts[$starts.Count - 1] + $pageCols
            $sc = $sc2
            $Ctx.Log.Info(('capture {0}: columns {1}..{2} ({3})' -f $tiles.Count, $starts[$starts.Count - 1], ($covered - 1), $moved['how']))
        }
        if (@($how | Where-Object { $_ -eq 'page' }).Count -gt 0 -and $atEnd) {
            [void]$warnings.Add(@{ code = 'seam_assumed'; message = 'the scroll bar could not be read: the captures were joined a page apart, and the last scroll may have been cut short at the end -- look at the last seam' })
        }
        $layout = @{ textLeft = [int]$g['textLeft']; textRight = [int]$g['textRight']; bandTop = [int]$g['bandTop']; bandBottom = [int]$g['bandBottom']; charPx = [int]$g['charPx']; pageCols = $pageCols; tileW = [int]$g['tileW']; tileH = [int]$g['tileH']; starts = $starts.ToArray() }
        if ($wide -ne '') {
            $lay = Get-EbiStitchLayout -Starts $starts.ToArray() -Geometry $g -Frame ([string]$In['frame'])
            $r = New-EbiStitchedPng -Tiles $tiles.ToArray() -Layout $lay -Dest $wide
            if (-not $r['ok']) { return (& $fail 'save_failed' ('merging the captures: ' + $r['message'])) }
            $outW = [int]$r['width']; $outH = [int]$r['height']
            $Ctx.Log.Info(('merged {0} captures into {1} ({2} x {3} px)' -f $tiles.Count, $wide, $outW, $outH))
        }
    } finally {
        if ([bool]$In['close'] -and [bool]$pick['ours']) {
            $c = Close-EbiWindow -HWnd $h -WaitMs 3000
            if (-not $c['ok']) { try { if (-not $proc.HasExited) { $proc.Kill() } } catch { } }
        }
    }
    $reached = if ($need -gt 0) { $covered -ge $need } else { $atEnd }
    $out = ScreenStitchScroll-Result -Tiles $tiles.ToArray() -Starts $starts.ToArray() -How $how.ToArray() -Stitched $wide -W $outW -H $outH -Layout $layout -Covered $covered -Need $need -Reached $reached -AtEnd $atEnd
    if ($warnings.Count -gt 0) { $out['warnings'] = $warnings.ToArray() }
    if ($need -gt 0 -and -not $reached) {
        $out['ok'] = $false; $out['failure'] = 'not_at_line_end'
        $out['message'] = ('the captures reach column {0}, the line end needs {1} ({2}; {3} captures kept in {4}) -- does the viewer wrap lines, or stop short of the line width?' -f $covered, $need, $(if ($atEnd) { 'the view would not scroll further' } else { 'maxTiles reached' }), $tiles.Count, $dir)
        return $out
    }
    $out['ok'] = $true
    return $out
}
