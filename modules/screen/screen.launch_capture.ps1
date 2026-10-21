# modules/screen/screen.launch_capture.ps1
# Run a desktop program once per argument set (a diff tool on a pair of
# files), size its window, capture it, optionally send keys and capture
# again (top of the data, then its end), and close it. One step for the
# whole list because a workflow has no inner loop: a transfer of three
# files is three runs of the tool inside one item.
#
# Captures use the window's VISIBLE bounds (kernel/Native.ps1
# Get-EbiWindowVisibleRect), so no strip of the desktop shows at the
# edges; the size asked for is the visible size too.
#
# fitStrip (10/21, SJODWSPN review: the first picture showed one line too
# many): a fixed window height is a guess -- 458 was derived, never seen,
# and the 17th line showed. With fitStrip set, each first / END capture
# counts the text bands in the line-number strip and resizes the window
# until it shows exactly the lines wanted, no part of the next one.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')    # Save-EbiScreenRegionPng, Get-EbiInkCounts
. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')   # Get-EbiInkBands

$Manifest = @{
  id         = 'screen.launch_capture'
  group      = 'screen'
  summary    = 'Run a program per argument set, capture its window (and again after keys), close it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    exe          = @{ type='path';   required=$true; desc='the program' }
    argSets      = @{ type='list';   required=$true; desc='one map per run (e.g. verify.pair_files pairs)' }
    argFields    = @{ type='list';   default=@('left', 'right'); desc='fields of each map passed as quoted arguments, in order' }
    nameField    = @{ type='string'; default='rightName'; desc='field naming the capture files (its file stem is used); blank = run number' }
    secondField  = @{ type='string'; default='long'; desc='field that is true when a second capture is wanted' }
    keysFirst    = @{ type='string'; default='^{HOME}'; desc='keys before the first capture (SendKeys syntax)' }
    keysSecond   = @{ type='string'; default='^{END}'; desc='keys before the second capture (the END view)' }
    screenLines  = @{ type='int';    default=0; desc='data lines one capture shows; > 0 switches to the line rule: lines <= screenLines one capture, <= 2 x screenLines the second capture after keysMid, more after keysSecond' }
    linesField   = @{ type='string'; default='lines'; desc='field of each map holding its line count (the line rule)' }
    keysMid      = @{ type='string'; default='^{HOME}{PGDN}'; desc='keys before the second capture of a file of screenLines+1 .. 2 x screenLines lines: the next page, so it starts at line screenLines+1' }
    saveDir      = @{ type='path';   required=$true }
    x            = @{ type='int';    default=40 }
    y            = @{ type='int';    default=40 }
    width        = @{ type='int';    default=0; desc='visible window width; 0 = leave as the program opens it' }
    height       = @{ type='int';    default=0 }
    heightEnd    = @{ type='int';    default=0; desc='line rule: the visible window height for the END-view capture (a file of more than 2 x screenLines lines); 0 = height. 10/19: the HOST screen shows 17 records, so the END picture shows 17 DF lines while the first shows 16' }
    linesEnd     = @{ type='int';    default=0; desc='line rule: the lines the END-view capture must show (fitStrip); 0 = screenLines' }
    fitStrip     = @{ type='map';    default=@{}; desc='the line-number strip of a capture, @{ x; width; yFrom; yTo; ink; maxLuma } (screen.find_ink_rows): each first / END capture is resized until it shows exactly the lines wanted (first: min(screenLines, lines); END: linesEnd), no part of the next; empty = no fit' }
    fitPitchPx   = @{ type='int';    default=18; desc='px a line (DF: 18, measured 10/19)' }
    fitStepPx    = @{ type='int';    default=6; desc='fine resize step near the target' }
    fitTries     = @{ type='int';    default=8; desc='resizes per capture before giving up (warning fit_failed, the capture kept)' }
    windowWaitSec = @{ type='int';   default=15 }
    windowTitle  = @{ type='string'; default=''; desc='fallback when the started process shows no window of its own (a launcher / single-instance program): the top-level window whose title contains this' }
    closeOthers  = @{ type='bool';   default=$false; desc='close every window whose title contains windowTitle before each start, so the only such window is the new one' }
    settleMs     = @{ type='int';    default=1200; desc='wait after the window appears / after keys' }
  }
  outputs    = @{
    shots = @{ type='list'; desc='@{ first; last; name; firstLines; lastLines; firstHeight; lastHeight } per run (last empty when no second capture; *Lines = text lines counted by the fit, -1 when not fitted)' }
    paths = @{ type='list'; desc='every capture, in order' }
    runs  = @{ type='int' }
  }
  failures   = @(
    @{ id = 'file_not_found';  transient = $false }
    @{ id = 'no_window';       transient = $true  }
    @{ id = 'foreground_lost'; transient = $true  }
    @{ id = 'save_failed';     transient = $true  }
  )
  example    = @{ use = 'screen.launch_capture'; with = @{ exe = '{{profile.df.exe}}'; argSets = '{{steps.pair.out.pairs}}'; saveDir = 'capture/df/{{item.keySafe}}'; width = 1133; height = 429 } }
  notes      = 'An empty argSets is a successful no-op (runs = 0). The program is closed with WM_CLOSE and killed if it is still there 3 s later; a capture already saved is kept either way. With fitStrip, a height that had to change is carried to the next captures of the same call, and a capture that could not be fitted is kept with warning fit_failed (look at it before review).'
}

function ScreenLaunchCapture-FitPlan {
    <#
      PURE. Does this capture show exactly Want lines? Bands are the text
      bands of its line-number strip (top to bottom). A line is whole when
      its band is about as tall as the median band; a shorter LAST band is
      the next line cut by the window's bottom (a shorter first band is
      only the strip's own top edge). -> @{ done; delta; full; partial }
        delta  px to add to the window height before the next try
      Far off: whole lines at a time (Pitch), the last one in Step px, so
      the window grows / shrinks to the first height that shows Want whole
      lines and nothing of the next.
    #>
    param($Bands, [int]$Want, [int]$Pitch = 18, [int]$Step = 6)
    $b = @(@($Bands) | Where-Object { $null -ne $_ })
    if ($Want -le 0) { return @{ done = $true; delta = 0; full = $b.Count; partial = $false } }
    if ($b.Count -eq 0) { return @{ done = $false; delta = 0; full = 0; partial = $false } }
    $hs = @($b | ForEach-Object { [int]$_['height'] } | Sort-Object)
    $med = [int]$hs[[int][Math]::Floor($hs.Count / 2)]
    $partial = ($b.Count -ge 2 -and [int]$b[$b.Count - 1]['height'] -lt ($med - 2))
    $full = $b.Count - $(if ($partial) { 1 } else { 0 })
    if ($full -eq $Want -and -not $partial) { return @{ done = $true; delta = 0; full = $full; partial = $false } }
    if ($full -gt $Want) { $d = -((($full - $Want - 1) * $Pitch) + $Step) }
    elseif ($full -eq $Want) { $d = -$Step }
    else { $d = (($Want - $full - 1) * $Pitch) + $Step }
    return @{ done = $false; delta = $d; full = $full; partial = $partial }
}

function ScreenLaunchCapture-CountBands {
    # Text bands of the fit strip of a saved capture, or $null when unread.
    param([string]$Path, $Strip)
    $c = Get-EbiInkCounts -Path $Path -X ([int]$Strip['x']) -Width ([int]$Strip['width']) -YFrom ([int]$Strip['yFrom']) -YTo ([int]$Strip['yTo']) -Ink $(if ($Strip['ink']) { [string]$Strip['ink'] } else { 'dark' }) -MaxLuma $(if ($Strip['maxLuma']) { [int]$Strip['maxLuma'] } else { 160 })
    if (-not $c['ok']) { return $null }
    # minHeight 2, not find_ink_rows' 4: a sliver of the next line is what
    # the fit is there to remove
    return @(Get-EbiInkBands -Counts ([int[]]$c['counts']) -MinInk 2 -MergeGap 1 | Where-Object { [int]$_['height'] -ge 2 })
}

function ScreenLaunchCapture-FitShot {
    <#
      Capture to Dest; with a fit strip, resize and capture again until the
      picture shows exactly Want lines (Keys re-sent after each resize: the
      view has to be rebuilt for the new page size). -> @{ ok; failure;
      message; lines; height; fitted }
    #>
    param([IntPtr]$HWnd, [string]$Dest, $In, [int]$Height, [int]$Want, [string]$Keys)
    $r = ScreenLaunchCapture-Shot -HWnd $HWnd -Dest $Dest
    if (-not $r['ok']) { return @{ ok = $false; failure = $r['failure']; message = $r['message']; lines = -1; height = $Height; fitted = $false } }
    $strip = $In['fitStrip']
    if (-not ($strip -is [System.Collections.IDictionary]) -or $strip.Count -eq 0 -or $Want -le 0 -or [int]$In['width'] -le 0 -or $Height -le 0) {
        return @{ ok = $true; lines = -1; height = $Height; fitted = $true }
    }
    $h = $Height
    $plan = $null
    for ($try = 0; $try -le [int]$In['fitTries']; $try++) {
        $bands = ScreenLaunchCapture-CountBands -Path $Dest -Strip $strip
        if ($null -eq $bands) { return @{ ok = $true; lines = -1; height = $h; fitted = $false } }
        $plan = ScreenLaunchCapture-FitPlan -Bands $bands -Want $Want -Pitch ([int]$In['fitPitchPx']) -Step ([int]$In['fitStepPx'])
        if ([bool]$plan['done']) { return @{ ok = $true; lines = [int]$plan['full']; height = $h; fitted = $true } }
        if ([int]$plan['delta'] -eq 0 -or $try -eq [int]$In['fitTries']) { break }
        $h = $h + [int]$plan['delta']
        [void](Set-EbiWindowVisibleSize -HWnd $HWnd -X ([int]$In['x']) -Y ([int]$In['y']) -W ([int]$In['width']) -H $h)
        $fg = Set-EbiForeground -HWnd $HWnd
        if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; lines = -1; height = $h; fitted = $false } }
        if (-not [string]::IsNullOrEmpty($Keys)) { Send-EbiKeys -Keys $Keys -WaitMs ([int]$In['settleMs']) } else { Start-Sleep -Milliseconds ([int]$In['settleMs']) }
        $r = ScreenLaunchCapture-Shot -HWnd $HWnd -Dest $Dest
        if (-not $r['ok']) { return @{ ok = $false; failure = $r['failure']; message = $r['message']; lines = -1; height = $h; fitted = $false } }
    }
    return @{ ok = $true; lines = $(if ($plan) { [int]$plan['full'] } else { -1 }); height = $h; fitted = $false }
}

function ScreenLaunchCapture-Shot {
    # Capture the visible window to Dest. @{ ok; failure; message }.
    param([IntPtr]$HWnd, [string]$Dest)
    $v = Get-EbiWindowVisibleRect -HWnd $HWnd
    if (-not $v['ok']) { return @{ ok = $false; failure = 'no_window'; message = 'the window has no rectangle' } }
    return (Save-EbiScreenRegionPng -X $v['X'] -Y $v['Y'] -W $v['W'] -H $v['H'] -Dest $Dest)
}

function ScreenLaunchCapture-SecondPlan {
    <#
      PURE. Is there a second capture for this argument set, and with which
      keys? @{ two; keys }. The line rule (screenLines > 0, 2026-10-12, the
      reviewer's request): one capture shows screenLines lines; a file of up
      to 2 x screenLines lines gets its second capture starting at line
      screenLines+1 (keysMid: top, then one page down), a longer one the END
      view (keysSecond; view = 'end', the window is first resized to
      heightEnd). Without the rule: secondField decides, keysSecond.
    #>
    param($Set, [int]$ScreenLines, [string]$LinesField, [string]$SecondField, [string]$KeysMid, [string]$KeysSecond)
    if ($ScreenLines -gt 0 -and $LinesField -ne '' -and $Set.Contains($LinesField) -and $null -ne $Set[$LinesField]) {
        $n = [int]$Set[$LinesField]
        if ($n -le $ScreenLines) { return @{ two = $false; keys = ''; view = '' } }
        if ($n -le (2 * $ScreenLines)) { return @{ two = $true; keys = $KeysMid; view = 'mid' } }
        return @{ two = $true; keys = $KeysSecond; view = 'end' }
    }
    $two = ($SecondField -ne '' -and $Set.Contains($SecondField) -and [bool]$Set[$SecondField])
    return @{ two = $two; keys = $(if ($two) { $KeysSecond } else { '' }); view = $(if ($two) { 'second' } else { '' }) }
}

function ScreenLaunchCapture-Want {
    <#
      PURE. Lines a fitted capture must show: the first view min(screenLines,
      lines) (a short file shows all it has, blank below); the END view
      linesEnd (else screenLines). 0 = no fit (no line rule / line count).
    #>
    param($Set, $In, [string]$View)
    $sl = [int]$In['screenLines']
    if ($sl -le 0) { return 0 }
    if ($View -eq 'end') { if ([int]$In['linesEnd'] -gt 0) { return [int]$In['linesEnd'] } else { return $sl } }
    $lf = [string]$In['linesField']
    if ($lf -eq '' -or -not $Set.Contains($lf) -or $null -eq $Set[$lf]) { return 0 }
    return [Math]::Min($sl, [int]$Set[$lf])
}

function ScreenLaunchCapture-Stem {
    param($Set, [string]$Field, [int]$N)
    if (-not [string]::IsNullOrWhiteSpace($Field) -and $Set.Contains($Field) -and -not [string]::IsNullOrWhiteSpace([string]$Set[$Field])) {
        return [System.IO.Path]::GetFileNameWithoutExtension([string]$Set[$Field])
    }
    return ('run{0}' -f $N)
}

function ScreenLaunchCapture-PickWindow {
    <#
      PURE. Which listed top window is the run just started? -> handle or 0.
        Windows   @(@{ handle; title }) now on screen
        Before    handles that matched the title BEFORE the start (an old
                  diff window the operator left open must never be captured)
        Title     the title fragment (DF - )
        Names     file names the right window's title must show
      A new window with every name > any new window > a window with every
      name even if it was already there (a single-instance program that
      reused its old window for the new pair) > 0.
    #>
    param($Windows, $Before, [string]$Title, $Names)
    $cand = @(@($Windows) | Where-Object { [string]$_['title'] -ne '' -and ([string]$_['title']).IndexOf($Title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })
    $names = @(@($Names) | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    $hasAll = { param($t) foreach ($n in $names) { if (([string]$t).IndexOf([string]$n, [System.StringComparison]::OrdinalIgnoreCase) -lt 0) { return $false } }; return $true }
    $old = @(@($Before) | ForEach-Object { [string]$_ })
    $new = @($cand | Where-Object { $old -notcontains [string]$_['handle'] })
    foreach ($w in $new) { if (& $hasAll $w['title']) { return $w['handle'] } }
    if ($new.Count -gt 0 -and $names.Count -eq 0) { return $new[0]['handle'] }
    if ($names.Count -gt 0) { foreach ($w in $cand) { if (& $hasAll $w['title']) { return $w['handle'] } } }
    if ($new.Count -gt 0) { return $new[0]['handle'] }
    return [IntPtr]::Zero
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $exe = [string]$In['exe']
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['saveDir']) -WorkDir $work
    $sets = @(@($In['argSets']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $fields = @(@($In['argFields']) | ForEach-Object { [string]$_ })
    $shots = New-Object System.Collections.ArrayList
    $paths = New-Object System.Collections.ArrayList
    if ($Ctx['DryRun']) {
        $n = 0
        foreach ($s in $sets) {
            $n++
            $stem = ScreenLaunchCapture-Stem -Set $s -Field ([string]$In['nameField']) -N $n
            $two = (ScreenLaunchCapture-SecondPlan -Set $s -ScreenLines ([int]$In['screenLines']) -LinesField ([string]$In['linesField']) -SecondField ([string]$In['secondField']) -KeysMid ([string]$In['keysMid']) -KeysSecond ([string]$In['keysSecond']))['two']
            $f = Join-Path $dir ($stem + '__first.png'); $l = if ($two) { Join-Path $dir ($stem + '__last.png') } else { '' }
            $Ctx.Log.Info(('would run {0} {1} and capture {2}{3}' -f $exe, (@($fields | ForEach-Object { '"' + [string]$s[$_] + '"' }) -join ' '), $f, $(if ($two) { ' + ' + $l } else { '' })))
            [void]$shots.Add(@{ first = $f; last = $l; name = $stem }); [void]$paths.Add($f); if ($two) { [void]$paths.Add($l) }
        }
        if ($sets.Count -eq 0) { $Ctx.Log.Info(('would run {0} for no argument set' -f $exe)) }
        return @{ ok = $true; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $sets.Count }
    }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('program not found: ' + $exe); shots = @(); paths = @(); runs = 0 } }
    $warnings = New-Object System.Collections.ArrayList
    $fitDelta = 0   # px the fit had to add to height / heightEnd; carried to the next pair
    $n = 0
    foreach ($s in $sets) {
        $n++
        $stem = ScreenLaunchCapture-Stem -Set $s -Field ([string]$In['nameField']) -N $n
        $argList = @($fields | ForEach-Object { '"' + ([string]$s[$_]).Replace('"', '') + '"' })
        $wt = [string]$In['windowTitle']
        # file STEMS: a long title is cut by the program (DF showed
        # "...F202610090033.cs]"), so the extension cannot be relied on
        $names = @($fields | ForEach-Object { [System.IO.Path]::GetFileNameWithoutExtension([string]$s[$_]) })
        if ([bool]$In['closeOthers'] -and $wt -ne '') {
            foreach ($ow in @(Get-EbiTopWindows | Where-Object { ([string]$_['title']).IndexOf($wt, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 })) {
                $c = Close-EbiWindow -HWnd $ow['handle'] -WaitMs 3000
                if (-not $c['ok']) { return @{ ok = $false; failure = 'no_window'; message = ('an old window would not close: ' + $ow['title'] + ' -- close it by hand and r'); shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
            }
            Start-Sleep -Milliseconds 300
        }
        $before = @()
        if ($wt -ne '') { $before = @(Get-EbiTopWindows | Where-Object { ([string]$_['title']).IndexOf($wt, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 } | ForEach-Object { [string]$_['handle'] }) }
        $proc = $null
        try { $proc = Start-Process -FilePath $exe -ArgumentList $argList -PassThru }
        catch { return @{ ok = $false; failure = 'file_not_found'; message = $_.Exception.Message; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
        $h = [IntPtr]::Zero
        $deadline = (Get-Date).AddSeconds([Math]::Max(2, [int]$In['windowWaitSec']))
        while ((Get-Date) -lt $deadline) {
            try { $proc.Refresh() } catch { }
            if (-not $proc.HasExited -and $proc.MainWindowHandle -ne [IntPtr]::Zero) { $h = $proc.MainWindowHandle; break }
            if ($wt -ne '') {
                # the started process showed no window (single instance): find
                # the one showing THESE files, never an old one left open
                $p = ScreenLaunchCapture-PickWindow -Windows @(Get-EbiTopWindows) -Before $before -Title $wt -Names $names
                if ($p -ne [IntPtr]::Zero) { $h = $p; break }
            } elseif ($proc.HasExited) { break }
            Start-Sleep -Milliseconds 250
        }
        # the window must show THIS pair before it is captured (the title of a
        # diff tool names its files; an old window, or one still loading,
        # does not) -- wait for it within the same deadline
        if ($h -ne [IntPtr]::Zero -and $wt -ne '' -and $names.Count -gt 0) {
            $ok = $false
            while ($true) {
                $t = ''
                try { $t = [EbiNative]::TitleOf($h) } catch { }
                $miss = @($names | Where-Object { $t.IndexOf([string]$_, [System.StringComparison]::OrdinalIgnoreCase) -lt 0 })
                if ($miss.Count -eq 0) { $ok = $true; break }
                if ((Get-Date) -ge $deadline) { break }
                Start-Sleep -Milliseconds 250
            }
            if (-not $ok) {
                try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
                return @{ ok = $false; failure = 'no_window'; message = ('the window found does not show ' + ($names -join ' / ') + ' (title: ' + $t + ') -- close old ' + $wt + 'windows and r'); shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 }
            }
        }
        if ($h -eq [IntPtr]::Zero) {
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
            return @{ ok = $false; failure = 'no_window'; message = ('{0} showed no window within {1}s' -f $exe, $In['windowWaitSec']); shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 }
        }
        try {
            Start-Sleep -Milliseconds ([int]$In['settleMs'])
            $hFirst = [int]$In['height']
            if ($hFirst -gt 0) { $hFirst += $fitDelta }
            if ([int]$In['width'] -gt 0 -and $hFirst -gt 0) { [void](Set-EbiWindowVisibleSize -HWnd $h -X ([int]$In['x']) -Y ([int]$In['y']) -W ([int]$In['width']) -H $hFirst) }
            $fg = Set-EbiForeground -HWnd $h
            if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1; warnings = $warnings.ToArray() } }
            if (-not [string]::IsNullOrEmpty([string]$In['keysFirst'])) { Send-EbiKeys -Keys ([string]$In['keysFirst']) -WaitMs ([int]$In['settleMs']) }
            $first = Join-Path $dir ($stem + '__first.png')
            $wantFirst = ScreenLaunchCapture-Want -Set $s -In $In -View 'first'
            $r = ScreenLaunchCapture-FitShot -HWnd $h -Dest $first -In $In -Height $hFirst -Want $wantFirst -Keys ([string]$In['keysFirst'])
            if (-not $r['ok']) { return @{ ok = $false; failure = $(if ($r['failure'] -in @('no_window', 'foreground_lost')) { $r['failure'] } else { 'save_failed' }); message = $r['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1; warnings = $warnings.ToArray() } }
            [void]$paths.Add($first)
            $firstLines = [int]$r['lines']; $firstHeight = [int]$r['height']
            if ([bool]$r['fitted']) { if ($hFirst -gt 0) { $fitDelta += ($firstHeight - $hFirst) } }
            else { [void]$warnings.Add(@{ code = 'fit_failed'; message = ('{0}: the first picture shows {1} lines, wanted {2} -- look at it before review' -f $stem, $firstLines, $wantFirst); data = @{ path = $first; lines = $firstLines; want = $wantFirst; height = $firstHeight } }) }
            $last = ''; $lastLines = -1; $lastHeight = 0
            $plan2 = ScreenLaunchCapture-SecondPlan -Set $s -ScreenLines ([int]$In['screenLines']) -LinesField ([string]$In['linesField']) -SecondField ([string]$In['secondField']) -KeysMid ([string]$In['keysMid']) -KeysSecond ([string]$In['keysSecond'])
            if ([bool]$plan2['two']) {
                $hLast = $firstHeight
                if ([string]$plan2['view'] -eq 'end' -and [int]$In['heightEnd'] -gt 0 -and [int]$In['width'] -gt 0) {
                    $hLast = [int]$In['heightEnd'] + $fitDelta
                    if ($hLast -ne $firstHeight) {
                        [void](Set-EbiWindowVisibleSize -HWnd $h -X ([int]$In['x']) -Y ([int]$In['y']) -W ([int]$In['width']) -H $hLast)
                        Start-Sleep -Milliseconds 300
                    }
                }
                $fg = Set-EbiForeground -HWnd $h
                if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1; warnings = $warnings.ToArray() } }
                Send-EbiKeys -Keys ([string]$plan2['keys']) -WaitMs ([int]$In['settleMs'])
                $last = Join-Path $dir ($stem + '__last.png')
                # the PgDn page shows what is left of the file under a window
                # already fitted: nothing to fit; the END view is fitted
                $wantLast = if ([string]$plan2['view'] -eq 'end') { ScreenLaunchCapture-Want -Set $s -In $In -View 'end' } else { 0 }
                $r = ScreenLaunchCapture-FitShot -HWnd $h -Dest $last -In $In -Height $hLast -Want $wantLast -Keys ([string]$plan2['keys'])
                if (-not $r['ok']) { return @{ ok = $false; failure = $(if ($r['failure'] -in @('no_window', 'foreground_lost')) { $r['failure'] } else { 'save_failed' }); message = $r['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1; warnings = $warnings.ToArray() } }
                [void]$paths.Add($last)
                $lastLines = [int]$r['lines']; $lastHeight = [int]$r['height']
                if (-not [bool]$r['fitted']) { [void]$warnings.Add(@{ code = 'fit_failed'; message = ('{0}: the END picture shows {1} lines, wanted {2} -- look at it before review' -f $stem, $lastLines, $wantLast); data = @{ path = $last; lines = $lastLines; want = $wantLast; height = $lastHeight } }) }
            }
            [void]$shots.Add(@{ first = $first; last = $last; name = $stem; firstLines = $firstLines; lastLines = $lastLines; firstHeight = $firstHeight; lastHeight = $lastHeight })
            $Ctx.Log.Info(('captured {0}{1}{2}' -f $first, $(if ($last) { ' + ' + $last } else { '' }), $(if ($firstLines -ge 0) { (' (lines {0}{1}, height {2}{3})' -f $firstLines, $(if ($last) { '/' + $lastLines } else { '' }), $firstHeight, $(if ($last) { '/' + $lastHeight } else { '' })) } else { '' })))
        } finally {
            $c = Close-EbiWindow -HWnd $h -WaitMs 3000
            if (-not $c['ok']) { try { if (-not $proc.HasExited) { $proc.Kill() } } catch { } }
        }
    }
    return @{ ok = $true; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $sets.Count; warnings = $warnings.ToArray() }
}
