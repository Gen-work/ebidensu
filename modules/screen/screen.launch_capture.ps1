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

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')    # Save-EbiScreenRegionPng

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
    keysSecond   = @{ type='string'; default='^{END}'; desc='keys before the second capture' }
    saveDir      = @{ type='path';   required=$true }
    x            = @{ type='int';    default=40 }
    y            = @{ type='int';    default=40 }
    width        = @{ type='int';    default=0; desc='visible window width; 0 = leave as the program opens it' }
    height       = @{ type='int';    default=0 }
    windowWaitSec = @{ type='int';   default=15 }
    windowTitle  = @{ type='string'; default=''; desc='fallback when the started process shows no window of its own (a launcher / single-instance program): the top-level window whose title contains this' }
    settleMs     = @{ type='int';    default=1200; desc='wait after the window appears / after keys' }
  }
  outputs    = @{
    shots = @{ type='list'; desc='@{ first; last; name } per run (last empty when no second capture)' }
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
  notes      = 'An empty argSets is a successful no-op (runs = 0). The program is closed with WM_CLOSE and killed if it is still there 3 s later; a capture already saved is kept either way.'
}

function ScreenLaunchCapture-Shot {
    # Capture the visible window to Dest. @{ ok; failure; message }.
    param([IntPtr]$HWnd, [string]$Dest)
    $v = Get-EbiWindowVisibleRect -HWnd $HWnd
    if (-not $v['ok']) { return @{ ok = $false; failure = 'no_window'; message = 'the window has no rectangle' } }
    return (Save-EbiScreenRegionPng -X $v['X'] -Y $v['Y'] -W $v['W'] -H $v['H'] -Dest $Dest)
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
            $two = ([string]$In['secondField'] -ne '' -and $s.Contains([string]$In['secondField']) -and [bool]$s[[string]$In['secondField']])
            $f = Join-Path $dir ($stem + '__first.png'); $l = if ($two) { Join-Path $dir ($stem + '__last.png') } else { '' }
            $Ctx.Log.Info(('would run {0} {1} and capture {2}{3}' -f $exe, (@($fields | ForEach-Object { '"' + [string]$s[$_] + '"' }) -join ' '), $f, $(if ($two) { ' + ' + $l } else { '' })))
            [void]$shots.Add(@{ first = $f; last = $l; name = $stem }); [void]$paths.Add($f); if ($two) { [void]$paths.Add($l) }
        }
        if ($sets.Count -eq 0) { $Ctx.Log.Info(('would run {0} for no argument set' -f $exe)) }
        return @{ ok = $true; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $sets.Count }
    }
    if (-not (Test-Path -LiteralPath $exe -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('program not found: ' + $exe); shots = @(); paths = @(); runs = 0 } }
    $n = 0
    foreach ($s in $sets) {
        $n++
        $stem = ScreenLaunchCapture-Stem -Set $s -Field ([string]$In['nameField']) -N $n
        $argList = @($fields | ForEach-Object { '"' + ([string]$s[$_]).Replace('"', '') + '"' })
        $wt = [string]$In['windowTitle']
        $names = @($fields | ForEach-Object { [System.IO.Path]::GetFileName([string]$s[$_]) })
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
        if ($h -eq [IntPtr]::Zero) {
            try { if (-not $proc.HasExited) { $proc.Kill() } } catch { }
            return @{ ok = $false; failure = 'no_window'; message = ('{0} showed no window within {1}s' -f $exe, $In['windowWaitSec']); shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 }
        }
        try {
            Start-Sleep -Milliseconds ([int]$In['settleMs'])
            if ([int]$In['width'] -gt 0 -and [int]$In['height'] -gt 0) { [void](Set-EbiWindowVisibleSize -HWnd $h -X ([int]$In['x']) -Y ([int]$In['y']) -W ([int]$In['width']) -H ([int]$In['height'])) }
            $fg = Set-EbiForeground -HWnd $h
            if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
            if (-not [string]::IsNullOrEmpty([string]$In['keysFirst'])) { Send-EbiKeys -Keys ([string]$In['keysFirst']) -WaitMs ([int]$In['settleMs']) }
            $first = Join-Path $dir ($stem + '__first.png')
            $r = ScreenLaunchCapture-Shot -HWnd $h -Dest $first
            if (-not $r['ok']) { return @{ ok = $false; failure = $(if ($r['failure'] -eq 'no_window') { 'no_window' } else { 'save_failed' }); message = $r['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
            [void]$paths.Add($first)
            $last = ''
            $sf = [string]$In['secondField']
            if ($sf -ne '' -and $s.Contains($sf) -and [bool]$s[$sf]) {
                $fg = Set-EbiForeground -HWnd $h
                if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
                Send-EbiKeys -Keys ([string]$In['keysSecond']) -WaitMs ([int]$In['settleMs'])
                $last = Join-Path $dir ($stem + '__last.png')
                $r = ScreenLaunchCapture-Shot -HWnd $h -Dest $last
                if (-not $r['ok']) { return @{ ok = $false; failure = 'save_failed'; message = $r['message']; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $n - 1 } }
                [void]$paths.Add($last)
            }
            [void]$shots.Add(@{ first = $first; last = $last; name = $stem })
            $Ctx.Log.Info(('captured {0}{1}' -f $first, $(if ($last) { ' + ' + $last } else { '' })))
        } finally {
            $c = Close-EbiWindow -HWnd $h -WaitMs 3000
            if (-not $c['ok']) { try { if (-not $proc.HasExited) { $proc.Kill() } } catch { } }
        }
    }
    return @{ ok = $true; shots = $shots.ToArray(); paths = $paths.ToArray(); runs = $sets.Count }
}
