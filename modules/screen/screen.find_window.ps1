# modules/screen/screen.find_window.ps1
# Find a top-level window by title (and optionally process), bring it to
# the front, optionally maximize / resize it, and register it as a
# 'window' resource. browser.ensure takes "the" main window of a process;
# with two browser windows open -- one per site -- only the title tells
# them apart, and that is what this step matches.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')

$Manifest = @{
  id         = 'screen.find_window'
  group      = 'screen'
  summary    = 'Find a window by title text, bring it to front, register it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @('window')
  releases   = @()
  idempotent = $true
  inputs     = @{
    title    = @{ type='string'; required=$true; desc='text the window title must contain (case-insensitive)' }
    process  = @{ type='string'; default=''; desc='process name the window must belong to (without .exe); empty = any' }
    state    = @{ type='string'; default='keep'; enum=@('keep', 'maximize', 'restore'); desc='window state to put it in' }
    width    = @{ type='int'; default=0; desc='with state restore: resize the VISIBLE window to this (0 = keep)' }
    height   = @{ type='int'; default=0 }
    settleMs = @{ type='int'; default=500 }
  }
  outputs    = @{
    title     = @{ type='string'; desc='the full title of the window found' }
    processId = @{ type='int' }
    found     = @{ type='int'; desc='how many windows matched (the front-most is used)' }
    width     = @{ type='int'; desc='visible width after the step' }
    height    = @{ type='int'; desc='visible height after the step' }
  }
  failures   = @(
    @{ id = 'not_found';       transient = $true }
    @{ id = 'foreground_lost'; transient = $true }
  )
  example    = @{ use = 'screen.find_window'; with = @{ title = '{{profile.windows.list.title}}'; state = 'maximize'; as = 'listWindow' } }
  notes      = 'not_found is transient: the usual cause is the page not being open yet; with onError ask the operator opens it and answers r. Several matches: the front-most (Z order) wins and found says how many.'
}

function Invoke-Step {
    param($In, $Ctx)
    $title = [string]$In['title']; $proc = [string]$In['process']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would find the window titled *{0}*{1} and bring it to front ({2})' -f $title, $(if ($proc) { ' of ' + $proc } else { '' }), $In['state'])); return @{ ok = $true; resource = $null; title = ''; processId = 0; found = 0; width = 0; height = 0 } }
    $hits = @(Get-EbiTopWindows | Where-Object { $_['title'].IndexOf($title, [System.StringComparison]::OrdinalIgnoreCase) -ge 0 -and ($proc -eq '' -or [string]::Equals($_['processName'], $proc, [System.StringComparison]::OrdinalIgnoreCase)) })
    if ($hits.Count -eq 0) { return @{ ok = $false; failure = 'not_found'; message = ('no window title contains "' + $title + '"' + $(if ($proc) { ' (process ' + $proc + ')' } else { '' }) + ' -- is the page open?'); title = ''; processId = 0; found = 0; width = 0; height = 0 } }
    $w = $hits[0]
    $h = $w['handle']
    switch ([string]$In['state']) {
        'maximize' { if (-not [bool]$w['maximized']) { Set-EbiWindowState -HWnd $h -State 'maximize' -SettleMs ([int]$In['settleMs']) } }
        'restore'  {
            if ([bool]$w['maximized'] -or [bool]$w['minimized']) { Set-EbiWindowState -HWnd $h -State 'restore' -SettleMs ([int]$In['settleMs']) }
            if ([int]$In['width'] -gt 0 -and [int]$In['height'] -gt 0) {
                $r = Get-EbiWindowVisibleRect -HWnd $h
                [void](Set-EbiWindowVisibleSize -HWnd $h -X ([Math]::Max(0, [int]$r['X'])) -Y ([Math]::Max(0, [int]$r['Y'])) -W ([int]$In['width']) -H ([int]$In['height']) -SettleMs ([int]$In['settleMs']))
            }
        }
    }
    $fg = Set-EbiForeground -HWnd $h -SettleMs ([int]$In['settleMs'])
    $vr = Get-EbiWindowVisibleRect -HWnd $h
    if (-not $fg['ok']) { return @{ ok = $false; failure = 'foreground_lost'; message = $fg['message']; title = [string]$w['title']; processId = [int]$w['processId']; found = $hits.Count; width = [int]$vr['W']; height = [int]$vr['H'] } }
    $out = @{ ok = $true; resource = $h; title = [string]$w['title']; processId = [int]$w['processId']; found = $hits.Count; width = [int]$vr['W']; height = [int]$vr['H'] }
    if ([string]$In['state'] -eq 'restore' -and [int]$In['width'] -gt 0 -and ([int]$vr['W'] -ne [int]$In['width'] -or [int]$vr['H'] -ne [int]$In['height'])) {
        $out['warnings'] = @(@{ code = 'size_not_applied'; message = ('asked for ' + $In['width'] + 'x' + $In['height'] + ', the window is ' + $vr['W'] + 'x' + $vr['H'] + ' (DPI scaling?)') })
    }
    if ($hits.Count -gt 1) { $out['warnings'] = @(@($out['warnings']) + @(@{ code = 'several_windows'; message = ('' + $hits.Count + ' windows match "' + $title + '"; using the front-most: ' + $w['title']) }) | Where-Object { $null -ne $_ }) }
    return $out
}
