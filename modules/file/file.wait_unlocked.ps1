# modules/file/file.wait_unlocked.ps1
# Wait until nobody holds a shared file (10/20, the operator's own
# isWbsReady.bat made into a step): try to open it for read+write with NO
# sharing -- the strictest test, stricter than Excel's read-only flag --
# and while that fails say who holds it (Excel's ~$ owner file next to
# it) and look again every everySec seconds. q or Esc in the console
# stops the wait (Ctrl+C would kill the whole run). Put it in front of an
# excel.open of a shared workbook (the WBS, the review check sheet, an
# evidence book a reviewer may have open): the open then never comes back
# read-only for "someone has it" reasons. openWhenFree hands the file to
# its own program once it is free (the .bat's "open the WBS for me").

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath

$Manifest = @{
  id         = 'file.wait_unlocked'
  group      = 'file'
  summary    = 'Wait until a shared file can be opened exclusively; say who holds it meanwhile'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path         = @{ type='path';   required=$true; desc='the shared file; relative paths resolve under the work dir' }
    everySec     = @{ type='int';    default=60; desc='seconds between two looks (the .bat waited a minute)' }
    timeoutMin   = @{ type='int';    default=30; desc='give up after this many minutes (failure locked, r waits again); 0 = until it is free or q' }
    openWhenFree = @{ type='bool';   default=$false; desc='open the file in its own program once it is free (standalone use)' }
  }
  outputs    = @{
    free      = @{ type='bool';   desc='the file could be opened exclusively' }
    heldBy    = @{ type='string'; desc='who held it at the last look that failed ("" when Excel''s owner file says nothing or it was free at once)' }
    waitedSec = @{ type='int' }
    polls     = @{ type='int' }
    path      = @{ type='path' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
    @{ id = 'locked';         transient = $true }
    @{ id = 'operator_quit';  transient = $false }
  )
  example    = @{ use = 'file.wait_unlocked'; with = @{ path = '{{profile.paths.wbs}}'; everySec = 60; timeoutMin = 30 } }
  notes      = 'A file the run itself has open (an excel.open earlier in the same run) counts as held: put this step BEFORE that open. Under DryRun nothing is opened or waited for (free=true).'
}

function FileWaitUnlocked-Owner {
    <#
      PURE. The bytes of Excel's owner file (~$<name>) -> the user name, or
      ''. Byte 0 = length, the ANSI (CP932) name after it; at byte 54 the
      same length and the name in UTF-16LE (preferred).
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 2) { return '' }
    $n = [int]$Bytes[0]
    try {
        if ($Bytes.Length -ge (55 + 2 * $n) -and [int]$Bytes[54] -eq $n -and $n -gt 0) {
            $u = [System.Text.Encoding]::Unicode.GetString($Bytes, 55, 2 * $n).Trim()
            if ($u -ne '') { return $u }
        }
        if ($n -gt 0 -and $Bytes.Length -ge (1 + $n)) { return ([System.Text.Encoding]::GetEncoding(932)).GetString($Bytes, 1, $n).Trim() }
    } catch { }
    return ''
}

function FileWaitUnlocked-HeldBy {
    # The ~$ owner file next to Path -> name or '' (never throws).
    param([string]$Path)
    try {
        $lock = Join-Path (Split-Path -Path $Path -Parent) ('~$' + (Split-Path -Path $Path -Leaf))
        if (-not (Test-Path -LiteralPath $lock -PathType Leaf)) { return '' }
        return (FileWaitUnlocked-Owner -Bytes ([System.IO.File]::ReadAllBytes($lock)))
    } catch { return '' }
}

function FileWaitUnlocked-IsFree {
    # Can the file be opened for read+write with no sharing right now?
    param([string]$Path)
    $fs = $null
    try {
        $fs = [System.IO.File]::Open($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::ReadWrite, [System.IO.FileShare]::None)
        return $true
    } catch { return $false }
    finally { if ($null -ne $fs) { try { $fs.Close() } catch { } } }
}

function FileWaitUnlocked-StopKey {
    # q or Esc pressed in the console while waiting. No console -> never.
    try {
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Escape -or $k.KeyChar -eq 'q' -or $k.KeyChar -eq 'Q') { return $true }
        }
    } catch { }
    return $false
}

function Invoke-Step {
    param($In, $Ctx)
    $p = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $every = [Math]::Max(1, [int]$In['everySec'])
    $limit = [int]$In['timeoutMin']
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would wait until {0} is free (every {1} s)' -f $p, $every))
        return @{ ok = $true; free = $true; heldBy = ''; waitedSec = 0; polls = 0; path = $p }
    }
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $p; free = $false; heldBy = ''; waitedSec = 0; polls = 0; path = $p } }
    $start = Get-Date
    $polls = 0
    $who = ''
    $leaf = Split-Path -Path $p -Leaf
    while ($true) {
        $polls++
        if (FileWaitUnlocked-IsFree -Path $p) {
            $waited = [int]((Get-Date) - $start).TotalSeconds
            if ($polls -gt 1) { $Ctx.Log.Info(('{0} is free now (waited {1} s)' -f $leaf, $waited)) }
            if ([bool]$In['openWhenFree']) {
                try { Start-Process -FilePath $p | Out-Null; $Ctx.Log.Info(('opened ' + $leaf)) } catch { $Ctx.Log.Warn(('could not open ' + $leaf + ': ' + $_.Exception.Message)) }
            }
            return @{ ok = $true; free = $true; heldBy = $who; waitedSec = $waited; polls = $polls; path = $p }
        }
        $who = FileWaitUnlocked-HeldBy -Path $p
        $elapsed = ((Get-Date) - $start).TotalSeconds
        if ($limit -gt 0 -and $elapsed -ge ($limit * 60)) {
            return @{ ok = $false; failure = 'locked'; message = ('{0} is still held{1} after {2} min -- r waits again' -f $leaf, $(if ($who -ne '') { ' by ' + $who } else { '' }), $limit); free = $false; heldBy = $who; waitedSec = [int]$elapsed; polls = $polls; path = $p }
        }
        Write-Host ('  [{0}] {1} is held{2}; looking again in {3} s (q / Esc = stop)' -f (Get-Date).ToString('HH:mm:ss'), $leaf, $(if ($who -ne '') { ' by ' + $who } else { ' (Excel''s owner file says nobody: another program, or a lock not yet released)' }), $every) -ForegroundColor DarkYellow
        for ($s = 0; $s -lt $every; $s++) {
            if (FileWaitUnlocked-StopKey) { return @{ ok = $false; failure = 'operator_quit'; message = ('stopped waiting for ' + $leaf); free = $false; heldBy = $who; waitedSec = [int]((Get-Date) - $start).TotalSeconds; polls = $polls; path = $p } }
            Start-Sleep -Seconds 1
        }
    }
}
