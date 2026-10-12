# modules/file/file.move.ps1
# Move (or copy) a file to a folder or a new name, creating the folder
# (P4-14). The target folder that does not exist yet -- DATA\GFIX\<job>
# the first time a job is received -- is created, not an error.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath

$Manifest = @{
  id         = 'file.move'
  group      = 'file'
  summary    = 'Move or copy a file into a folder or to a new name (folders created)'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    from      = @{ type='path'; required=$true; desc='the file (relative: under the work dir)' }
    to        = @{ type='path'; required=$true; desc='destination file, or a folder when it ends with \ or / or already is one' }
    copy      = @{ type='bool'; default=$false; desc='copy instead of move' }
    overwrite = @{ type='bool'; default=$false; desc='replace an existing destination' }
  }
  outputs    = @{
    path = @{ type='path'; desc='where the file is now' }
    name = @{ type='string' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'target_exists';  transient = $false }
    @{ id = 'move_failed';    transient = $true  }
  )
  example    = @{ use = 'file.move'; with = @{ from = '{{steps.dl.out.path}}'; to = 'log/GFIXReceive/{{run.mmdd}}.log' } }
  notes      = 'Not idempotent: a move done once has no source the second time (put a flow.checkpoint after it, STEP-CONTRACT 6.4). A destination that already holds the very same bytes is not target_exists -- the source is just removed (move) or left (copy).'
}

function FileMove-SameBytes {
    param([string]$A, [string]$B)
    try {
        $fa = Get-Item -LiteralPath $A; $fb = Get-Item -LiteralPath $B
        if ($fa.Length -ne $fb.Length) { return $false }
        $x = [System.IO.File]::ReadAllBytes($A); $y = [System.IO.File]::ReadAllBytes($B)
        for ($i = 0; $i -lt $x.Length; $i++) { if ($x[$i] -ne $y[$i]) { return $false } }
        return $true
    } catch { return $false }
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['from']) -WorkDir $work
    $toRaw = [string]$In['to']
    $dst = Resolve-EbiWorkPath -PathValue $toRaw -WorkDir $work
    $isDir = $toRaw.EndsWith('\') -or $toRaw.EndsWith('/') -or (Test-Path -LiteralPath $dst -PathType Container)
    if ($isDir) { $dst = Join-Path $dst (Split-Path -Path $src -Leaf) }
    $verb = if ([bool]$In['copy']) { 'copy' } else { 'move' }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would {0} {1} -> {2}' -f $verb, $src, $dst)); return @{ ok = $true; path = $dst; name = (Split-Path -Path $dst -Leaf) } }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $src; path = ''; name = '' } }
    try {
        $dir = Split-Path -Path $dst -Parent
        if (-not [string]::IsNullOrWhiteSpace($dir) -and -not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if (Test-Path -LiteralPath $dst) {
            if (FileMove-SameBytes -A $src -B $dst) {
                if ($verb -eq 'move' -and $src -ne $dst) { Remove-Item -LiteralPath $src -Force }
                return @{ ok = $true; path = $dst; name = (Split-Path -Path $dst -Leaf); warnings = @(@{ code = 'already_there'; message = ('identical file already at ' + $dst) }) }
            }
            if (-not [bool]$In['overwrite']) { return @{ ok = $false; failure = 'target_exists'; message = ('a different file is already at ' + $dst); path = $dst; name = '' } }
        }
        if ($verb -eq 'copy') { Copy-Item -LiteralPath $src -Destination $dst -Force } else { Move-Item -LiteralPath $src -Destination $dst -Force }
    } catch { return @{ ok = $false; failure = 'move_failed'; message = $_.Exception.Message; path = $dst; name = '' } }
    $Ctx.Log.Info(('{0} {1} -> {2}' -f $verb, $src, $dst))
    return @{ ok = $true; path = $dst; name = (Split-Path -Path $dst -Leaf) }
}
