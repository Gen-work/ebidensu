# modules/file/file.set_aside.ps1
# Move files out of the way into a folder, never overwriting (10/23 track:
# with the wrong start time, another job's file was downloaded into
# DATA\GFIX\<job>; delivery copies that folder whole, so a "t 10:15" at the
# gate sets the files this item just downloaded aside before it goes back).
# A name already in the folder gets _2, _3 ... A path that is not there
# any more is passed over.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath

$Manifest = @{
  id         = 'file.set_aside'
  group      = 'file'
  summary    = 'Move files into a folder out of the way (never overwrites; missing files passed over)'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    paths = @{ type='list'; default=@(); desc='the files (relative: under the work dir)' }
    toDir = @{ type='path'; required=$true; desc='the folder they go to (created)' }
  }
  outputs    = @{
    moved = @{ type='list'; desc='where each moved file is now' }
    total = @{ type='int'; desc='files moved' }
  }
  failures   = @(
    @{ id = 'move_failed'; transient = $true }
  )
  example    = @{ use = 'file.set_aside'; with = @{ paths = '{{steps.dl.out.paths}}'; toDir = 'capture/{{item.keySafe}}/set-aside' } }
  notes      = 'Idempotent: a second run finds the files gone and moves nothing.'
}

function FileSetAside-FreeName {
    # PURE given an "is taken" test: name.ext -> name.ext, name_2.ext, ...
    param([string]$Dir, [string]$Name, [scriptblock]$Taken)
    $base = [System.IO.Path]::GetFileNameWithoutExtension($Name)
    $ext = [System.IO.Path]::GetExtension($Name)
    $cand = $Name
    $n = 1
    while (& $Taken (Join-Path $Dir $cand)) { $n++; $cand = $base + '_' + $n + $ext }
    return (Join-Path $Dir $cand)
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['toDir']) -WorkDir ([string]$Ctx['WorkDir'])
    $files = @(@($In['paths']) | Where-Object { $null -ne $_ -and [string]$_ -ne '' } | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir ([string]$Ctx['WorkDir']) })
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would set {0} file(s) aside into {1}' -f $files.Count, $dir)); return @{ ok = $true; moved = @(); total = 0 } }
    $moved = New-Object System.Collections.ArrayList
    foreach ($f in $files) {
        if (-not (Test-Path -LiteralPath $f -PathType Leaf)) { continue }
        try {
            if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
            $dest = FileSetAside-FreeName -Dir $dir -Name (Split-Path -Path $f -Leaf) -Taken { param($p) Test-Path -LiteralPath $p }
            Move-Item -LiteralPath $f -Destination $dest -ErrorAction Stop
            [void]$moved.Add($dest)
            $Ctx.Log.Info(('set aside {0} -> {1}' -f $f, $dest))
        } catch { return @{ ok = $false; failure = 'move_failed'; message = ($f + ': ' + $_.Exception.Message); moved = $moved.ToArray(); total = $moved.Count } }
    }
    return @{ ok = $true; moved = $moved.ToArray(); total = $moved.Count }
}
