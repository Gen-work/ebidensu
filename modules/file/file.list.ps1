# modules/file/file.list.ps1
# The files in a folder, with what pairing and evidence need to know about
# each: size, time, and (optionally) the line count. Sorted by name or by
# time; the order is the arrival order verify.pair_files relies on.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'file.list'
  group      = 'file'
  summary    = 'List the files in a folder (size, time, optional line count) in order'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    dir        = @{ type='path';   required=$true; desc='relative: under the work dir' }
    glob       = @{ type='string'; default='*' }
    orderBy    = @{ type='string'; default='name'; enum=@('name', 'time') }
    countLines = @{ type='bool';   default=$false; desc='read each file and count its lines (CRLF / LF)' }
    mustExist  = @{ type='bool';   default=$true; desc='false: a missing folder lists as empty' }
  }
  outputs    = @{
    files = @{ type='list'; desc='@{ name; path; size; modified (ISO); lines; order } in order' }
    names = @{ type='list' }
    paths = @{ type='list' }
    total = @{ type='int' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
  )
  example    = @{ use = 'file.list'; with = @{ dir = 'DATA/GIFT/{{item.JOB}}'; glob = '*.csv'; countLines = $true } }
  notes      = 'order is the sort key used (the name, or the time as yyyyMMddHHmmss), so a later pairing step can sort the two sides the same way.'
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['dir']) -WorkDir ([string]$Ctx['WorkDir'])
    $glob = [string]$In['glob']; if ([string]::IsNullOrWhiteSpace($glob)) { $glob = '*' }
    if (-not (Test-Path -LiteralPath $dir -PathType Container)) {
        if ($Ctx['DryRun'] -or -not [bool]$In['mustExist']) {
            if ($Ctx['DryRun']) { $Ctx.Log.Info(('would list {0} in {1}' -f $glob, $dir)) }
            return @{ ok = $true; files = @(); names = @(); paths = @(); total = 0 }
        }
        return @{ ok = $false; failure = 'file_not_found'; message = ('no such folder: ' + $dir); files = @(); names = @(); paths = @(); total = 0 }
    }
    $items = @(Get-ChildItem -LiteralPath $dir -File -Filter $glob -ErrorAction SilentlyContinue)
    $items = if ([string]$In['orderBy'] -eq 'time') { @($items | Sort-Object LastWriteTime, Name) } else { @($items | Sort-Object Name) }
    $files = New-Object System.Collections.ArrayList
    foreach ($f in $items) {
        $lines = -1
        if ([bool]$In['countLines']) { $lines = @(Get-EbiTextLines -Text ((ConvertFrom-EbiMixedBytes -Bytes ([System.IO.File]::ReadAllBytes($f.FullName)))['text'])).Count }
        $order = if ([string]$In['orderBy'] -eq 'time') { $f.LastWriteTime.ToString('yyyyMMddHHmmss') } else { $f.Name }
        [void]$files.Add(@{ name = $f.Name; path = $f.FullName; size = [int64]$f.Length; modified = $f.LastWriteTime.ToString('s'); lines = $lines; order = $order })
    }
    return @{ ok = $true; files = $files.ToArray(); names = @($files | ForEach-Object { $_['name'] }); paths = @($files | ForEach-Object { $_['path'] }); total = $files.Count }
}
