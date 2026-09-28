# modules/file/file.write_json.ps1
# Write a value as a JSON file through kernel/Json.ps1 (atomic, UTF-8 no
# BOM, Japanese as characters). The sidecar step: <keySafe>.meta.json next
# to a capture, a step's verdict next to its evidence (P1-21, P0-R14).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Json.ps1')

$Manifest = @{
  id         = 'file.write_json'
  group      = 'file'
  summary    = 'Write a value to a JSON file (atomic, UTF-8 without BOM)'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path = @{ type='path'; required=$true; desc='where to write; relative paths resolve under the work dir' }
    data = @{ type='any';  required=$true; desc='the value (map, list, string, number, bool)' }
  }
  outputs    = @{ path = @{ type='path'; desc='the file that was written' } }
  failures   = @(
    @{ id = 'write_failed';    transient = $true  }
    @{ id = 'not_serializable'; transient = $false }
  )
  example    = @{ use = 'file.write_json'; with = @{ path = 'capture/before_transferStatus/{{item.keySafe}}.meta.json'; data = @{ key = '{{item.key}}'; capturedAt = '{{run.startedAt}}' } } }
}

function Invoke-Step {
    param($In, $Ctx)
    $dest = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if (-not (Test-EbiJsonSerializable -Value $In['data'])) { return @{ ok = $false; failure = 'not_serializable'; message = 'data holds a value JSON cannot carry (a COM object, a handle, or nesting past depth 20)'; path = $dest } }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would write JSON to {0}' -f $dest)); return @{ ok = $true; path = $dest } }
    $w = Write-EbiJson -Path $dest -Value $In['data']
    if (-not $w['ok']) { return @{ ok = $false; failure = 'write_failed'; message = $w['message']; path = $dest } }
    return @{ ok = $true; path = $dest }
}
