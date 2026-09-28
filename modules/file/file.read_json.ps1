# modules/file/file.read_json.ps1
# Read a JSON file through kernel/Json.ps1 into a hashtable. A file that is
# not there is NOT a failure: data is null and a warning says so, because
# the usual caller is "is there a sidecar from last time?" (P1-21, P0-R14).
# A file that is there but not JSON is a failure -- that is corruption, not
# absence.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Json.ps1')

$Manifest = @{
  id         = 'file.read_json'
  group      = 'file'
  summary    = 'Read a JSON file; a missing file gives data=null and a warning'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path = @{ type='path'; required=$true; desc='the file; relative paths resolve under the work dir' }
  }
  outputs    = @{
    data   = @{ type='any';  desc='the parsed value (hashtables, not objects); null when the file is missing' }
    exists = @{ type='bool' }
    path   = @{ type='path' }
  }
  failures   = @(
    @{ id = 'json_invalid'; transient = $false }
    @{ id = 'read_failed';  transient = $true  }
  )
  example    = @{ use = 'file.read_json'; with = @{ path = 'capture/before_transferStatus/{{item.keySafe}}.meta.json' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $r = Read-EbiJson -Path $src
    if (-not $r['exists']) {
        return @{ ok = $true; data = $null; exists = $false; path = $src; warnings = @( @{ code = 'not_found'; message = ('no file at ' + $src); data = @{} } ) }
    }
    if (-not $r['ok']) {
        $id = if ([string]$r['message'] -match 'not valid JSON|parse|invalid|unexpected|depth') { 'json_invalid' } else { 'read_failed' }
        return @{ ok = $false; failure = $id; message = $r['message']; data = $null; exists = $true; path = $src }
    }
    return @{ ok = $true; data = $r['value']; exists = $true; path = $src }
}
