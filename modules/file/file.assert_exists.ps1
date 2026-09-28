# modules/file/file.assert_exists.ps1
# Fail unless a path exists (P1-23). What happens on the failure is the
# workflow's onError policy (retry / skip / ask): the step only states the
# fact. In a dry run a missing path is a warning, so `ebi dryrun` walks the
# whole plan on a machine that has none of the files.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath

$Manifest = @{
  id         = 'file.assert_exists'
  group      = 'file'
  summary    = 'Fail with file_not_found unless the path exists'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path = @{ type='path';   required=$true; desc='relative paths resolve under the work dir' }
    kind = @{ type='string'; default='any'; enum=@('any','file','dir'); desc='what it must be' }
  }
  outputs    = @{
    exists = @{ type='bool' }
    path   = @{ type='path'; desc='the resolved path that was checked' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true }
  )
  example    = @{ use = 'file.assert_exists'; with = @{ path = '{{steps.find.out.path}}'; kind = 'file' } }
  notes      = 'file_not_found is transient here for the same reason as in file.find: the file is usually on its way.'
}

function Invoke-Step {
    param($In, $Ctx)
    $p = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    $kind = [string]$In['kind']
    $exists = $false
    if ($p -ne '') {
        switch ($kind) {
            'file' { $exists = Test-Path -LiteralPath $p -PathType Leaf }
            'dir'  { $exists = Test-Path -LiteralPath $p -PathType Container }
            default { $exists = Test-Path -LiteralPath $p }
        }
    }
    if ($exists) { return @{ ok = $true; exists = $true; path = $p } }
    if ($Ctx['DryRun']) {
        $Ctx.Log.Warn(('would fail: {0} {1} does not exist' -f $kind, $p))
        return @{ ok = $true; exists = $false; path = $p; warnings = @( @{ code = 'would_fail'; message = ('no ' + $kind + ' at ' + $p); data = @{} } ) }
    }
    return @{ ok = $false; failure = 'file_not_found'; message = ('no ' + $kind + ' at ' + $p); exists = $false; path = $p }
}
