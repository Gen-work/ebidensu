# modules/screen/screen.save.ps1
# Put a captured image where the naming rule says (VOCABULARY.md 2.5):
#     <dir>/<keySafe>.<ext>            one image per key
#     <dir>/<keySafe>__<tag>.<ext>     several, told apart by tag
# The key is folded through kernel/Key.ps1's file-safe form so a workflow
# that passes {{item.key}} by mistake still lands on the same name as one
# that passes {{item.keySafe}} (P0-R4: file names never carry a bare key).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'screen.save'
  group      = 'screen'
  summary    = 'Move or copy an image to <dir>/<keySafe>[__<tag>].<ext>'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    source = @{ type='path';   required=$true; desc='the image to place' }
    dir    = @{ type='path';   required=$true; desc='destination folder, e.g. capture/before_transferStatus' }
    key    = @{ type='string'; required=$true; desc='the item key (keySafe form; a raw key is folded)' }
    tag    = @{ type='string'; default=''; desc='distinguishes several images of one key: <key>__<tag>' }
    ext    = @{ type='string'; default='png' }
    copy   = @{ type='bool';   default=$false; desc='copy instead of move (the source stays)' }
  }
  outputs    = @{
    path = @{ type='path';   desc='where the image is now' }
    name = @{ type='string'; desc='the file name that was chosen' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
    @{ id = 'save_failed';    transient = $true  }
  )
  example    = @{ use = 'screen.save'; with = @{ source = '{{steps.shot.out.path}}'; dir = 'capture/before_transferStatus'; key = '{{item.keySafe}}'; tag = 'row' } }
  notes      = 'An existing file of the same name is replaced: re-running a key re-captures it. A key that had to be folded is reported as a warning (key_folded).'
}

function ScreenSave-Name {
    # PURE. The file name for a key, tag and extension.
    param([string]$Key, [string]$Tag = '', [string]$Ext = 'png')
    $safe = ConvertTo-EbiKeySafeSegment -Value $Key
    $e = if ([string]::IsNullOrWhiteSpace($Ext)) { 'png' } else { $Ext.TrimStart('.') }
    $t = if ([string]::IsNullOrWhiteSpace($Tag)) { '' } else { '__' + (ConvertTo-EbiKeySafeSegment -Value $Tag) }
    return @{ name = ($safe + $t + '.' + $e); folded = ($safe -ne $Key) }
}

function Invoke-Step {
    param($In, $Ctx)
    $src = Resolve-EbiWorkPath -PathValue ([string]$In['source']) -WorkDir ([string]$Ctx['WorkDir'])
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['dir']) -WorkDir ([string]$Ctx['WorkDir'])
    $n = ScreenSave-Name -Key ([string]$In['key']) -Tag ([string]$In['tag']) -Ext ([string]$In['ext'])
    $dest = [System.IO.Path]::Combine($dir, $n['name'])
    $warnings = @()
    if ($n['folded']) { $warnings = @( @{ code = 'key_folded'; message = ('key "' + [string]$In['key'] + '" folded to the file-safe form'); data = @{ name = $n['name'] } } ) }
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would {0} {1} -> {2}' -f $(if ([bool]$In['copy']) { 'copy' } else { 'move' }), $src, $dest))
        return @{ ok = $true; path = $dest; name = $n['name']; warnings = $warnings }
    }
    if (-not (Test-Path -LiteralPath $src -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $src; path = $dest; name = $n['name'] } }
    try {
        if (-not (Test-Path -LiteralPath $dir)) { New-Item -ItemType Directory -Path $dir -Force | Out-Null }
        if ([bool]$In['copy']) { Copy-Item -LiteralPath $src -Destination $dest -Force }
        elseif ([System.IO.Path]::GetFullPath($src) -ne [System.IO.Path]::GetFullPath($dest)) { Move-Item -LiteralPath $src -Destination $dest -Force }
    } catch { return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message; path = $dest; name = $n['name'] } }
    return @{ ok = $true; path = $dest; name = $n['name']; warnings = $warnings }
}
