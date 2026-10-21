# modules/excel/excel.export_pictures.ps1
# Export every picture of a sheet to PNG files, top-to-bottom, left-to-
# right (10/21: the HOST screen shots on the send-data sheet, so their line
# numbers can be read and the DF pictures given the same ranges). Read
# only: the sheet is left as it was (a temporary chart per picture,
# deleted again). Clobbers the clipboard.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.export_pictures'
  group      = 'excel'
  summary    = 'Export every picture of a sheet to PNG, top-to-bottom'
  tier       = 'core'
  effects    = 'read'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet    = @{ type='string';  required=$true }
    saveDir  = @{ type='path';    required=$true; desc='pic1.png, pic2.png, ... in sheet order; relative paths resolve under the work dir' }
    scale    = @{ type='int';     default=3; desc='export size against the size on the sheet (bigger reads better for OCR)' }
  }
  outputs    = @{
    paths  = @{ type='list'; desc='the exported PNGs, top-to-bottom (a picture that would not export is left out, with a warning)' }
    pictures = @{ type='int';  desc='pictures on the sheet' }
    groups = @{ type='list'; desc='names of grouped shapes (Ctrl+G) on the sheet: their pictures are not exported' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
  )
  example    = @{ use = 'excel.export_pictures'; with = @{ workbook = 'wb'; sheet = 'sendData'; saveDir = 'capture/{{item.keySafe}}/host' } }
  notes      = 'Old PNGs in saveDir named pic<n>.png are removed first, so a sheet with fewer pictures than last time leaves none behind.'
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['saveDir']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would export the pictures of {0} to {1}' -f $In['sheet'], $dir)); return @{ ok = $true; paths = @(); pictures = 0; groups = @() } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('sheet "' + $In['sheet'] + '" not found'); paths = @(); pictures = 0; groups = @() } }
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    foreach ($old in @(Get-ChildItem -LiteralPath $dir -Filter 'pic*.png' -File -ErrorAction SilentlyContinue)) { Remove-Item -LiteralPath $old.FullName -Force -ErrorAction SilentlyContinue }
    $pics = @(Get-EbiPictureShapes -Worksheet $ws)
    $groups = @(Get-EbiGroupShapesAbove -Worksheet $ws -Top 1.0e9)
    $warn = New-Object System.Collections.ArrayList
    if ($groups.Count -gt 0) { [void]$warn.Add(@{ code = 'group_shape'; message = ('{0} has grouped shape(s) (Ctrl+G: {1}): pictures inside a group are not exported' -f $In['sheet'], ($groups -join ', ')) }) }
    $paths = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($p in $pics) {
        $n++
        $dest = Join-Path $dir ('pic' + $n + '.png')
        $x = Export-EbiShapePng -Worksheet $ws -Shape $p['shape'] -OutPath $dest -Scale ([double]$In['scale'])
        if ($x['ok']) { [void]$paths.Add($dest) } else { [void]$warn.Add(@{ code = 'png_not_saved'; message = ('picture #' + $n + ' (' + $p['name'] + ') not exported: ' + $x['message']) }) }
    }
    try { $In['workbook'].Application.CutCopyMode = $false } catch { }
    $Ctx.Log.Info(('exported {0} of {1} picture(s) of {2}' -f $paths.Count, $pics.Count, $In['sheet']))
    return @{ ok = $true; paths = $paths.ToArray(); pictures = $pics.Count; groups = $groups; warnings = $warn.ToArray() }
}
