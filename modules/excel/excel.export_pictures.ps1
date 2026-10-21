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
  summary    = 'Export every picture (and picture group / embedded object) of a sheet to PNG, top-to-bottom'
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
    hideMarks = @{ type='bool';   default=$true; desc='in a group, hide the shapes that are not pictures (red boxes) while it is exported, then show them again' }
    scale    = @{ type='int';     default=3; desc='export size against the size on the sheet (bigger reads better for OCR)' }
  }
  outputs    = @{
    paths  = @{ type='list'; desc='the exported PNGs, top-to-bottom (a picture that would not export is left out, with a warning)' }
    pictures = @{ type='int';  desc='pictures on the sheet' }
    groups = @{ type='list'; desc='names of grouped shapes (Ctrl+G) exported as one image each' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
  )
  example    = @{ use = 'excel.export_pictures'; with = @{ workbook = 'wb'; sheet = 'sendData'; saveDir = 'capture/{{item.keySafe}}/host' } }
  notes      = 'Old PNGs in saveDir named pic<n>.png are removed first, so a sheet with fewer pictures than last time leaves none behind.'
}

function ExcelExportPictures-Shapes {
    <#
      The shapes to export, sorted top, then left: pictures (13), linked
      pictures (11), embedded / linked OLE objects (7 / 10), canvases (20)
      and GROUPS (6) as one image each -- 10/21 KJODWWBV: the HOST shots on
      the send-data sheet were not plain pictures, and the first cut found
      none. Also @{ type; name } of everything else, for the warning.
    #>
    param($Worksheet)
    $take = New-Object System.Collections.ArrayList
    $other = New-Object System.Collections.ArrayList
    try {
        foreach ($sh in $Worksheet.Shapes) {
            $t = -1
            try { $t = [int]$sh.Type } catch { }
            $nm = ''
            try { $nm = [string]$sh.Name } catch { }
            if (@(6, 7, 10, 11, 13, 20) -contains $t) {
                [void]$take.Add(@{ shape = $sh; name = $nm; type = $t; top = [double]$sh.Top; left = [double]$sh.Left })
            } else { [void]$other.Add(('{0} (type {1})' -f $nm, $t)) }
        }
    } catch { }
    $sorted = @($take | Sort-Object -Property @{ Expression = { $_['top'] } }, @{ Expression = { $_['left'] } })
    return @{ shapes = $sorted; other = $other.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    $dir = Resolve-EbiWorkPath -PathValue ([string]$In['saveDir']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would export the pictures of {0} to {1}' -f $In['sheet'], $dir)); return @{ ok = $true; paths = @(); pictures = 0; groups = @() } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('sheet "' + $In['sheet'] + '" not found'); paths = @(); pictures = 0; groups = @() } }
    if (-not (Test-Path -LiteralPath $dir)) { [void](New-Item -ItemType Directory -Path $dir -Force) }
    foreach ($old in @(Get-ChildItem -LiteralPath $dir -Filter 'pic*.png' -File -ErrorAction SilentlyContinue)) { Remove-Item -LiteralPath $old.FullName -Force -ErrorAction SilentlyContinue }
    $found = ExcelExportPictures-Shapes -Worksheet $ws
    $pics = @($found['shapes'])
    $groups = @($pics | Where-Object { [int]$_['type'] -eq 6 } | ForEach-Object { [string]$_['name'] })
    $warn = New-Object System.Collections.ArrayList
    if ($pics.Count -eq 0) {
        $msg = ('{0}: no picture to export; its shapes: {1}' -f $In['sheet'], $(if (@($found['other']).Count) { (@($found['other']) | Select-Object -First 12) -join ', ' } else { 'none' }))
        [void]$warn.Add(@{ code = 'no_picture'; message = $msg })
        $Ctx.Log.Warn($msg)
    }
    $paths = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($p in $pics) {
        $n++
        $dest = Join-Path $dir ('pic' + $n + '.png')
        # a group: hide its marks (red boxes, arrows -- not pictures) for the
        # export, so a box drawn over a line number does not hide it from OCR
        # (10/21 KJODWWBV: 007568 under a red box); shown again right after
        $hidden = New-Object System.Collections.ArrayList
        if ([int]$p['type'] -eq 6 -and [bool]$In['hideMarks']) {
            try {
                foreach ($gi in $p['shape'].GroupItems) {
                    $gt = -1; try { $gt = [int]$gi.Type } catch { }
                    if (@(11, 13) -notcontains $gt) { try { if ([int]$gi.Visible -ne 0) { $gi.Visible = 0; [void]$hidden.Add($gi) } } catch { } }
                }
            } catch { }
        }
        try { $x = Export-EbiShapePng -Worksheet $ws -Shape $p['shape'] -OutPath $dest -Scale ([double]$In['scale']) }
        finally { foreach ($gi in $hidden) { try { $gi.Visible = -1 } catch { } } }
        if ($x['ok']) { [void]$paths.Add($dest) } else { [void]$warn.Add(@{ code = 'png_not_saved'; message = ('picture #' + $n + ' (' + $p['name'] + ') not exported: ' + $x['message']) }) }
    }
    try { $In['workbook'].Application.CutCopyMode = $false } catch { }
    $Ctx.Log.Info(('exported {0} of {1} picture(s) of {2}' -f $paths.Count, $pics.Count, $In['sheet']))
    return @{ ok = $true; paths = $paths.ToArray(); pictures = $pics.Count; groups = $groups; warnings = $warn.ToArray() }
}
