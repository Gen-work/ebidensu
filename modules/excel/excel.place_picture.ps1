# modules/excel/excel.place_picture.ps1
# Put one PNG on a sheet at native size, optionally REPLACING a picture
# that is already there: 'atCell' takes away the pictures whose top-left
# corner sits in the target cell (a rerun replaces its own copy), 'first'
# takes away the sheet's first picture (top-to-bottom, left-to-right) and
# puts the new one where it was. 10/16: evidence snaps picture 1 from the
# overview list (excel.snap_list) and puts it on GFIX's result sheet, and
# on GIFT's in place of the old picture 1 when the snap had to run on to
# the source-file column.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.place_picture'
  group      = 'excel'
  summary    = 'Put a PNG on a sheet at a cell, replacing the picture there (or the first picture) when asked'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    workbook = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet    = @{ type='string'; required=$true }
    path     = @{ type='path';   required=$true; desc='the PNG (relative: under the work dir)' }
    cell     = @{ type='string'; default='B3'; desc='top-left cell (replace = first: used only when the sheet has no picture)' }
    replace  = @{ type='string'; default='atCell'; enum=@('none', 'atCell', 'first'); desc='none = just add; atCell = remove pictures whose top-left is in cell first; first = remove the sheet''s first picture and put the new one where it was' }
    gapRows  = @{ type='int';    default=0; desc='added to nextRow' }
    name     = @{ type='string'; default=''; desc='shape name for the new picture (empty: Excel''s)' }
  }
  outputs    = @{
    name     = @{ type='string' }
    nextRow  = @{ type='int';    desc='first row whose top is below the picture, plus gapRows' }
    replaced = @{ type='list';   desc='names of the pictures taken away' }
    groups   = @{ type='list';   desc='replace = first: grouped shapes (Ctrl+G) at or above the replaced picture -- the "first picture" may not be the one a person sees first' }
    note     = @{ type='string'; desc='one line for the gate when groups is not empty, else empty' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'file_not_found';  transient = $true  }
    @{ id = 'insert_failed';   transient = $true  }
  )
  example    = @{ use = 'excel.place_picture'; with = @{ workbook = 'wb'; sheet = 'result'; path = '{{steps.snap.out.png}}'; cell = 'B3'; replace = 'atCell' } }
  notes      = 'Not idempotent for replace = none (a rerun adds a second copy); atCell and first replace what is there, so a rerun ends with one picture.'
}

function ExcelPlacePicture-InCell {
    # PURE. Is a shape's top-left corner (points) inside the cell's box?
    param([double]$ShapeLeft, [double]$ShapeTop, [double]$CellLeft, [double]$CellTop, [double]$CellWidth, [double]$CellHeight)
    return ($ShapeLeft -ge ($CellLeft - 0.5) -and $ShapeLeft -lt ($CellLeft + $CellWidth - 0.5) -and $ShapeTop -ge ($CellTop - 0.5) -and $ShapeTop -lt ($CellTop + $CellHeight - 0.5))
}

function Invoke-Step {
    param($In, $Ctx)
    $none = @{ name = ''; nextRow = 0; replaced = @(); groups = @(); note = '' }
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would put {0} on {1}!{2} (replace: {3})' -f $path, $In['sheet'], $In['cell'], $In['replace'])); $r = $none.Clone(); $r['ok'] = $true; return $r }
    if ([string]::IsNullOrWhiteSpace($path) -or -not (Test-Path -LiteralPath $path -PathType Leaf)) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'file_not_found'; $r['message'] = ('no picture at ' + $path); return $r }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'sheet_not_found'; $r['message'] = ('sheet "' + $In['sheet'] + '" not found'); return $r }
    try {
        $target = $ws.Range([string]$In['cell'])
        $left = [double]$target.Left; $top = [double]$target.Top
        $gone = New-Object System.Collections.ArrayList
        $groups = @(); $note = ''
        $pics = @(Get-EbiPictureShapes -Worksheet $ws)
        if ([string]$In['replace'] -eq 'first' -and $pics.Count -gt 0) {
            $p0 = $pics[0]
            $left = [double]$p0['left']; $top = [double]$p0['top']
            $groups = @(Get-EbiGroupShapesAbove -Worksheet $ws -Top $top)
            if ($groups.Count -gt 0) {
                $note = ('{0} has {1} grouped shape(s) (Ctrl+G: {2}) at or above its first picture {3}: the replaced one may not be the picture you see first -- check it' -f $In['sheet'], $groups.Count, ($groups -join ', '), $p0['name'])
                $Ctx.Log.Warn($note)
            }
            [void]$gone.Add([string]$p0['name'])
            [void]$p0['shape'].Delete()
        } elseif ([string]$In['replace'] -eq 'atCell') {
            foreach ($p in $pics) {
                if (ExcelPlacePicture-InCell -ShapeLeft ([double]$p['left']) -ShapeTop ([double]$p['top']) -CellLeft $left -CellTop $top -CellWidth ([double]$target.Width) -CellHeight ([double]$target.Height)) {
                    [void]$gone.Add([string]$p['name'])
                    [void]$p['shape'].Delete()
                }
            }
        }
        $pic = $ws.Shapes.AddPicture($path, 0, -1, $left, $top, -1, -1)
        if ([string]$In['name'] -ne '') { try { $pic.Name = [string]$In['name'] } catch { } }
        try { [void]$pic.ZOrder(1) } catch { }   # behind any red box drawn later
        $startRow = 1
        try { $startRow = [int]$ws.Cells.Item(1, 1).Row; $startRow = [int]$pic.TopLeftCell.Row } catch { }
        $next = Get-EbiRowAtOrBelow -Worksheet $ws -TopPoints ([double]$pic.Top + [double]$pic.Height) -StartRow $startRow
        if ($gone.Count -gt 0) { $Ctx.Log.Info(('{0}: replaced {1} with {2}' -f $In['sheet'], ($gone -join ', '), (Split-Path -Leaf $path))) }
        Show-EbiExcelAt -Worksheet $ws -Row $startRow -Front
        $out = @{ ok = $true; name = [string]$pic.Name; nextRow = ($next + [int]$In['gapRows']); replaced = $gone.ToArray(); groups = $groups; note = $note }
        if ($note -ne '') { $out['warnings'] = @(@{ code = 'group_shape'; message = $note }) }
        return $out
    } catch {
        $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'insert_failed'; $r['message'] = (Format-EbiExcelError -Stage 'place picture' -ErrorRecord $_); return $r
    }
}
