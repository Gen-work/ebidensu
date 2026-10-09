# modules/excel/excel.insert_pictures.ps1
# Stack pictures down a sheet, each starting on the first row below the
# previous one (+ its gap), at native size, and draw red rectangles on
# them at IMAGE pixel positions (Excel's scaling is taken from the shape,
# kernel/Layout.ps1 ConvertTo-EbiSheetRect). A rectangle is one of:
#   @{ x; y; w; h; anchor }   fixed, relative to a corner (tl tr bl br) --
#                             a status-bar cell stays put when the window
#                             is resized
#   @{ lastInk = @{ x; width; yFrom; yTo; ink; maxLuma }; x; w; padTop;
#      padBottom }            on the LAST band of ink in that strip of the
#                             picture (the last line of a listing): found
#                             in the picture itself, kernel/Image.ps1
# Ported from ExcelHelpers.ps1 Insert-PictureSendToBack + Add-RedRectangle.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')    # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # ConvertTo-EbiColumnNumber

$Manifest = @{
  id         = 'excel.insert_pictures'
  group      = 'excel'
  summary    = 'Stack pictures down a sheet at native size and draw red boxes on them'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    workbook = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet    = @{ type='string'; required=$true }
    row      = @{ type='int';    required=$true; desc='row of the first picture' }
    skipRows = @{ type='int';    default=0; desc='blank rows left before the first picture' }
    column   = @{ type='string'; default='B'; desc='default column' }
    pictures = @{ type='list';   required=$true; desc='paths, or @{ path; column; gapRows; rows; rects } (excel layout plan)' }
    gapRows  = @{ type='int';    default=1; desc='blank rows between pictures when an entry does not say' }
    rects    = @{ type='list';   default=@(); desc='boxes for EVERY picture given as a plain path' }
    lineWeight = @{ type='int';  default=0; desc='box line weight in points; 0 = 1.5' }
    tag      = @{ type='string'; default=''; desc='shape name prefix (verifyMark_<tag>_n) so a rerun can find them' }
  }
  outputs    = @{
    shapes  = @{ type='list'; desc='@{ name; path; row; top; left; width; height; boxes }' }
    nextRow = @{ type='int';  desc='first row below the last picture' }
    boxes   = @{ type='int';  desc='rectangles drawn' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'file_not_found';  transient = $true  }
    @{ id = 'insert_failed';   transient = $true  }
  )
  example    = @{ use = 'excel.insert_pictures'; with = @{ workbook = 'wb'; sheet = 'result'; row = '{{steps.log.out.nextRow}}'; pictures = @('{{steps.shot.out.path}}'); rects = '{{steps.lr.out.rects}}' } }
  notes      = 'Not idempotent: a second run adds the pictures again. The evidence workflows start from a clean sheet section and checkpoint after.'
}

function ExcelInsertPictures-Entry {
    param($P, [string]$DefaultColumn, [int]$DefaultGap, $DefaultRects)
    if ($P -is [System.Collections.IDictionary]) {
        return @{ path = [string]$P['path']; column = $(if ($P.Contains('column') -and [string]$P['column'] -ne '') { [string]$P['column'] } else { $DefaultColumn });
                  gapRows = $(if ($P.Contains('gapRows') -and $null -ne $P['gapRows']) { [int]$P['gapRows'] } else { $DefaultGap });
                  rows = $(if ($P.Contains('rows') -and $null -ne $P['rows']) { [int]$P['rows'] } else { 0 });
                  rects = @(if ($P.Contains('rects') -and $null -ne $P['rects']) { $P['rects'] }) }
    }
    return @{ path = [string]$P; column = $DefaultColumn; gapRows = $DefaultGap; rows = 0; rects = @($DefaultRects) }
}

function ExcelInsertPictures-ImageRect {
    # One rect spec -> image px @{ x; y; w; h } or $null (ink not found).
    param($Spec, [string]$Path, [int]$W, [int]$H)
    if ($Spec.Contains('lastInk') -and ($Spec['lastInk'] -is [System.Collections.IDictionary])) {
        $s = $Spec['lastInk']
        $ink = if ($s.Contains('ink')) { [string]$s['ink'] } else { 'dark' }
        $ml = if ($s.Contains('maxLuma')) { [int]$s['maxLuma'] } else { 160 }
        $c = Get-EbiInkCounts -Path $Path -X ([int]$s['x']) -Width ([int]$s['width']) -YFrom $(if ($s.Contains('yFrom')) { [int]$s['yFrom'] } else { 0 }) -YTo $(if ($s.Contains('yTo')) { [int]$s['yTo'] } else { 0 }) -Ink $ink -MaxLuma $ml
        if (-not $c['ok']) { return $null }
        $minH = if ($s.Contains('minHeight')) { [int]$s['minHeight'] } else { 4 }
        $bands = @(Get-EbiInkBands -Counts ([int[]]$c['counts']) -MinInk $(if ($s.Contains('minInk')) { [int]$s['minInk'] } else { 2 }) -MergeGap $(if ($s.Contains('mergeGap')) { [int]$s['mergeGap'] } else { 1 }) | Where-Object { [int]$_['height'] -ge $minH })
        if ($bands.Count -eq 0) { return $null }
        $b = $bands[$bands.Count - 1]
        $pt = if ($Spec.Contains('padTop')) { [int]$Spec['padTop'] } else { 3 }
        $pb = if ($Spec.Contains('padBottom')) { [int]$Spec['padBottom'] } else { 3 }
        $top = [int]$c['yFrom'] + [int]$b['top'] - $pt
        return @{ x = [double]$Spec['x']; y = [double]$top; w = [double]$Spec['w']; h = [double]([int]$b['height'] + $pt + $pb) }
    }
    return (Resolve-EbiAnchoredRect -Rect $Spec -ImageWidth $W -ImageHeight $H)
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $entries = @(foreach ($p in @($In['pictures'])) { if ($null -ne $p) { ExcelInsertPictures-Entry -P $p -DefaultColumn ([string]$In['column']) -DefaultGap ([int]$In['gapRows']) -DefaultRects @($In['rects']) } })
    $entries = @($entries | Where-Object { -not [string]::IsNullOrWhiteSpace($_['path']) })
    foreach ($e in $entries) { $e['path'] = Resolve-EbiWorkPath -PathValue $e['path'] -WorkDir $work }
    if ($Ctx['DryRun']) {
        foreach ($e in $entries) { $Ctx.Log.Info(('would insert {0} on {1} at column {2} with {3} box(es)' -f $e['path'], $In['sheet'], $e['column'], @($e['rects']).Count)) }
        if ($entries.Count -eq 0) { $Ctx.Log.Info('would insert no picture') }
        return @{ ok = $true; shapes = @(); nextRow = [int]$In['row']; boxes = 0 }
    }
    foreach ($e in $entries) { if (-not (Test-Path -LiteralPath $e['path'] -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $e['path']; shapes = @(); nextRow = [int]$In['row']; boxes = 0 } } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); shapes = @(); nextRow = [int]$In['row']; boxes = 0 } }
    $weight = if ([int]$In['lineWeight'] -gt 0) { [double]$In['lineWeight'] } else { 1.5 }
    $tag = [string]$In['tag']
    $row = [Math]::Max(1, [int]$In['row'] + [Math]::Max(0, [int]$In['skipRows']))
    $shapes = New-Object System.Collections.ArrayList
    $boxes = 0; $n = 0
    $warn = New-Object System.Collections.ArrayList
    foreach ($e in $entries) {
        $n++
        if ($n -gt 1) { $row = $row + [Math]::Max(0, [int]$e['gapRows']) }
        try {
            $col = ConvertTo-EbiColumnNumber $e['column']
            $cell = $ws.Cells.Item($row, $col)
            $left = [double]$cell.Left; $top = [double]$cell.Top
            $pic = $ws.Shapes.AddPicture($e['path'], 0, -1, $left, $top, -1, -1)
            try { [void]$pic.ZOrder(1) } catch { }
            if ($tag -ne '') { try { $pic.Name = ('verifyPic_{0}_{1}' -f $tag, $n) } catch { } }
            $size = Get-EbiPngSize -Path $e['path']
            $iw = if ($size['ok']) { [int]$size['width'] } else { 0 }; $ih = if ($size['ok']) { [int]$size['height'] } else { 0 }
            $pl = [double]$pic.Left; $ptop = [double]$pic.Top; $pw = [double]$pic.Width; $ph = [double]$pic.Height
            $k = 0
            foreach ($spec in @($e['rects'])) {
                if (-not ($spec -is [System.Collections.IDictionary])) { continue }
                $ir = ExcelInsertPictures-ImageRect -Spec $spec -Path $e['path'] -W $iw -H $ih
                if ($null -eq $ir) { [void]$warn.Add(@{ code = 'box_not_placed'; message = ('no ink found for a box on ' + $e['path']) }); continue }
                $sr = ConvertTo-EbiSheetRect -PicLeft $pl -PicTop $ptop -PicWidth $pw -PicHeight $ph -ImageWidth $iw -ImageHeight $ih -Rect $ir
                $k++
                $nm = if ($tag -ne '') { ('verifyMark_{0}_{1}_{2}' -f $tag, $n, $k) } else { '' }
                [void](Add-EbiRedRect -Worksheet $ws -Left $sr['left'] -Top $sr['top'] -Width $sr['width'] -Height $sr['height'] -Name $nm -Weight $weight -AltText $(if ($tag -ne '') { 'verifyMark|' + $tag + '|' + $n + '|' + $k } else { '' }))
                $boxes++
            }
            $bottom = $ptop + $ph
            $next = Get-EbiRowAtOrBelow -Worksheet $ws -TopPoints $bottom -StartRow $row
            if ([int]$e['rows'] -gt 0) { $next = [Math]::Max($next, $row + [int]$e['rows']) }
            [void]$shapes.Add(@{ name = [string]$pic.Name; path = $e['path']; row = $row; top = $ptop; left = $pl; width = $pw; height = $ph; boxes = $k })
            $row = $next
        } catch { return @{ ok = $false; failure = 'insert_failed'; message = ($e['path'] + ': ' + $_.Exception.Message); shapes = $shapes.ToArray(); nextRow = $row; boxes = $boxes } }
    }
    $Ctx.Log.Info(('inserted {0} picture(s), {1} box(es) on {2}; next free row {3}' -f $shapes.Count, $boxes, $In['sheet'], $row))
    return @{ ok = $true; shapes = $shapes.ToArray(); nextRow = $row; boxes = $boxes; warnings = $warn.ToArray() }
}
