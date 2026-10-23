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
#   @{ textPanel = @{ fromRight; toRight; fromBottom; toBottom }; x; y;
#      w; h; anchor }         the cell around the LAST text in that
#                             bottom-right strip (DF's "same content" cell, 10/21),
#                             kernel/Layout.ps1 Find-EbiTextPanel; the
#                             fixed x/y/w/h/anchor box when none is found
#   + template = 'profiles/<p>/assets/x.png' (optional): that crop searched
#     for in the same strip first (the operator's own picture of the cell)
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
    pictures = @{ type='list';   required=$true; desc='paths, or @{ path; column; gapRows; rows; rects; role } (excel layout plan)' }
    gapRows  = @{ type='int';    default=1; desc='blank rows between pictures when an entry does not say' }
    rects    = @{ type='list';   default=@(); desc='boxes for EVERY picture given as a plain path' }
    lineWeight = @{ type='int';  default=0; desc='box line weight in points; 0 = 1.5' }
    tag      = @{ type='string'; default=''; desc='shape name prefix (verifyMark_<tag>_n) so a rerun can find them' }
    replace  = @{ type='bool';   default=$false; desc='with tag: first delete the pictures and boxes an earlier run left on this sheet under the same tag (verifyPic_<tag>_* / verifyMark_<tag>_*), so a rerun replaces instead of stacking a second copy on top' }
  }
  outputs    = @{
    shapes  = @{ type='list'; desc='@{ name; path; row; top; left; width; height; boxes; role } (role from an excel.stack_plan entry: first / separator / last / only; blank for a plain path)' }
    nextRow = @{ type='int';  desc='first row below the last picture' }
    boxes   = @{ type='int';  desc='rectangles drawn' }
    removed = @{ type='int';  desc='shapes of an earlier run deleted first (replace)' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'file_not_found';  transient = $true  }
    @{ id = 'insert_failed';   transient = $true  }
  )
  example    = @{ use = 'excel.insert_pictures'; with = @{ workbook = 'wb'; sheet = 'result'; row = '{{steps.log.out.nextRow}}'; pictures = @('{{steps.shot.out.path}}'); rects = '{{steps.lr.out.rects}}' } }
  notes      = 'Not idempotent without replace: a second run adds the pictures again. With tag + replace = true a rerun deletes its own earlier pictures and boxes (by name) first -- pictures someone put there by hand are never touched. Excel is scrolled to each picture after it is drawn and held there a moment (EBI_WATCH_MS).'
}

function ExcelInsertPictures-Entry {
    param($P, [string]$DefaultColumn, [int]$DefaultGap, $DefaultRects)
    if ($P -is [System.Collections.IDictionary]) {
        return @{ path = [string]$P['path']; column = $(if ($P.Contains('column') -and [string]$P['column'] -ne '') { [string]$P['column'] } else { $DefaultColumn });
                  gapRows = $(if ($P.Contains('gapRows') -and $null -ne $P['gapRows']) { [int]$P['gapRows'] } else { $DefaultGap });
                  rows = $(if ($P.Contains('rows') -and $null -ne $P['rows']) { [int]$P['rows'] } else { 0 });
                  rects = @(if ($P.Contains('rects') -and $null -ne $P['rects']) { $P['rects'] });
                  role = $(if ($P.Contains('role') -and $null -ne $P['role']) { [string]$P['role'] } else { '' }) }
    }
    return @{ path = [string]$P; column = $DefaultColumn; gapRows = $DefaultGap; rows = 0; rects = @($DefaultRects); role = '' }
}

function ExcelInsertPictures-IsOwnShape {
    # PURE. A shape name -> true when an earlier run of this step drew it
    # under this tag: verifyPic_<tag>_<n> or verifyMark_<tag>_<n>_<k>.
    param([string]$Name, [string]$Tag)
    if ([string]::IsNullOrWhiteSpace($Tag)) { return $false }
    $t = [regex]::Escape($Tag)
    return ($Name -match ('^verifyPic_' + $t + '_\d+$') -or $Name -match ('^verifyMark_' + $t + '_\d+_\d+$'))
}

function ExcelInsertPictures-FindTemplate {
    # A template PNG (relative = under the tool folder) searched for inside
    # the strip the spec's textPanel names (default: 260 px from the right,
    # 22 from the bottom) -> image px rect of the best place, or $null when
    # it is not there (mean grey difference above templateMaxDiff, 40).
    param($Spec, [string]$Path, [int]$W, [int]$H)
    $tp = [string]$Spec['template']
    if (-not [System.IO.Path]::IsPathRooted($tp)) { $tp = Join-Path (Join-Path $PSScriptRoot '..\..') $tp }
    if (-not (Test-Path -LiteralPath $tp -PathType Leaf)) { return $null }
    $t = if ($Spec.Contains('textPanel') -and ($Spec['textPanel'] -is [System.Collections.IDictionary])) { $Spec['textPanel'] } else { @{} }
    $fr = if ($t.Contains('fromRight')) { [int]$t['fromRight'] } else { 260 }
    $fb = if ($t.Contains('fromBottom')) { [int]$t['fromBottom'] } else { 22 }
    $tg = Get-EbiGrayRegion -Path $tp -X 0 -Y 0 -W 100000 -H 100000
    $rg = Get-EbiGrayRegion -Path $Path -X ($W - $fr) -Y ($H - $fb) -W $fr -H $fb
    if (-not $tg['ok'] -or -not $rg['ok'] -or $tg['w'] -gt $rg['w'] -or $tg['h'] -gt $rg['h']) { return $null }
    $cx = [int](($rg['w'] - $tg['w']) / 2); $cy = [int](($rg['h'] - $tg['h']) / 2)
    $m = [EbiImageMath]::BestDiff($tg['g'], $tg['w'], $tg['h'], $rg['g'], $rg['w'], $rg['h'], $cx, $cy, $cx + 1, $cy + 1)
    $max = if ($Spec.Contains('templateMaxDiff')) { [double]$Spec['templateMaxDiff'] } else { 40.0 }
    if ([double]$m[0] -lt 0 -or [double]$m[0] -gt $max) { return $null }
    return @{ x = [double]([int]$rg['x0'] + [int]$m[1]); y = [double]([int]$rg['y0'] + [int]$m[2]); w = [double]$tg['w']; h = [double]$tg['h'] }
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
    if ($Spec.Contains('template') -and -not [string]::IsNullOrWhiteSpace([string]$Spec['template'])) {
        # the operator's own crop of the cell (10/21: "use the pic I provide"),
        # searched for in the bottom-right strip; textPanel / the fixed box after it
        $r = ExcelInsertPictures-FindTemplate -Spec $Spec -Path $Path -W $W -H $H
        if ($null -ne $r) { return $r }
    }
    if ($Spec.Contains('textPanel') -and ($Spec['textPanel'] -is [System.Collections.IDictionary])) {
        # the cell around the last text in a strip near the bottom-right (DF's
        # "same content" cell), found in the picture; the fixed box below it is the fallback
        $t = $Spec['textPanel']
        $fr = if ($t.Contains('fromRight')) { [int]$t['fromRight'] } else { 260 }
        $tr = if ($t.Contains('toRight')) { [int]$t['toRight'] } else { 16 }
        $fb = if ($t.Contains('fromBottom')) { [int]$t['fromBottom'] } else { 22 }
        $tb = if ($t.Contains('toBottom')) { [int]$t['toBottom'] } else { 0 }
        $g = Get-EbiGrayRegion -Path $Path -X ($W - $fr) -Y ($H - $fb) -W ($fr - $tr) -H ($fb - $tb)
        if ($g['ok']) {
            $p = Find-EbiTextPanel -Gray $g['g'] -W $g['w'] -H $g['h'] -TextMaxLuma $(if ($t.Contains('textMaxLuma')) { [int]$t['textMaxLuma'] } else { 100 }) -Gap $(if ($t.Contains('gap')) { [int]$t['gap'] } else { 10 })
            if ($p['ok']) { return @{ x = [double]([int]$g['x0'] + [int]$p['left']); y = [double]([int]$g['y0'] + [int]$p['top']); w = [double]([int]$p['right'] - [int]$p['left'] + 1); h = [double]([int]$p['bottom'] - [int]$p['top'] + 1) } }
        }
        if (-not $Spec.Contains('anchor') -and -not $Spec.Contains('x')) { return $null }
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
        return @{ ok = $true; shapes = @(); nextRow = [int]$In['row']; boxes = 0; removed = 0 }
    }
    foreach ($e in $entries) { if (-not (Test-Path -LiteralPath $e['path'] -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $e['path']; shapes = @(); nextRow = [int]$In['row']; boxes = 0; removed = 0 } } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('no sheet "' + $In['sheet'] + '"'); shapes = @(); nextRow = [int]$In['row']; boxes = 0; removed = 0 } }
    $weight = if ([int]$In['lineWeight'] -gt 0) { [double]$In['lineWeight'] } else { 1.5 }
    $tag = [string]$In['tag']
    $row = [Math]::Max(1, [int]$In['row'] + [Math]::Max(0, [int]$In['skipRows']))
    # 10/19: a rerun on a finished book rewrote the logs but stacked a second
    # Jenkins / DF picture on top of the first -- delete our own first
    $removed = 0
    if ([bool]$In['replace'] -and $tag -ne '') {
        $old = New-Object System.Collections.ArrayList
        try { foreach ($sh in @($ws.Shapes)) { if (ExcelInsertPictures-IsOwnShape -Name ([string]$sh.Name) -Tag $tag) { [void]$old.Add($sh) } } } catch { }
        foreach ($sh in $old) { try { [void]$sh.Delete(); $removed++ } catch { } }
        if ($removed -gt 0) { $Ctx.Log.Info(('{0}: deleted {1} picture(s)/box(es) an earlier run left under tag "{2}"' -f $In['sheet'], $removed, $tag)) }
    }
    $shapes = New-Object System.Collections.ArrayList
    $boxes = 0; $n = 0
    $warn = New-Object System.Collections.ArrayList
    foreach ($e in $entries) {
        $n++
        if ($n -gt 1) { $row = $row + [Math]::Max(0, [int]$e['gapRows']) }
        Show-EbiExcelAt -Worksheet $ws -Row $row -Front
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
                if ($null -eq $ir) { [void]$warn.Add(@{ code = 'box_not_placed'; message = ('no ink found for a box on ' + $e['path'] + ' (a listing with no lines has no last line to box -- expected for a 0-byte transfer)') }); continue }
                $sr = ConvertTo-EbiSheetRect -PicLeft $pl -PicTop $ptop -PicWidth $pw -PicHeight $ph -ImageWidth $iw -ImageHeight $ih -Rect $ir
                $k++
                $nm = if ($tag -ne '') { ('verifyMark_{0}_{1}_{2}' -f $tag, $n, $k) } else { '' }
                [void](Add-EbiRedRect -Worksheet $ws -Left $sr['left'] -Top $sr['top'] -Width $sr['width'] -Height $sr['height'] -Name $nm -Weight $weight -AltText $(if ($tag -ne '') { 'verifyMark|' + $tag + '|' + $n + '|' + $k } else { '' }))
                $boxes++
            }
            $bottom = $ptop + $ph
            $next = Get-EbiRowAtOrBelow -Worksheet $ws -TopPoints $bottom -StartRow $row
            if ([int]$e['rows'] -gt 0) { $next = [Math]::Max($next, $row + [int]$e['rows']) }
            [void]$shapes.Add(@{ name = [string]$pic.Name; path = $e['path']; row = $row; top = $ptop; left = $pl; width = $pw; height = $ph; boxes = $k; role = [string]$e['role'] })
            Show-EbiExcelAt -Worksheet $ws -Row $row -Front -Dwell
            $row = $next
        } catch { return @{ ok = $false; failure = 'insert_failed'; message = ($e['path'] + ': ' + $_.Exception.Message); shapes = $shapes.ToArray(); nextRow = $row; boxes = $boxes; removed = $removed } }
    }
    $Ctx.Log.Info(('inserted {0} picture(s), {1} box(es) on {2}; next free row {3}' -f $shapes.Count, $boxes, $In['sheet'], $row))
    return @{ ok = $true; shapes = $shapes.ToArray(); nextRow = $row; boxes = $boxes; removed = $removed; warnings = $warn.ToArray() }
}
