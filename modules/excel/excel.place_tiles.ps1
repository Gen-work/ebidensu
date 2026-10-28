# modules/excel/excel.place_tiles.ps1
# Put the page captures of a long line (screen.stitch_scroll tiles +
# layout) on a sheet side by side, lined up column for column, and group
# them (Ctrl+G) -- the long snap merged INSIDE Excel, the way the J4
# evidence books hold it: every capture at native size, each one after the
# first moved right by its first column x px a column and cropped by its
# line-number strip (the books' later captures carry a 1.5 % left crop),
# so the text runs on across the seams. The pictures stay separate
# originals inside the group (Excel never resamples them), unlike
# the merged PNG -- which goes in through this step too (one picture, no
# layout), so the two can be compared on one sheet and a rerun replaces
# both by name.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.place_tiles'
  group      = 'excel'
  summary    = 'Put page captures of a long line side by side on a sheet, lined up by column, cropped and grouped'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    workbook   = @{ type='session'; sessionKind='workbook'; required=$true }
    sheet      = @{ type='string'; required=$true }
    cell       = @{ type='string'; default='B4'; desc='top-left cell of the first capture' }
    row        = @{ type='int';    default=0; desc='> 0: that row of cell''s column instead (excel.place_picture nextRow: under the merged picture)' }
    tiles      = @{ type='list';   required=$true; desc='the captures, left to right (screen.stitch_scroll tiles)' }
    layout     = @{ type='map';    default=@{}; desc='screen.stitch_scroll layout: starts, charPx, textLeft, tileW; empty = the pictures one after another, uncropped (the merged PNG: one picture)' }
    cropStrip  = @{ type='bool';   default=$true; desc='crop the line-number strip (textLeft px) off every capture after the first' }
    group      = @{ type='bool';   default=$true; desc='group the pictures into one shape (Ctrl+G)' }
    name       = @{ type='string'; default='verifyWide'; desc='name of the group (pictures: <name>_1 ..; a single picture is called name itself); an earlier run''s shapes of that name are replaced' }
    gapRows    = @{ type='int';    default=0; desc='added to nextRow' }
  }
  outputs    = @{
    name     = @{ type='string'; desc='the group (or the single picture; group = false: the first picture)' }
    shapes   = @{ type='list';   desc='picture names, left to right' }
    widthPt  = @{ type='any';    desc='width of the whole, points' }
    scale    = @{ type='any';    desc='points a capture pixel (0.75 = native at 96 dpi)' }
    nextRow  = @{ type='int';    desc='first row below the pictures, plus gapRows' }
    removed  = @{ type='int';    desc='shapes of an earlier run replaced' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'file_not_found';  transient = $true  }
    @{ id = 'input_invalid';   transient = $false }
    @{ id = 'insert_failed';   transient = $true  }
  )
  example    = @{ use = 'excel.place_tiles'; with = @{ workbook = 'wb'; sheet = 'data'; cell = 'B20'; tiles = '{{steps.wide.out.tiles}}'; layout = '{{steps.wide.out.layout}}' } }
  notes      = 'Idempotent by name: a rerun deletes the group / pictures an earlier run named <name> / <name>_<n> first; pictures put there by hand are never touched. Every picture goes in at its native size (AddPicture -1 / -1), never scaled.'
}

function ExcelPlaceTiles-Plan {
    <#
      PURE. Where each capture goes, in capture pixels from the first
      capture's left edge: @(@{ index; leftPx; cropPx }) -- capture k's
      column 0 must land where the first capture's column (starts[k] -
      starts[0]) is; cropPx = the strip cropped off its left.
    #>
    param($Layout, [int]$Count, [bool]$CropStrip)
    $st = @(@($Layout['starts']) | ForEach-Object { [int]$_ })
    $cpx = [int]$Layout['charPx']; $L = [int]$Layout['textLeft']
    $out = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $Count; $k++) {
        $s = if ($k -lt $st.Count) { $st[$k] } else { 0 }
        $base = if ($st.Count -gt 0) { $st[0] } else { 0 }
        $crop = if ($CropStrip -and $k -gt 0) { $L } else { 0 }
        [void]$out.Add(@{ index = $k; leftPx = (($s - $base) * $cpx + $crop); cropPx = $crop })
    }
    return $out.ToArray()
}

function ExcelPlaceTiles-IsOwn {
    # PURE. A shape this step named under Name on an earlier run?
    param([string]$ShapeName, [string]$Name)
    if ($Name -eq '') { return $false }
    if ($ShapeName -eq $Name) { return $true }
    return ($ShapeName -match ('^' + [regex]::Escape($Name) + '_\d+$'))
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $none = @{ name = ''; shapes = @(); widthPt = 0; scale = 0; nextRow = 0; removed = 0 }
    $paths = @(@($In['tiles']) | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir $work } | Where-Object { $_ -ne '' })
    $lay = $In['layout']
    $lined = ($lay -is [System.Collections.IDictionary] -and $lay.Contains('starts'))
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would put {0} captures side by side on {1}!{2} and group them as {3}' -f $paths.Count, $In['sheet'], $In['cell'], $In['name'])); $r = $none.Clone(); $r['ok'] = $true; return $r }
    if ($paths.Count -eq 0 -or ($lined -and (@($lay['starts']).Count -lt $paths.Count -or [int]$lay['charPx'] -le 0))) {
        $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'input_invalid'; $r['message'] = 'tiles and a layout with a start for each (screen.stitch_scroll) are needed'; return $r
    }
    foreach ($p in $paths) { if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'file_not_found'; $r['message'] = ('no capture at ' + $p); return $r } }
    $ws = Get-EbiSheet -Workbook $In['workbook'] -Sheet $In['sheet']
    if ($null -eq $ws) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'sheet_not_found'; $r['message'] = ('sheet "' + $In['sheet'] + '" not found'); return $r }
    $name = [string]$In['name']
    try {
        $removed = 0
        $old = New-Object System.Collections.ArrayList
        try { foreach ($sh in @($ws.Shapes)) { if (ExcelPlaceTiles-IsOwn -ShapeName ([string]$sh.Name) -Name $name) { [void]$old.Add($sh) } } } catch { }
        foreach ($sh in $old) { try { [void]$sh.Delete(); $removed++ } catch { } }
        $target = $ws.Range([string]$In['cell'])
        if ([int]$In['row'] -gt 0) { $target = $ws.Cells.Item([int]$In['row'], [int]$target.Column) }
        $left = [double]$target.Left; $top = [double]$target.Top
        $plan = if ($lined) { ExcelPlaceTiles-Plan -Layout $lay -Count $paths.Count -CropStrip ([bool]$In['cropStrip']) } else { $null }
        $names = New-Object System.Collections.ArrayList
        $scale = 0.0; $right = $left; $bottom = $top
        for ($k = 0; $k -lt $paths.Count; $k++) {
            $pic = $ws.Shapes.AddPicture($paths[$k], 0, -1, $left, $top, -1, -1)
            if ($k -eq 0) {
                $tw = if ($lined) { [int]$lay['tileW'] } else { 0 }
                $scale = if ($tw -gt 0) { [double]$pic.Width / $tw } else { 0.75 }
            }
            if ($lined) {
                $pc = $plan[$k]
                if ([int]$pc['cropPx'] -gt 0) { $pic.PictureFormat.CropLeft = [double]$pc['cropPx'] * $scale }
                $pic.Left = $left + [double]$pc['leftPx'] * $scale
            } else {
                $pic.Left = $right   # no layout: right after the previous one
            }
            $pic.Top = $top
            $pic.Name = $(if ($paths.Count -eq 1) { $name } else { '{0}_{1}' -f $name, ($k + 1) })
            [void]$names.Add([string]$pic.Name)
            $right = [Math]::Max($right, [double]$pic.Left + [double]$pic.Width)
            $bottom = [Math]::Max($bottom, [double]$pic.Top + [double]$pic.Height)
        }
        $outName = [string]$names[0]
        if ([bool]$In['group'] -and $names.Count -gt 1) {
            $grp = $ws.Shapes.Range([object[]]$names.ToArray()).Group()
            $grp.Name = $name
            try { [void]$grp.ZOrder(1) } catch { }   # behind any red box drawn later
            $outName = [string]$grp.Name
        }
        $startRow = 1
        try { $startRow = [int]$target.Row } catch { }
        $next = Get-EbiRowAtOrBelow -Worksheet $ws -TopPoints $bottom -StartRow $startRow
        Show-EbiExcelAt -Worksheet $ws -Row $startRow -Front -Dwell
        $Ctx.Log.Info(('{0}: {1} picture(s) side by side at {2} ({3:N0} pt wide, {4} pt a pixel){5}' -f $In['sheet'], $names.Count, [string]$target.Address($false, $false), ($right - $left), [Math]::Round($scale, 4), $(if ($removed -gt 0) { ', replacing ' + $removed + ' shape(s) of the last run' } else { '' })))
        return @{ ok = $true; name = $outName; shapes = $names.ToArray(); widthPt = [Math]::Round($right - $left, 2); scale = [Math]::Round($scale, 4); nextRow = ($next + [int]$In['gapRows']); removed = $removed }
    } catch {
        $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'insert_failed'; $r['message'] = (Format-EbiExcelError -Stage 'place tiles' -ErrorRecord $_); return $r
    }
}
