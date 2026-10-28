# modules/verify/verify.picture_pixels.ps1
# Did a picture keep its pixels in the saved workbook? (10/28: "a big
# snap pasted loses resolution, but not sure".) Read the SAVED .xlsx
# without Excel (kernel/Xlsx.ps1): for every named picture -- the members
# of a named group too -- the stored media's format and pixel size, and
# the size it is shown at. Lost resolution shows as one of:
#   * stored as JPEG (lossy) instead of PNG
#   * stored with fewer pixels than the capture (Excel compressed it)
#   * shown smaller than its pixels (scaled down on the sheet: every
#     pixel no longer one screen pixel)
# Run it after excel.close save = true.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Xlsx.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Stitch.ps1')   # Get-EbiImageHeadSize

$Manifest = @{
  id         = 'verify.picture_pixels'
  group      = 'verify'
  summary    = 'Read a saved workbook: are the named pictures stored and shown at their full pixels (PNG, not resampled)?'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path     = @{ type='path';   required=$true; desc='the saved .xlsx' }
    sheet    = @{ type='string'; required=$true }
    names    = @{ type='list';   default=@(); desc='picture or group names (a group = all its pictures); empty = every picture on the sheet' }
    expect   = @{ type='list';   default=@(); desc='the source PNGs of the matched pictures, in drawing order: their pixel sizes must be kept' }
    minScale = @{ type='any';    default=0.995; desc='shown / stored (after a crop) below this = scaled down on the sheet' }
  }
  outputs    = @{
    code     = @{ type='string'; desc='ok | ng | unknown' }
    pictures = @{ type='list';   desc='@{ name; group; media; format; bytes; mediaW; mediaH; shownW; shownH; scale; sourceW; sourceH; verdict; why }' }
    reason   = @{ type='string' }
    alert    = @{ type='string'; desc='the reason when code is not ok' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'read_failed';    transient = $true  }
    @{ id = 'sheet_not_found'; transient = $false }
  )
  example    = @{ use = 'verify.picture_pixels'; with = @{ path = '{{vars.book}}'; sheet = '{{vars.sheet}}'; names = @('verifyWide', 'verifyWidePng'); expect = @() } }
  notes      = 'A check, never a stop: ng is the gate''s alert. Excel keeps a PNG as it was given unless the file''s picture compression applies (it resamples pictures shown smaller than they are, down to the default resolution) -- so ng with scale < 1 is the one to fix.'
}

function VerifyPicturePixels-Judge {
    <#
      PURE. One picture -> @{ verdict; why; scale }.
        P   @{ format; mediaW; mediaH; shownW; shownH; crop = @{ l; r; t; b } (1/100000) }
        SourceW / SourceH  the capture's size (0 = not known)
    #>
    param($P, [int]$SourceW, [int]$SourceH, [double]$MinScale)
    $why = New-Object System.Collections.ArrayList
    $mw = [double]$P['mediaW']; $mh = [double]$P['mediaH']
    if ([string]$P['format'] -eq '' -or $mw -le 0) { return @{ verdict = 'unknown'; why = 'the stored picture could not be read'; scale = 0 } }
    $c = $P['crop']
    $vw = $mw * (1.0 - ([double]$c['l'] + [double]$c['r']) / 100000.0)
    $vh = $mh * (1.0 - ([double]$c['t'] + [double]$c['b']) / 100000.0)
    $sc = if ($vw -gt 0 -and $vh -gt 0) { [Math]::Min([double]$P['shownW'] / $vw, [double]$P['shownH'] / $vh) } else { 0 }
    $bad = $false
    if ([string]$P['format'] -eq 'jpeg') { [void]$why.Add('stored as JPEG (lossy)'); $bad = $true }
    elseif (@('png', 'gif', 'bmp') -notcontains [string]$P['format']) { [void]$why.Add(('stored as {0}' -f $P['format'])) }
    if ($SourceW -gt 0 -and ([int]$mw -ne $SourceW -or [int]$mh -ne $SourceH)) { [void]$why.Add(('stored {0}x{1}, the capture is {2}x{3}: resampled' -f [int]$mw, [int]$mh, $SourceW, $SourceH)); $bad = $true }
    if ($sc -gt 0 -and $sc -lt $MinScale) { [void]$why.Add(('shown at {0:N1} % of its pixels' -f ($sc * 100))); $bad = $true }
    if (-not $bad -and $why.Count -eq 0) { [void]$why.Add(('{0} {1}x{2}, shown 1:1' -f $P['format'], [int]$mw, [int]$mh)) }
    return @{ verdict = $(if ($bad) { 'ng' } else { 'ok' }); why = ($why.ToArray() -join '; '); scale = [Math]::Round($sc, 4) }
}

function VerifyPicturePixels-Pick {
    # PURE. Pictures whose name or group is one of Names (all when Names is empty), drawing order kept.
    param($Pictures, $Names)
    $n = @(@($Names) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' })
    if ($n.Count -eq 0) { return @($Pictures) }
    return @(@($Pictures) | Where-Object { $n -contains [string]$_['name'] -or ([string]$_['group'] -ne '' -and $n -contains [string]$_['group']) })
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $path = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir $work
    $none = @{ code = 'unknown'; pictures = @(); reason = ''; alert = '' }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would read the pictures of {0}!{1} and check their pixels' -f $path, $In['sheet'])); $r = $none.Clone(); $r['ok'] = $true; return $r }
    $book = Read-EbiXlsx -Path $path -Media
    if (-not $book['ok']) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = $(if ($book['failure'] -eq 'file_not_found') { 'file_not_found' } else { 'read_failed' }); $r['message'] = $book['message']; return $r }
    $sh = @(Get-EbiXlsxSheets -Book $book | Where-Object { [string]$_['name'] -eq [string]$In['sheet'] })
    if ($sh.Count -eq 0) { $r = $none.Clone(); $r['ok'] = $false; $r['failure'] = 'sheet_not_found'; $r['message'] = ('no sheet "' + $In['sheet'] + '" in ' + $path); return $r }
    $pics = @(VerifyPicturePixels-Pick -Pictures @(Get-EbiXlsxPictures -Book $book -Part ([string]$sh[0]['part'])) -Names $In['names'])
    $src = @(@($In['expect']) | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir $work })
    $rows = New-Object System.Collections.ArrayList
    $k = 0
    foreach ($p in $pics) {
        $m = if ($book['media'].Contains([string]$p['media'])) { $book['media'][[string]$p['media']] } else { $null }
        $hs = if ($null -ne $m) { Get-EbiImageHeadSize -Bytes ([byte[]]$m['head']) } else { @{ format = ''; width = 0; height = 0 } }
        $sw = 0; $shh = 0
        if ($k -lt $src.Count -and $src[$k] -ne '' -and (Test-Path -LiteralPath $src[$k] -PathType Leaf)) {
            $fs = [System.IO.File]::OpenRead($src[$k])
            try { $buf = New-Object byte[] ([int][Math]::Min([int64]65536, $fs.Length)); [void]$fs.Read($buf, 0, $buf.Length) } finally { $fs.Dispose() }
            $ss = Get-EbiImageHeadSize -Bytes $buf; $sw = [int]$ss['width']; $shh = [int]$ss['height']
        }
        $row = @{ name = [string]$p['name']; group = [string]$p['group']; media = [string]$p['media']; format = [string]$hs['format']; bytes = $(if ($null -ne $m) { [int64]$m['length'] } else { 0 });
                  mediaW = [int]$hs['width']; mediaH = [int]$hs['height']; shownW = $p['shownW']; shownH = $p['shownH']; crop = $p['crop']; sourceW = $sw; sourceH = $shh }
        $j = VerifyPicturePixels-Judge -P $row -SourceW $sw -SourceH $shh -MinScale ([double]$In['minScale'])
        $row['verdict'] = $j['verdict']; $row['why'] = $j['why']; $row['scale'] = $j['scale']
        $row.Remove('crop')
        [void]$rows.Add($row)
        $Ctx.Log.Info(('{0}{1}: {2} -- {3}' -f $row['name'], $(if ($row['group'] -ne '') { ' (in ' + $row['group'] + ')' } else { '' }), $row['verdict'], $row['why']))
        $k++
    }
    $ng = @($rows | Where-Object { $_['verdict'] -eq 'ng' })
    $unk = @($rows | Where-Object { $_['verdict'] -eq 'unknown' })
    $code = if ($rows.Count -eq 0) { 'unknown' } elseif ($ng.Count -gt 0) { 'ng' } elseif ($unk.Count -gt 0) { 'unknown' } else { 'ok' }
    $reason = switch ($code) {
        'ok'    { ('{0}: {1} picture(s) stored and shown at their full pixels' -f $In['sheet'], $rows.Count) }
        'ng'    { ('{0}: {1} of {2} picture(s) lost pixels -- {3}' -f $In['sheet'], $ng.Count, $rows.Count, (@($ng | ForEach-Object { $_['name'] + ': ' + $_['why'] }) -join ' / ')) }
        default { if ($rows.Count -eq 0) { ('{0}: no picture named {1}' -f $In['sheet'], (@($In['names']) -join ', ')) } else { ('{0}: {1} picture(s) could not be read' -f $In['sheet'], $unk.Count) } }
    }
    return @{ ok = $true; code = $code; pictures = $rows.ToArray(); reason = $reason; alert = $(if ($code -eq 'ok') { '' } else { $reason }) }
}
