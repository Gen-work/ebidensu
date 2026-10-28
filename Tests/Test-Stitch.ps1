#Requires -Version 5.1
# Test-Stitch.ps1 -- the horizontally long snap (10/28): kernel/Stitch.ps1
# (line ends + widest line in editor columns, the stitch geometry measured
# on the J4 GFIX data pictures, the merged-picture layout, the page-shift
# decision, picture header sizes, the C# overlap search), the picture walk
# of kernel/Xlsx.ps1, and the pure helpers of file.line_ends /
# screen.stitch_scroll / excel.place_tiles / verify.picture_pixels, plus
# file.line_ends run for real on temp files. The GDI+ / Win32 / COM halves
# need Windows; their DryRun contract is Test-StepDryRun.ps1.
# ASCII source; Japanese via [char]. No param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Stitch.ps1')
. (Join-Path $repoRoot 'kernel/Xlsx.ps1')

Reset-Tests 'Stitch'

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-stitch-' + [guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
$log = New-Object PSObject
$log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
$log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
$ctx = @{ WorkDir = $tmp; RunId = 'test'; Profile = @{}; Log = $log; DryRun = $false; Session = @{}; Item = $null; KeyColumns = @('Excel_NAME') }
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
$loaded = @{}
foreach ($u in @('file.line_ends', 'screen.stitch_scroll', 'excel.place_tiles', 'verify.picture_pixels')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    if (-not $r['ok']) { throw ('cannot load ' + $u + ': ' + $r['message']) }
    $loaded[$u] = $r['Entry']
}
function Invoke-TestStep {
    param([string]$Use, [hashtable]$With)
    $e = $loaded[$Use]
    $res = Resolve-EbiStepInputs -Manifest $e['Manifest'] -With $With -Session $ctx['Session']
    if (-not $res['ok']) { throw ($Use + ' inputs: ' + $res['message']) }
    return (& $e['Invoke'] $res['In'] $ctx)
}
function Get-Ascii { param([string]$S) return [System.Text.Encoding]::ASCII.GetBytes($S) }
function Join-Bytes { param([object[]]$Parts) $l = New-Object System.Collections.Generic.List[byte]; foreach ($p in $Parts) { $l.AddRange([byte[]]$p) }; return ,$l.ToArray() }

try {
    # ============================================================ line ends
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "AB`nABCDE`nX`n")
    Assert-True ($i['eol'] -eq 'LF' -and $i['lf'] -eq 3 -and $i['lines'] -eq 3 -and $i['lastEol']) 'eol: three LF lines'
    Assert-True ($i['longestLine'] -eq 2 -and $i['longestCols'] -eq 5 -and $i['needCols'] -eq 7) 'eol: the widest is line 2, 5 columns; a capture must reach 5 + the mark (7)'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "AB`r`nABCDE`r`n")
    Assert-True ($i['eol'] -eq 'CRLF' -and $i['crlf'] -eq 2 -and $i['lf'] -eq 0 -and $i['cr'] -eq 0) 'eol: CRLF is one line end, not a CR and an LF'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "A`rBC`r")
    Assert-True ($i['eol'] -eq 'CR' -and $i['cr'] -eq 2 -and $i['longestCols'] -eq 2) 'eol: CR only'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "A`r`nB`nC`r")
    Assert-True ($i['eol'] -eq 'mixed' -and $i['crlf'] -eq 1 -and $i['lf'] -eq 1 -and $i['cr'] -eq 1 -and $i['lines'] -eq 3) 'eol: one of each = mixed'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "A`nABCDEFG")
    Assert-True (-not $i['lastEol'] -and $i['longestLine'] -eq 2 -and $i['needCols'] -eq 12) 'eol: the widest is the last line with no line end -> [EOF] follows it (7 + 5)'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "ABCDEFG`nA")
    Assert-True (-not $i['lastEol'] -and $i['longestLine'] -eq 1 -and $i['needCols'] -eq 9) 'eol: a short last line without a line end does not move the target'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "`n`n")
    Assert-True ($i['lines'] -eq 2 -and $i['longestCols'] -eq 0 -and $i['longestLine'] -eq 0) 'eol: empty lines'
    $i = Get-EbiLineEndInfo -Bytes ([byte[]]@())
    Assert-True ($i['eol'] -eq 'none' -and $i['lines'] -eq 0 -and $i['needCols'] -eq 0) 'eol: an empty file'
    # SJIS: 'DS' + two kanji + the full-width space + 'X' = 2 + 2+2 + 2 + 1 = 9 columns, 6 characters
    $sj = Join-Bytes @((Get-Ascii 'DS'), [byte[]]@(0x96, 0x7B, 0x93, 0x63, 0x81, 0x40), (Get-Ascii "X`nA`n"))
    $i = Get-EbiLineEndInfo -Bytes $sj -Encoding 'sjis'
    Assert-True ($i['longestCols'] -eq 9 -and $i['longestChars'] -eq 6) ('eol: SJIS double-byte = 2 columns (got ' + $i['longestCols'] + ' / ' + $i['longestChars'] + ')')
    $i = Get-EbiLineEndInfo -Bytes ([byte[]]@(0xB1, 0xB2, 0xB3, 0x0A)) -Encoding 'sjis'
    Assert-Equal 3 $i['longestCols'] 'eol: SJIS half-width kana = 1 column each'
    $u8 = [System.Text.Encoding]::UTF8.GetBytes('DS' + [char]0x672C + [char]0x7530 + [char]0x3000 + "X`n")
    $i = Get-EbiLineEndInfo -Bytes $u8 -Encoding 'utf8'
    Assert-True ($i['longestCols'] -eq 9 -and $i['longestChars'] -eq 6) 'eol: UTF-8 wide characters = 2 columns'
    $i = Get-EbiLineEndInfo -Bytes (Join-Bytes @([byte[]]@(0xEF, 0xBB, 0xBF), $u8)) -Encoding 'sjis'
    Assert-True ($i['bom'] -and $i['longestCols'] -eq 9) 'eol: a UTF-8 BOM switches to utf8 and is not a column'
    $i = Get-EbiLineEndInfo -Bytes (Get-Ascii "A`tB`n") -TabWidth 4
    Assert-Equal 5 $i['longestCols'] 'eol: a tab moves to the next multiple of 4'

    # ============================================================ geometry
    # the J4 GFIX data pictures (10/28): 1906 x 166, column 0 at x 30, the
    # vertical bar from x 1888, the ruler at y 76, the horizontal bar from y 148, 7 px a column
    $g = Get-EbiStitchGeometry -TileW 1906 -TileH 166 -TextLeft 30 -TextRight 19 -BandTop 76 -BandBottom -18 -CharPx 7
    Assert-True ($g['ok'] -and $g['pageCols'] -eq 265 -and $g['pagePx'] -eq 1855 -and $g['bandBottom'] -eq 148) ('geometry: 265 columns = 1855 px a page, band 76..148 (got ' + $g['pageCols'] + ')')
    Assert-True (-not (Get-EbiStitchGeometry -TileW 1906 -TileH 166 -TextLeft 30 -TextRight 19 -BandTop 160 -BandBottom -18 -CharPx 7)['ok']) 'geometry: a band below its own bottom is refused'
    Assert-True (-not (Get-EbiStitchGeometry -TileW 40 -TileH 166 -TextLeft 30 -TextRight 19 -BandTop 76 -BandBottom -18 -CharPx 7)['ok']) 'geometry: no text area is refused'
    Assert-True (-not (Get-EbiStitchGeometry -TileW 1906 -TileH 166 -TextLeft 30 -TextRight 19 -BandTop 76 -BandBottom -18 -CharPx 0)['ok']) 'geometry: charPx 0 is refused'

    # ============================================================ layout
    # five page scrolls: the J4 book's five pictures (0, 265, 530, 795, 1060) join into 9326 x 148
    $lay = Get-EbiStitchLayout -Starts @(0, 265, 530, 795, 1060) -Geometry $g
    Assert-True ($lay['width'] -eq 9326 -and $lay['height'] -eq 148) ('layout: 30 + 5 x 1855 + the 21 px edge = 9326 wide, cut above the horizontal bar (got ' + $lay['width'] + ' x ' + $lay['height'] + ')')
    $pc = @($lay['pieces'])
    Assert-True ($pc.Count -eq 8 -and $pc[0]['h'] -eq 76 -and $pc[0]['w'] -eq 1906 -and $pc[1]['w'] -eq 30) 'layout: the first capture''s header and its line-number strip, then five bands and the edge'
    Assert-True ($pc[3]['tile'] -eq 1 -and $pc[3]['sx'] -eq 30 -and $pc[3]['dx'] -eq 1885 -and $pc[3]['w'] -eq 1855) 'layout: capture 2 adds its whole text area right after capture 1''s (no gap, no repeat)'
    Assert-True ($pc[7]['tile'] -eq 4 -and $pc[7]['sx'] -eq 1885 -and $pc[7]['w'] -eq 21 -and $pc[7]['dx'] -eq 9305) 'layout: the last capture''s right edge closes the band'
    # a scroll cut short at the end: capture 3 starts at 400, repeating 130 columns -> only its new 135
    $lay = Get-EbiStitchLayout -Starts @(0, 265, 400) -Geometry $g
    $pc = @($lay['pieces'])
    Assert-True ($pc[4]['tile'] -eq 2 -and $pc[4]['sx'] -eq 940 -and $pc[4]['w'] -eq 945 -and $pc[4]['dx'] -eq 3740 -and $lay['width'] -eq 4706) 'layout: a capture that repeats columns adds only the new ones'
    Assert-Equal 0 (Get-EbiStitchLayout -Starts @() -Geometry $g)['width'] 'layout: no captures -> nothing'

    # ============================================================ shift
    $s = Resolve-EbiTileShift -ScrollA @{ ok = $true; pos = 0 } -ScrollB @{ ok = $true; pos = 265 } -Match $null -PageCols 265 -CharPx 7
    Assert-True ($s['how'] -eq 'scroll' -and $s['cols'] -eq 265 -and -not $s['same']) 'shift: the scroll bar positions are the truth'
    $s = Resolve-EbiTileShift -ScrollA @{ ok = $true; pos = 795 } -ScrollB @{ ok = $true; pos = 900 } -Match $null -PageCols 265 -CharPx 7
    Assert-True ($s['how'] -eq 'scroll' -and $s['cols'] -eq 105) 'shift: a scroll cut short at the end reads as its real distance'
    $s = Resolve-EbiTileShift -ScrollA @{ ok = $true; pos = 1060 } -ScrollB @{ ok = $true; pos = 1060 } -Match $null -PageCols 265 -CharPx 7
    Assert-True ($s['same'] -and $s['how'] -eq 'same') 'shift: the same position = the view did not move'
    $s = Resolve-EbiTileShift -ScrollA @{ ok = $false } -ScrollB @{ ok = $false } -Match @{ ok = $true; dx = 0; bad = 0.0; ink = 0.2 } -PageCols 265 -CharPx 7
    Assert-True $s['same'] 'shift: no scroll bar, the same picture twice = the end'
    $s = Resolve-EbiTileShift -ScrollA $null -ScrollB $null -Match @{ ok = $true; dx = 945; bad = 0.0005; ink = 0.1 } -PageCols 265 -CharPx 7
    Assert-True ($s['how'] -eq 'match' -and $s['cols'] -eq 135) 'shift: no scroll bar, 130 columns repeated -> moved 135 (match)'
    $s = Resolve-EbiTileShift -ScrollA $null -ScrollB $null -Match @{ ok = $true; dx = 945; bad = 0.0; ink = 0.0 } -PageCols 265 -CharPx 7
    Assert-Equal 'page' $s['how'] 'shift: a blank stretch matches anywhere -> not trusted, a page assumed'
    $s = Resolve-EbiTileShift -ScrollA $null -ScrollB $null -Match @{ ok = $true; dx = 1600; bad = 0.0; ink = 0.1 } -PageCols 265 -CharPx 7
    Assert-Equal 'page' $s['how'] 'shift: an overlap under 40 columns is too short to trust -> a page'
    $s = Resolve-EbiTileShift -ScrollA $null -ScrollB $null -Match @{ ok = $true; dx = 948; bad = 0.0; ink = 0.1 } -PageCols 265 -CharPx 7
    Assert-Equal 'page' $s['how'] 'shift: an offset off the column grid -> not a scroll, a page'
    $s = Resolve-EbiTileShift -ScrollA $null -ScrollB $null -Match @{ ok = $true; dx = 945; bad = 0.2; ink = 0.1 } -PageCols 265 -CharPx 7
    Assert-Equal 'page' $s['how'] 'shift: no good match after a full page scroll -> a page'

    # ============================================================ C# overlap search
    [void](Get-EbiStitchMath)
    $w = 60; $h = 4
    $rnd = New-Object System.Random 7   # not periodic: a periodic pattern matches itself at its period too
    $rows = @(for ($y = 0; $y -lt $h; $y++) { ,@(for ($x = 0; $x -lt 100; $x++) { if ($rnd.Next(4) -eq 0) { [byte]20 } else { [byte]250 } }) })
    $a = New-Object byte[] ($w * $h); $b = New-Object byte[] ($w * $h)
    for ($y = 0; $y -lt $h; $y++) { for ($x = 0; $x -lt $w; $x++) { $a[$y * $w + $x] = $rows[$y][$x]; $b[$y * $w + $x] = $rows[$y][$x + 20] } }
    $o = [EbiStitchMath]::Overlap($a, $b, $w, $h, 10, 40)
    Assert-True ([int]$o[0] -eq 20 -and [double]$o[1] -eq 0 -and [double]$o[2] -gt 0) ('overlap: b = a moved 20 px -> dx 20, nothing differs, ink there (got ' + ($o -join ',') + ')')
    $o = [EbiStitchMath]::Overlap($a, $a, $w, $h, 10, 40)
    Assert-True ([int]$o[0] -eq 0 -and [double]$o[1] -eq 0) 'overlap: the same band -> dx 0'
    $o = [EbiStitchMath]::Overlap($a, $b, $w, $h, 100, 40)
    Assert-Equal -1 ([int]$o[0]) 'overlap: a minimum overlap wider than the band -> not measured'

    # ============================================================ picture headers
    $png = [byte[]]@(0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 13, 0x49, 0x48, 0x44, 0x52, 0, 0, 0x24, 0x6E, 0, 0, 0, 0x94, 8, 2, 0, 0, 0)
    $hs = Get-EbiImageHeadSize -Bytes $png
    Assert-True ($hs['format'] -eq 'png' -and $hs['width'] -eq 9326 -and $hs['height'] -eq 148) 'header: PNG IHDR 9326 x 148'
    $jpg = [byte[]]@(0xFF, 0xD8, 0xFF, 0xE0, 0x00, 0x04, 0x00, 0x00, 0xFF, 0xC0, 0x00, 0x11, 0x08, 0x00, 0xA6, 0x07, 0x72, 0x03, 0, 0, 0, 0, 0, 0)
    $hs = Get-EbiImageHeadSize -Bytes $jpg
    Assert-True ($hs['format'] -eq 'jpeg' -and $hs['width'] -eq 1906 -and $hs['height'] -eq 166) ('header: JPEG SOF0 1906 x 166 (got ' + $hs['width'] + ' x ' + $hs['height'] + ')')
    $hs = Get-EbiImageHeadSize -Bytes ([byte[]]@(0x47, 0x49, 0x46, 0x38, 0x39, 0x61, 0x10, 0x00, 0x20, 0x00, 0, 0))
    Assert-True ($hs['format'] -eq 'gif' -and $hs['width'] -eq 16 -and $hs['height'] -eq 32) 'header: GIF'
    Assert-Equal '' (Get-EbiImageHeadSize -Bytes ([byte[]]@(1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 11)))['format'] 'header: not a picture'

    # ============================================================ xlsx picture walk
    $xdr = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
    $am = 'http://schemas.openxmlformats.org/drawingml/2006/main'
    $rr = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $pic = { param($id, $name, $rid, $cx, $cy, $src) ('<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="{0}" name="{1}"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill><a:blip r:embed="{2}"/>{5}<a:stretch/></xdr:blipFill><xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="{3}" cy="{4}"/></a:xfrm></xdr:spPr></xdr:pic>' -f $id, $name, $rid, $cx, $cy, $src) }
    $drawing = '<xdr:wsDr xmlns:xdr="' + $xdr + '" xmlns:a="' + $am + '" xmlns:r="' + $rr + '">' +
        '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>3</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>9</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        (& $pic 2 'verifyWidePng' 'rId1' (9326 * 9525) (148 * 9525) '') + '<xdr:clientData/></xdr:twoCellAnchor>' +
        '<xdr:twoCellAnchor><xdr:from><xdr:col>1</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>20</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>9</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>29</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' +
        '<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id="5" name="verifyWide"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr><xdr:grpSpPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + (1906 * 9525) + '" cy="' + (166 * 9525) + '"/><a:chOff x="0" y="0"/><a:chExt cx="' + (3812 * 9525) + '" cy="' + (332 * 9525) + '"/></a:xfrm></xdr:grpSpPr>' +
        (& $pic 3 'verifyWide_1' 'rId2' (1906 * 9525 * 2) (166 * 9525 * 2) '') + (& $pic 4 'verifyWide_2' 'rId2' (1877 * 9525 * 2) (166 * 9525 * 2) '<a:srcRect l="1521"/>') +
        '</xdr:grpSp><xdr:clientData/></xdr:twoCellAnchor></xdr:wsDr>'
    $rel = { param($pairs) '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + (@($pairs | ForEach-Object { '<Relationship Id="' + $_[0] + '" Type="x" Target="' + $_[1] + '"/>' }) -join '') + '</Relationships>' }
    $book = @{ parts = @{
        'xl/worksheets/sheet1.xml' = '<worksheet/>'
        'xl/worksheets/_rels/sheet1.xml.rels' = (& $rel @(,@('rId9', '../drawings/drawing1.xml')))
        'xl/drawings/drawing1.xml' = $drawing
        'xl/drawings/_rels/drawing1.xml.rels' = (& $rel @(@('rId1', '../media/image1.png'), @('rId2', '../media/image2.png')))
    } }
    $ps = @(Get-EbiXlsxPictures -Book $book -Part 'xl/worksheets/sheet1.xml')
    Assert-Equal 3 $ps.Count 'xlsx: three pictures, the two inside the group too'
    Assert-True ($ps[0]['name'] -eq 'verifyWidePng' -and $ps[0]['group'] -eq '' -and $ps[0]['media'] -eq 'xl/media/image1.png' -and $ps[0]['shownW'] -eq 9326) 'xlsx: the merged picture, its media and its shown width in px'
    Assert-True ($ps[1]['group'] -eq 'verifyWide' -and $ps[1]['shownW'] -eq 1906 -and $ps[1]['shownH'] -eq 166) 'xlsx: a group member''s size goes through the group''s ext / chExt (half here)'
    Assert-True ($ps[2]['crop']['l'] -eq 1521 -and $ps[2]['shownW'] -eq 1877) 'xlsx: the left crop is read (1.5 %, the J4 books'' later captures)'

    # ============================================================ step helpers
    Assert-Equal 5 (ScreenStitchScroll-Plan -NeedCols 1213 -PageCols 265 -MaxTiles 40) 'stitch plan: 1213 columns at 265 a page = 5 captures'
    Assert-Equal 1 (ScreenStitchScroll-Plan -NeedCols 200 -PageCols 265 -MaxTiles 40) 'stitch plan: a short file = 1 capture'
    Assert-Equal 3 (ScreenStitchScroll-Plan -NeedCols 99999 -PageCols 265 -MaxTiles 3) 'stitch plan: capped at maxTiles'
    $wins = @(@{ handle = 11; title = 'old.txt - sakura' }, @{ handle = 12; title = 'DATA1.261026 - sakura' }, @{ handle = 13; title = 'DATA1.261026 - sakura' })
    $pk = ScreenStitchScroll-PickWindow -Windows $wins -Before @('12') -Title 'DATA1.261026'
    Assert-True ($pk['handle'] -eq 13 -and $pk['ours']) 'stitch window: the new window showing the file, ours to close'
    $pk = ScreenStitchScroll-PickWindow -Windows $wins -Before @('12', '13') -Title 'DATA1.261026'
    Assert-True ($pk['handle'] -eq 12 -and -not $pk['ours']) 'stitch window: only an old one -> captured, never closed'
    Assert-True ((ScreenStitchScroll-PickWindow -Windows $wins -Before @() -Title 'nothing')['handle'] -eq [IntPtr]::Zero) 'stitch window: no title match -> none'
    $tp = ScreenStitchScroll-TroughPoint -BarRect @{ X = 100; Y = 500; W = 1800; H = 17 } -Visible @{ X = 0; Y = 0; W = 0; H = 0 }
    Assert-True ($tp['x'] -eq 1879 -and $tp['y'] -eq 508) 'stitch click: left of the bar''s right arrow, in its middle'
    $tp = ScreenStitchScroll-TroughPoint -BarRect $null -Visible @{ X = 40; Y = 40; W = 1906; H = 166 } -Right 40 -Bottom 8
    Assert-True ($tp['x'] -eq 1906 -and $tp['y'] -eq 198) 'stitch click: no bar found -> clickRight / clickBottom off the bottom-right corner'

    $tl = @(ExcelPlaceTiles-Plan -Layout @{ starts = @(0, 265, 530); charPx = 7; textLeft = 30 } -Count 3 -CropStrip $true)
    Assert-True ($tl[0]['leftPx'] -eq 0 -and $tl[0]['cropPx'] -eq 0 -and $tl[1]['leftPx'] -eq 1885 -and $tl[1]['cropPx'] -eq 30 -and $tl[2]['leftPx'] -eq 3740) 'place tiles: each later capture cropped by its strip and put where its column 0 continues the first'
    $tl = @(ExcelPlaceTiles-Plan -Layout @{ starts = @(0, 265); charPx = 7; textLeft = 30 } -Count 2 -CropStrip $false)
    Assert-True ($tl[1]['leftPx'] -eq 1855 -and $tl[1]['cropPx'] -eq 0) 'place tiles: uncropped, the later capture''s strip covers the earlier''s last columns'
    Assert-True ((ExcelPlaceTiles-IsOwn -ShapeName 'verifyWide_3' -Name 'verifyWide') -and (ExcelPlaceTiles-IsOwn -ShapeName 'verifyWide' -Name 'verifyWide') -and -not (ExcelPlaceTiles-IsOwn -ShapeName 'verifyWideX' -Name 'verifyWide') -and -not (ExcelPlaceTiles-IsOwn -ShapeName 'Picture 3' -Name 'verifyWide')) 'place tiles: only its own names are replaced'

    Assert-Equal '' (FileLineEnds-Alert -Eol 'LF' -Expect 'LF' -CrLf 0 -Lf 3 -Cr 0) 'line ends alert: as expected'
    Assert-True ((FileLineEnds-Alert -Eol 'CRLF' -Expect 'LF' -CrLf 3 -Lf 0 -Cr 0) -like '*CRLF*expected LF*') 'line ends alert: not the expected one'
    Assert-True ((FileLineEnds-Alert -Eol 'mixed' -Expect '' -CrLf 1 -Lf 2 -Cr 0) -like 'mixed*') 'line ends alert: mixed always'

    $j = VerifyPicturePixels-Judge -P @{ format = 'png'; mediaW = 9326; mediaH = 148; shownW = 9326; shownH = 148; crop = @{ l = 0; t = 0; r = 0; b = 0 } } -SourceW 9326 -SourceH 148 -MinScale 0.995
    Assert-True ($j['verdict'] -eq 'ok' -and $j['scale'] -eq 1) 'pixels: PNG kept at its size and shown 1:1 -> ok'
    $j = VerifyPicturePixels-Judge -P @{ format = 'png'; mediaW = 1906; mediaH = 166; shownW = 1877; shownH = 166; crop = @{ l = 1521; t = 0; r = 0; b = 0 } } -SourceW 0 -SourceH 0 -MinScale 0.995
    Assert-Equal 'ok' $j['verdict'] 'pixels: a left crop is not a scale-down (the J4 later captures)'
    $j = VerifyPicturePixels-Judge -P @{ format = 'png'; mediaW = 562; mediaH = 428; shownW = 536; shownH = 408; crop = @{ l = 0; t = 0; r = 0; b = 0 } } -SourceW 0 -SourceH 0 -MinScale 0.995
    Assert-True ($j['verdict'] -eq 'ng' -and $j['why'] -like '*shown at 95*') ('pixels: shown smaller than its pixels -> ng (' + $j['why'] + ')')
    $j = VerifyPicturePixels-Judge -P @{ format = 'png'; mediaW = 4663; mediaH = 74; shownW = 9326; shownH = 148; crop = @{ l = 0; t = 0; r = 0; b = 0 } } -SourceW 9326 -SourceH 148 -MinScale 0.995
    Assert-True ($j['verdict'] -eq 'ng' -and $j['why'] -like '*resampled*') 'pixels: stored with fewer pixels than the capture -> ng'
    $j = VerifyPicturePixels-Judge -P @{ format = 'jpeg'; mediaW = 9326; mediaH = 148; shownW = 9326; shownH = 148; crop = @{ l = 0; t = 0; r = 0; b = 0 } } -SourceW 9326 -SourceH 148 -MinScale 0.995
    Assert-True ($j['verdict'] -eq 'ng' -and $j['why'] -like '*JPEG*') 'pixels: stored as JPEG -> ng'
    Assert-Equal 'unknown' (VerifyPicturePixels-Judge -P @{ format = ''; mediaW = 0; mediaH = 0; shownW = 1; shownH = 1; crop = @{ l = 0; t = 0; r = 0; b = 0 } } -SourceW 0 -SourceH 0 -MinScale 0.995)['verdict'] 'pixels: unreadable media -> unknown'
    $pk = @(VerifyPicturePixels-Pick -Pictures $ps -Names @('verifyWide'))
    Assert-True ($pk.Count -eq 2 -and $pk[0]['name'] -eq 'verifyWide_1') 'pixels: a group name picks its pictures'
    Assert-Equal 3 @(VerifyPicturePixels-Pick -Pictures $ps -Names @()).Count 'pixels: no names = every picture'

    # ============================================================ file.line_ends for real
    $f = Join-Path $tmp 'DATA1.261026.1442'
    [System.IO.File]::WriteAllBytes($f, (Join-Bytes @((Get-Ascii "H2026`n"), (Get-Ascii 'DS'), [byte[]]@(0x96, 0x7B, 0x93, 0x63), (Get-Ascii "1234`nDS`n"))))
    $r = Invoke-TestStep 'file.line_ends' @{ path = $f; expect = 'LF' }
    Assert-True ($r['ok'] -and $r['eol'] -eq 'LF' -and $r['lines'] -eq 3 -and $r['longestLine'] -eq 2 -and $r['longestCols'] -eq 10 -and $r['needCols'] -eq 12 -and $r['alert'] -eq '') ('line_ends: LF, line 2 is 10 columns, reach 12 (got ' + $r['longestCols'] + '/' + $r['needCols'] + ')')
    $r = Invoke-TestStep 'file.line_ends' @{ path = $f; expect = 'CRLF' }
    Assert-True ($r['ok'] -and $r['alert'] -like '*expected CRLF*') 'line_ends: the wrong line end is the alert, not a failure'
    $r = Invoke-TestStep 'file.line_ends' @{ path = (Join-Path $tmp 'none.txt') }
    Assert-Equal 'file_not_found' $r['failure'] 'line_ends: no file -> file_not_found'
    $r = Invoke-TestStep 'verify.picture_pixels' @{ path = (Join-Path $tmp 'none.xlsx'); sheet = 'x' }
    Assert-Equal 'file_not_found' $r['failure'] 'picture_pixels: no workbook -> file_not_found'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
