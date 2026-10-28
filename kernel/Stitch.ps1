#Requires -Version 5.1
# ============================================================
#  kernel/Stitch.ps1
#
#  A horizontally long snap (10/28): a text viewer shows one window-width
#  of a long line; the view is scrolled a page at a time, each page is
#  captured, and the pages are put side by side -- either merged into one
#  long PNG here (outside Excel) or placed on the sheet as separate
#  pictures lined up and grouped (inside Excel, excel.place_tiles). The
#  evidence books did this by hand: GFIX data opened in a text editor, five
#  window captures of 265 columns each, the later four cropped by the
#  line-number strip.
#
#  "Up to the end of the line": the data file itself is read
#  (Get-EbiLineEndInfo) -- the longest line's width in display columns and
#  what ends its lines (CRLF / LF / CR) -- so the capture knows how far
#  right it has to go and can say whether the last picture shows the line
#  end mark.
#
#    Get-EbiLineEndInfo -Bytes -Encoding -TabWidth
#                                 -> @{ lines; eol; crlf; lf; cr; lastEol;
#                                    longestLine; longestCols; longestChars;
#                                    needCols; bom }
#    Get-EbiStitchGeometry -TileW -TileH -TextLeft -TextRight -BandTop
#                          -BandBottom -CharPx
#                                 -> @{ ok; pageCols; pagePx; ... }
#    Resolve-EbiTileShift         one page step: scroll position / pixel
#                                 match / page assumption -> columns
#    Get-EbiStitchLayout -Starts -Geometry
#                                 -> @{ width; height; pieces } (what is
#                                    drawn where in the merged picture)
#    Get-EbiImageHeadSize -Bytes  -> @{ format; width; height } (PNG /
#                                    JPEG / GIF / BMP header, no GDI+)
#    Get-EbiStitchMath            C# sliding compare (no System.Drawing)
#    Find-EbiTileOverlap          two captures -> the best overlap
#    New-EbiStitchedPng           the merged PNG (GDI+)
#
#  Pure parts are unit-tested (Tests\Test-Stitch.ps1); the GDI+ halves
#  follow kernel/Image.ps1's split (a *Core names System.Drawing types).
#  Dot-source only (no param(), ASCII source, no class).
# ============================================================

function Get-EbiLineEndInfo {
    <#
      PURE. The bytes of a text file -> how its lines end and how wide the
      widest is, in the columns a text editor shows:
        sjis   (also plain ASCII) a lead byte 0x81-0x9F / 0xE0-0xFC starts a
               2-byte, 2-column character; any other byte is 1 column
               (half-width kana 0xA1-0xDF included)
        utf8   East Asian wide / full-width code points 2 columns, others 1
               (also chosen by a UTF-8 BOM)
      A tab moves to the next multiple of TabWidth. The line end itself is
      not counted. Counted in C# (Get-EbiStitchMath): a byte loop in
      PowerShell takes minutes on a large data file.
        eol       'CRLF' | 'LF' | 'CR' | 'mixed' | 'none'
        lastEol   the last line ends with a line end (else the editor's
                  [EOF] mark follows its last character on the same line)
        needCols  the columns a capture must reach so the line end mark of
                  the widest line is on it: longestCols + the mark (2
                  columns: the mark and the cell after it; [EOF] is 5)
    #>
    param([byte[]]$Bytes, [string]$Encoding = 'sjis', [int]$TabWidth = 4)
    $out = @{ lines = 0; eol = 'none'; crlf = 0; lf = 0; cr = 0; lastEol = $false; longestLine = 0; longestCols = 0; longestChars = 0; needCols = 0; bom = $false }
    if ($null -eq $Bytes -or $Bytes.Length -eq 0) { return $out }
    if ($TabWidth -le 0) { $TabWidth = 4 }
    $start = 0
    if ($Bytes.Length -ge 3 -and $Bytes[0] -eq 0xEF -and $Bytes[1] -eq 0xBB -and $Bytes[2] -eq 0xBF) { $out['bom'] = $true; $start = 3; $Encoding = 'utf8' }
    [void](Get-EbiStitchMath)
    $r = [EbiStitchMath]::Lines($Bytes, $start, ($Encoding -eq 'utf8'), $TabWidth)
    $out['lines'] = [int]$r[0]; $out['crlf'] = [int]$r[1]; $out['lf'] = [int]$r[2]; $out['cr'] = [int]$r[3]
    $out['lastEol'] = ([int]$r[4] -eq 0)
    $out['longestLine'] = [int]$r[5]; $out['longestCols'] = [int]$r[6]; $out['longestChars'] = [int]$r[7]
    $kinds = @(@('crlf', 'lf', 'cr') | Where-Object { $out[$_] -gt 0 })
    $out['eol'] = if ($kinds.Count -eq 0) { 'none' } elseif ($kinds.Count -gt 1) { 'mixed' } else { ([string]$kinds[0]).ToUpperInvariant() }
    # the mark after the widest line: a line end mark (1 column, plus one so
    # it is never the half-cut last column), or [EOF] when the file's last
    # line is the widest and has no line end
    $mark = if (-not $out['lastEol'] -and $out['longestLine'] -eq $out['lines']) { 5 } else { 2 }
    $out['needCols'] = $out['longestCols'] + $mark
    return $out
}

function Get-EbiStitchGeometry {
    <#
      PURE. Where the text is in one capture of the viewer's window.
        TextLeft    px from the capture's left edge to column 0 (past the
                    line-number strip)
        TextRight   px from the right edge back to the end of the text
                    area (the vertical scroll bar)
        BandTop     first row of the band that is stitched (the ruler)
        BandBottom  row after the band; <= 0 counts from the bottom edge
                    (-18 = stop above an 18 px horizontal scroll bar)
        CharPx      px a half-width column
      -> @{ ok; message; tileW; tileH; textLeft; textRight; textW;
            bandTop; bandBottom; charPx; pageCols; pagePx }
      pageCols = whole columns the text area shows = what a page scroll
      moves (10/28, measured on the J4 GFIX data pictures: 1906 px wide,
      column 0 at 30, the scroll bar at 1888, 7 px a column -> 265, and
      the five pictures join at exactly 265 x 7 = 1855 px).
    #>
    param([int]$TileW, [int]$TileH, [int]$TextLeft, [int]$TextRight, [int]$BandTop, [int]$BandBottom, [int]$CharPx)
    $g = @{ ok = $false; message = ''; tileW = $TileW; tileH = $TileH; textLeft = $TextLeft; textRight = $TextRight; textW = 0; bandTop = $BandTop; bandBottom = $BandBottom; charPx = $CharPx; pageCols = 0; pagePx = 0 }
    if ($TileW -le 0 -or $TileH -le 0) { $g['message'] = 'the capture has no size'; return $g }
    if ($CharPx -le 0) { $g['message'] = 'charPx must be > 0'; return $g }
    $bb = if ($BandBottom -le 0) { $TileH + $BandBottom } else { [Math]::Min($TileH, $BandBottom) }
    $g['bandBottom'] = $bb
    if ($BandTop -lt 0 -or $bb -le $BandTop) { $g['message'] = ('the band {0}..{1} is empty in a {2} px tall capture' -f $BandTop, $bb, $TileH); return $g }
    $tw = $TileW - $TextLeft - $TextRight
    if ($TextLeft -lt 0 -or $TextRight -lt 0 -or $tw -lt $CharPx) { $g['message'] = ('no text area: {0} px wide capture, textLeft {1}, textRight {2}' -f $TileW, $TextLeft, $TextRight); return $g }
    $g['textW'] = $tw
    $g['pageCols'] = [int][Math]::Floor($tw / $CharPx)
    $g['pagePx'] = $g['pageCols'] * $CharPx
    $g['ok'] = $true
    return $g
}

function Resolve-EbiTileShift {
    <#
      PURE. How far the view moved between two captures, in columns.
        Scroll   @{ ok; pos } of each capture (the horizontal scroll bar's
                 position, in columns) -- the truth when both are read
        Match    @{ ok; dx; bad; ink } Find-EbiTileOverlap of the two
                 captures' text bands
        PageCols what a page scroll moves when nothing else says
      -> @{ cols; how; same }
        how   'scroll'  the scroll bar positions
              'same'    the captures are the same picture: the view did
                        not move (the end of the scroll range)
              'match'   the overlap found: the scroll was cut short at
                        the end of the range (bad <= OkBad over a stretch
                        with ink, at least MinOverlapCols wide)
              'page'    neither: one page assumed (no overlap is expected
                        after a full page scroll)
    #>
    param($ScrollA, $ScrollB, $Match, [int]$PageCols, [int]$CharPx, [double]$OkBad = 0.002, [double]$MinInk = 0.003, [int]$MinOverlapCols = 40)
    if ($ScrollA -is [System.Collections.IDictionary] -and $ScrollB -is [System.Collections.IDictionary] -and [bool]$ScrollA['ok'] -and [bool]$ScrollB['ok']) {
        $d = [int]$ScrollB['pos'] - [int]$ScrollA['pos']
        if ($d -ge 0 -and $d -le $PageCols) { return @{ cols = $d; how = $(if ($d -eq 0) { 'same' } else { 'scroll' }); same = ($d -eq 0) } }
    }
    if ($Match -is [System.Collections.IDictionary] -and [bool]$Match['ok'] -and $CharPx -gt 0) {
        $dx = [int]$Match['dx']; $bad = [double]$Match['bad']
        if ($dx -eq 0 -and $bad -le $OkBad) { return @{ cols = 0; how = 'same'; same = $true } }
        $cols = [int][Math]::Round($dx / [double]$CharPx)
        $onGrid = ([Math]::Abs($dx - $cols * $CharPx) -le 1)
        if ($bad -le $OkBad -and [double]$Match['ink'] -ge $MinInk -and $onGrid -and $cols -gt 0 -and $cols -le ($PageCols - $MinOverlapCols)) {
            return @{ cols = $cols; how = 'match'; same = $false }
        }
    }
    return @{ cols = $PageCols; how = 'page'; same = $false }
}

function Get-EbiStitchLayout {
    <#
      PURE. The merged picture of the captures, given each capture's first
      column (Starts, Starts[0] = 0 for a view that began at the left) and
      the geometry. Frame:
        'aligned' (default, 10/28 operator's own sheet: "like this") every
                capture with its window frame (title, menu, toolbar,
                scroll bar), the later ones cropped by the line-number
                strip (textLeft) and laid on top of the one before so its
                column 0 continues it: ruler and text run on across the
                seams, each capture's frame shows from where it starts --
                what excel.place_tiles builds on the sheet
        'each'  every capture whole, side by side, no overlap
        'first' one frame: as below
      Frame 'first':
        * the whole first capture above the band (title, menu, toolbar --
          the file name and its encoding are evidence) and its line-number
          strip
        * the band (ruler + text rows) of every capture, each starting
          where the previous one's columns end; a capture that repeats
          columns (a scroll cut short) adds only the new ones
        * the last capture's right edge (the vertical scroll bar) closes
          the band
      Rows below the band (the horizontal scroll bar) are left out.
      -> @{ width; height; pieces = @(@{ tile; sx; sy; w; h; dx; dy }) }
    #>
    param($Starts, $Geometry, [string]$Frame = 'aligned')
    $g = $Geometry
    $st = @(@($Starts) | ForEach-Object { [int]$_ })
    $pieces = New-Object System.Collections.ArrayList
    if ($st.Count -eq 0 -or -not [bool]$g['ok']) { return @{ width = 0; height = 0; pieces = @() } }
    $L = [int]$g['textLeft']; $cpx = [int]$g['charPx']; $page = [int]$g['pagePx']
    if ($Frame -eq 'aligned') {
        $tw = [int]$g['tileW']; $base = $st[0]; $width = 0; $lastEnd = -1
        for ($k = 0; $k -lt $st.Count; $k++) {
            $end = $st[$k] + [int]$g['pageCols']
            if ($k -gt 0 -and $end -le $lastEnd) { continue }   # nothing new in it
            $crop = if ($k -gt 0) { $L } else { 0 }
            $dx = ($st[$k] - $base) * $cpx + $crop
            [void]$pieces.Add(@{ tile = $k; sx = $crop; sy = 0; w = ($tw - $crop); h = [int]$g['tileH']; dx = $dx; dy = 0 })
            $width = [Math]::Max($width, $dx + $tw - $crop); $lastEnd = $end
        }
        return @{ width = $width; height = [int]$g['tileH']; pieces = $pieces.ToArray() }
    }
    if ($Frame -eq 'each') {
        $tw = [int]$g['tileW']; $x = 0; $lastEnd = -1
        for ($k = 0; $k -lt $st.Count; $k++) {
            $end = $st[$k] + [int]$g['pageCols']
            if ($k -gt 0 -and $end -le $lastEnd) { continue }   # nothing new in it
            [void]$pieces.Add(@{ tile = $k; sx = 0; sy = 0; w = $tw; h = [int]$g['tileH']; dx = $x; dy = 0 })
            $x += $tw; $lastEnd = $end
        }
        return @{ width = $x; height = [int]$g['tileH']; pieces = $pieces.ToArray() }
    }
    $top = [int]$g['bandTop']; $bot = [int]$g['bandBottom']; $bh = $bot - $top
    $base = $st[0]
    # the first capture: everything above the band, then the strip left of column 0
    [void]$pieces.Add(@{ tile = 0; sx = 0; sy = 0; w = [int]$g['tileW']; h = $top; dx = 0; dy = 0 })
    [void]$pieces.Add(@{ tile = 0; sx = 0; sy = $top; w = $L; h = $bh; dx = 0; dy = $top })
    $drawn = $L   # stitched x up to which the band is drawn
    for ($k = 0; $k -lt $st.Count; $k++) {
        $x0 = $L + ($st[$k] - $base) * $cpx    # stitched x of this capture's column 0
        $x1 = $x0 + $page
        if ($x1 -le $drawn) { continue }        # nothing new in it
        $from = [Math]::Max($drawn, $x0)
        [void]$pieces.Add(@{ tile = $k; sx = ($L + ($from - $x0)); sy = $top; w = ($x1 - $from); h = $bh; dx = $from; dy = $top })
        $drawn = $x1
        $last = $k
    }
    # the right edge after the text area of the last capture drawn (a part column + the scroll bar)
    $edge = [int]$g['tileW'] - ($L + $page)
    if ($edge -gt 0) { [void]$pieces.Add(@{ tile = $last; sx = ($L + $page); sy = $top; w = $edge; h = $bh; dx = $drawn; dy = $top }) }
    $width = [Math]::Max([int]$g['tileW'], $drawn + [Math]::Max(0, $edge))
    return @{ width = $width; height = $bot; pieces = $pieces.ToArray() }
}

function Get-EbiImageHeadSize {
    <#
      PURE. The first bytes of a picture -> @{ format; width; height }
      (format '' when not recognised). PNG IHDR, GIF, BMP, and JPEG SOFn
      (a JPEG needs the bytes up to its frame header: pass the whole file
      or its first 64 KB).
    #>
    param([byte[]]$Bytes)
    $none = @{ format = ''; width = 0; height = 0 }
    if ($null -eq $Bytes -or $Bytes.Length -lt 10) { return $none }
    $b = $Bytes
    if ($b.Length -ge 24 -and $b[0] -eq 0x89 -and $b[1] -eq 0x50 -and $b[2] -eq 0x4E -and $b[3] -eq 0x47) {
        $w = ([int]$b[16] -shl 24) -bor ([int]$b[17] -shl 16) -bor ([int]$b[18] -shl 8) -bor [int]$b[19]
        $h = ([int]$b[20] -shl 24) -bor ([int]$b[21] -shl 16) -bor ([int]$b[22] -shl 8) -bor [int]$b[23]
        return @{ format = 'png'; width = $w; height = $h }
    }
    if ($b[0] -eq 0x47 -and $b[1] -eq 0x49 -and $b[2] -eq 0x46) { return @{ format = 'gif'; width = ([int]$b[6] -bor ([int]$b[7] -shl 8)); height = ([int]$b[8] -bor ([int]$b[9] -shl 8)) } }
    if ($b.Length -ge 26 -and $b[0] -eq 0x42 -and $b[1] -eq 0x4D) {
        $w = [BitConverter]::ToInt32($b, 18); $h = [Math]::Abs([BitConverter]::ToInt32($b, 22))
        return @{ format = 'bmp'; width = $w; height = $h }
    }
    if ($b[0] -eq 0xFF -and $b[1] -eq 0xD8) {
        $i = 2
        while (($i + 9) -lt $b.Length) {
            if ($b[$i] -ne 0xFF) { $i++; continue }
            $m = [int]$b[$i + 1]
            if ($m -eq 0xD8 -or $m -eq 0x01 -or ($m -ge 0xD0 -and $m -le 0xD7)) { $i += 2; continue }
            $len = ([int]$b[$i + 2] -shl 8) -bor [int]$b[$i + 3]
            if (($m -ge 0xC0 -and $m -le 0xCF) -and $m -ne 0xC4 -and $m -ne 0xC8 -and $m -ne 0xCC) {
                return @{ format = 'jpeg'; width = (([int]$b[$i + 7] -shl 8) -bor [int]$b[$i + 8]); height = (([int]$b[$i + 5] -shl 8) -bor [int]$b[$i + 6]) }
            }
            if ($len -lt 2) { break }
            $i += 2 + $len
        }
        return @{ format = 'jpeg'; width = 0; height = 0 }
    }
    if ($b.Length -ge 44 -and $b[0] -eq 0x01 -and $b[1] -eq 0x00 -and $b[2] -eq 0x00 -and $b[3] -eq 0x00 -and $b[40] -eq 0x20 -and $b[41] -eq 0x45 -and $b[42] -eq 0x4D -and $b[43] -eq 0x46) {
        return @{ format = 'emf'; width = 0; height = 0 }
    }
    return $none
}

function Get-EbiStitchMath {
    <#
      Compile the byte / pixel loops once per process (plain C#, no
      System.Drawing: unit-tested on Linux too).
        Lines(bytes, start, utf8, tab)   see Get-EbiLineEndInfo
        Overlap(a, b, w, h, minOverlap, tol)
          a, b: grey bands of the same size. For every dx in
          0 .. w - minOverlap, a's columns dx.. against b's columns 0..:
          the share of pixels differing by more than tol. -> double[]
          { best dx, its bad share, ink share of the overlap in b }.
          Ink = a pixel more than tol away from the band's commonest grey.
    #>
    if (-not ('EbiStitchMath' -as [type])) {
        Add-Type -TypeDefinition @"
public static class EbiStitchMath {
    static bool Wide(int c) {
        // East Asian Wide / Fullwidth, plus arrows and geometric shapes
        // (U+2190.., U+25A0..: ambiguous width, drawn 2 columns by the
        // Japanese fonts the editors use -- the full-width space mark)
        return (c >= 0x1100 && c <= 0x115F) || (c >= 0x2190 && c <= 0x21FF) || (c >= 0x25A0 && c <= 0x25FF)
            || (c >= 0x2E80 && c <= 0x303E) || (c >= 0x3041 && c <= 0x33FF) || (c >= 0x3400 && c <= 0x4DBF)
            || (c >= 0x4E00 && c <= 0x9FFF) || (c >= 0xA000 && c <= 0xA4CF) || (c >= 0xAC00 && c <= 0xD7A3)
            || (c >= 0xF900 && c <= 0xFAFF) || (c >= 0xFE30 && c <= 0xFE4F) || (c >= 0xFF00 && c <= 0xFF60)
            || (c >= 0xFFE0 && c <= 0xFFE6) || (c >= 0x20000 && c <= 0x3FFFD);
    }
    // { lines, crlf, lf, cr, last line open (1/0), longest line (1-based), its columns, its characters }
    public static long[] Lines(byte[] b, int start, bool utf8, int tab) {
        long lines = 0, crlf = 0, lf = 0, cr = 0, bestLine = 0, bestCols = 0, bestChars = 0;
        long cols = 0, chars = 0; bool open = false; int n = b.Length; int i = start;
        while (i < n) {
            int c = b[i];
            if (c == 0x0D || c == 0x0A) {
                if (c == 0x0D && i + 1 < n && b[i + 1] == 0x0A) { crlf++; i += 2; }
                else if (c == 0x0D) { cr++; i++; }
                else { lf++; i++; }
                lines++;
                if (cols > bestCols) { bestCols = cols; bestChars = chars; bestLine = lines; }
                cols = 0; chars = 0; open = false;
                continue;
            }
            open = true;
            if (c == 0x09) { cols = (cols / tab + 1) * tab; chars++; i++; continue; }
            if (!utf8) {
                if (((c >= 0x81 && c <= 0x9F) || (c >= 0xE0 && c <= 0xFC)) && i + 1 < n && b[i + 1] != 0x0D && b[i + 1] != 0x0A) { cols += 2; chars++; i += 2; continue; }
                cols++; chars++; i++; continue;
            }
            int cp = c, len = 1;
            if (c >= 0xF0 && c <= 0xF7) { len = 4; cp = c & 0x07; }
            else if (c >= 0xE0 && c <= 0xEF) { len = 3; cp = c & 0x0F; }
            else if (c >= 0xC0 && c <= 0xDF) { len = 2; cp = c & 0x1F; }
            if (len > 1) {
                if (i + len > n) { len = 1; cp = c; }
                else {
                    for (int k = 1; k < len; k++) {
                        int t = b[i + k];
                        if ((t & 0xC0) != 0x80) { len = 1; cp = c; break; }
                        cp = (cp << 6) | (t & 0x3F);
                    }
                }
            }
            cols += Wide(cp) ? 2 : 1; chars++; i += len;
        }
        if (open) { lines++; if (cols > bestCols) { bestCols = cols; bestChars = chars; bestLine = lines; } }
        return new long[] { lines, crlf, lf, cr, open ? 1 : 0, bestLine, bestCols, bestChars };
    }
    public static double[] Overlap(byte[] a, byte[] b, int w, int h, int minOverlap, int tol) {
        if (w <= 0 || h <= 0 || a.Length < w * h || b.Length < w * h || minOverlap < 1 || minOverlap > w) return new double[] { -1, 1, 0 };
        int[] hist = new int[256];
        for (int i = 0; i < w * h; i++) hist[b[i]]++;
        int mode = 0; for (int i = 1; i < 256; i++) if (hist[i] > hist[mode]) mode = i;
        double best = 2; int bdx = -1;
        for (int dx = 0; dx <= w - minOverlap; dx++) {
            int ov = w - dx; long bad = 0; long limit = (long)(best * ov * h) + 1;
            for (int y = 0; y < h && bad < limit; y++) {
                int ra = y * w + dx, rb = y * w;
                for (int x = 0; x < ov; x++) { int d = a[ra + x] - b[rb + x]; if (d > tol || d < -tol) bad++; }
            }
            double share = (double)bad / ((double)ov * h);
            if (share < best) { best = share; bdx = dx; }
        }
        if (bdx < 0) return new double[] { -1, 1, 0 };
        int bov = w - bdx; long ink = 0;
        for (int y = 0; y < h; y++) for (int x = 0; x < bov; x++) { int d = b[y * w + x] - mode; if (d > tol || d < -tol) ink++; }
        return new double[] { bdx, best, (double)ink / ((double)bov * h) };
    }
}
"@
    }
    return [type]'EbiStitchMath'
}

function Find-EbiTileOverlap {
    <#
      Two captures of the same window -> where the second's text band
      repeats the end of the first's: @{ ok; failure; message; dx; bad;
      ink } (dx in px of the text area; bad = share of pixels that differ;
      see Resolve-EbiTileShift for how it is used).
    #>
    param([string]$PathA, [string]$PathB, $Geometry, [int]$MinOverlapPx = 280, [int]$Tol = 40)
    $g = $Geometry
    $x = [int]$g['textLeft']; $y = [int]$g['bandTop']; $w = [int]$g['pagePx']; $h = [int]$g['bandBottom'] - $y
    $ra = Get-EbiGrayRegion -Path $PathA -X $x -Y $y -W $w -H $h
    if (-not $ra['ok']) { return @{ ok = $false; failure = $ra['failure']; message = $ra['message']; dx = -1; bad = 1; ink = 0 } }
    $rb = Get-EbiGrayRegion -Path $PathB -X $x -Y $y -W $w -H $h
    if (-not $rb['ok']) { return @{ ok = $false; failure = $rb['failure']; message = $rb['message']; dx = -1; bad = 1; ink = 0 } }
    if ([int]$ra['w'] -ne [int]$rb['w'] -or [int]$ra['h'] -ne [int]$rb['h']) { return @{ ok = $false; failure = 'geometry'; message = 'the two captures differ in size'; dx = -1; bad = 1; ink = 0 } }
    [void](Get-EbiStitchMath)
    $r = [EbiStitchMath]::Overlap([byte[]]$ra['g'], [byte[]]$rb['g'], [int]$ra['w'], [int]$ra['h'], [Math]::Min([int]$ra['w'], [Math]::Max(1, $MinOverlapPx)), $Tol)
    if ([double]$r[0] -lt 0) { return @{ ok = $false; failure = 'geometry'; message = 'no overlap could be measured'; dx = -1; bad = 1; ink = 0 } }
    return @{ ok = $true; failure = ''; message = ''; dx = [int]$r[0]; bad = [Math]::Round([double]$r[1], 5); ink = [Math]::Round([double]$r[2], 5) }
}

function New-EbiStitchedPng {
    <#
      Draw the Get-EbiStitchLayout pieces of the captures into one PNG,
      pixel for pixel (no scaling, nearest neighbour), saved at 96 dpi so
      Excel puts it in at 0.75 pt a pixel like any capture. Background
      white (the header right of the first capture).
      -> @{ ok; failure; message; width; height }
    #>
    param($Tiles, $Layout, [string]$Dest)
    $t = @(@($Tiles) | ForEach-Object { [string]$_ })
    if ($t.Count -eq 0 -or [int]$Layout['width'] -le 0 -or [int]$Layout['height'] -le 0) { return @{ ok = $false; failure = 'geometry'; message = 'nothing to stitch'; width = 0; height = 0 } }
    foreach ($p in $t) { if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('no capture at ' + $p); width = 0; height = 0 } } }
    return (New-EbiStitchedPngCore -Tiles $t -Layout $Layout -Dest $Dest)
}

function New-EbiStitchedPngCore {
    # The GDI+ half of New-EbiStitchedPng (see kernel/Image.ps1 Get-EbiPngSizeCore).
    param([string[]]$Tiles, $Layout, [string]$Dest)
    [void](Get-EbiDrawing)
    $imgs = @{}; $streams = New-Object System.Collections.ArrayList; $bmp = $null; $gfx = $null
    try {
        New-EbiParentDirectory -Path $Dest
        $W = [int]$Layout['width']; $H = [int]$Layout['height']
        $bmp = New-Object System.Drawing.Bitmap($W, $H, [System.Drawing.Imaging.PixelFormat]::Format24bppRgb)
        $bmp.SetResolution(96, 96)
        $gfx = [System.Drawing.Graphics]::FromImage($bmp)
        $gfx.Clear([System.Drawing.Color]::White)
        $gfx.InterpolationMode = [System.Drawing.Drawing2D.InterpolationMode]::NearestNeighbor
        $gfx.PixelOffsetMode = [System.Drawing.Drawing2D.PixelOffsetMode]::Half
        $gfx.CompositingMode = [System.Drawing.Drawing2D.CompositingMode]::SourceCopy
        foreach ($pc in @($Layout['pieces'])) {
            if ([int]$pc['w'] -le 0 -or [int]$pc['h'] -le 0) { continue }
            $k = [int]$pc['tile']
            if (-not $imgs.Contains($k)) {
                $ms = New-Object System.IO.MemoryStream(, [System.IO.File]::ReadAllBytes($Tiles[$k]))
                [void]$streams.Add($ms)
                $imgs[$k] = [System.Drawing.Image]::FromStream($ms)
            }
            $src = New-Object System.Drawing.Rectangle([int]$pc['sx'], [int]$pc['sy'], [int]$pc['w'], [int]$pc['h'])
            $dst = New-Object System.Drawing.Rectangle([int]$pc['dx'], [int]$pc['dy'], [int]$pc['w'], [int]$pc['h'])
            $gfx.DrawImage($imgs[$k], $dst, $src, [System.Drawing.GraphicsUnit]::Pixel)
        }
        $bmp.Save($Dest, [System.Drawing.Imaging.ImageFormat]::Png)
        return @{ ok = $true; failure = ''; message = ''; width = $W; height = $H }
    } catch {
        return @{ ok = $false; failure = 'save_failed'; message = $_.Exception.Message; width = 0; height = 0 }
    } finally {
        if ($null -ne $gfx) { $gfx.Dispose() }
        if ($null -ne $bmp) { $bmp.Dispose() }
        foreach ($v in $imgs.Values) { $v.Dispose() }
        foreach ($v in $streams) { $v.Dispose() }
    }
}
