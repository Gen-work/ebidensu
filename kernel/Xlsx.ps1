#Requires -Version 5.1
# ============================================================
#  kernel/Xlsx.ps1
#
#  Read an .xlsx WITHOUT Excel: the workbook is a zip of XML parts. The
#  review phase (verify.book_review, 10/28) checks the evidence books this
#  way -- the PowerShell port of tools/review/drawings.py + the openpyxl
#  reads of check_book.py, for the office PC that has no Python. Read-only;
#  a book open in Excel is read through a shared read handle. Dot-source
#  only (no param(), ASCII source, no class).
#
#    Read-EbiXlsx -Path [-Media]   -> @{ ok; failure; message; path; parts;
#                                     media } parts: part name -> XML text
#                                     (.xml / .rels); media (-Media only):
#                                     xl/media part -> @{ length; head }
#    Get-EbiXlsxPictures -Book -Part
#                                  -> @(@{ name; group; media; shownW;
#                                     shownH; crop }) every picture, group
#                                     members too, with its shown size in
#                                     px (EMU / 9525) and its media part
#    Get-EbiXlsxSheets -Book       -> @(@{ name; part; state }) in tab order
#    Get-EbiXlsxStyles -Book       -> @{ xf; fonts; fills } (cellXfs order)
#    Get-EbiXlsxCells -Book -Part -Styles
#                                  -> @(@{ ref; row; col; text; font; bold;
#                                     yellow }) for every cell with a value
#                                     or a style, in row / column order
#    Get-EbiXlsxView -Book -Part   -> @{ tabSelected; topLeft; active }
#    Get-EbiXlsxDrawingItems -Book -Part
#                                  -> @(@{ kind; name; box; col; row; z;
#                                     line }) every pic / sp / cxnSp (group
#                                     members flattened), box in EMU with
#                                     its y from the CELL ANCHOR, not the
#                                     shape's own xfrm (K-42)
#    ConvertFrom-EbiCellRef 'B12'  -> @{ row = 12; col = 2 }
#    ConvertTo-EbiColumnName 2     -> 'B'
#
#  Every function but Read-EbiXlsx is pure over the parts table, so a test
#  can hand it XML strings.
# ============================================================

function Get-EbiXlsxNs {
    param([xml]$Doc)
    $ns = New-Object System.Xml.XmlNamespaceManager($Doc.NameTable)
    $ns.AddNamespace('m', 'http://schemas.openxmlformats.org/spreadsheetml/2006/main')
    $ns.AddNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships')
    $ns.AddNamespace('pr', 'http://schemas.openxmlformats.org/package/2006/relationships')
    $ns.AddNamespace('xdr', 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing')
    $ns.AddNamespace('a', 'http://schemas.openxmlformats.org/drawingml/2006/main')
    # the comma: a namespace manager is IEnumerable, and a bare return would
    # hand back its prefixes as an array of strings (10/28 office run:
    # 'Cannot find an overload for "SelectNodes" and the argument count: 2')
    return ,$ns
}

function Read-EbiXlsx {
    # -Media: also keep each xl/media/* part's length and first 64 KB
    # (book.media: part -> @{ length; head }) -- enough for a picture's
    # format and pixel size (Get-EbiXlsxPictures), never the whole picture.
    param([string]$Path, [switch]$Media)
    $out = @{ ok = $false; failure = ''; message = ''; path = $Path; parts = @{}; media = @{} }
    if (-not [System.IO.File]::Exists($Path)) { $out['failure'] = 'file_not_found'; $out['message'] = ('no file at ' + $Path); return $out }
    try {
        Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
        $fs = New-Object System.IO.FileStream($Path, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::ReadWrite)
        try {
            $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Read)
            try {
                foreach ($e in $zip.Entries) {
                    if ($Media -and $e.FullName -like 'xl/media/*') {
                        $st = $e.Open()
                        try {
                            $buf = New-Object byte[] ([int][Math]::Min([int64]65536, [int64]$e.Length))
                            $got = 0
                            while ($got -lt $buf.Length) { $n = $st.Read($buf, $got, $buf.Length - $got); if ($n -le 0) { break }; $got += $n }
                            $out['media'][$e.FullName] = @{ length = [int64]$e.Length; head = $buf }
                        } finally { $st.Dispose() }
                        continue
                    }
                    if ($e.FullName -notmatch '\.(xml|rels)$') { continue }
                    $sr = New-Object System.IO.StreamReader($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
                    try { $out['parts'][$e.FullName] = $sr.ReadToEnd() } finally { $sr.Dispose() }
                }
            } finally { $zip.Dispose() }
        } finally { $fs.Dispose() }
    } catch {
        $out['failure'] = 'read_failed'; $out['message'] = ('cannot read ' + $Path + ' as an .xlsx: ' + $_.Exception.Message); $out['parts'] = @{}; $out['media'] = @{}
        return $out
    }
    if (-not $out['parts'].Contains('xl/workbook.xml')) { $out['failure'] = 'read_failed'; $out['message'] = ($Path + ' has no xl/workbook.xml: not an .xlsx'); $out['parts'] = @{}; return $out }
    $out['ok'] = $true
    return $out
}

function Get-EbiXlsxXml {
    # PURE. A part as an XmlDocument, or $null when the book has no such part.
    param($Book, [string]$Part)
    if (-not $Book['parts'].Contains($Part)) { return $null }
    $d = New-Object System.Xml.XmlDocument
    $d.LoadXml([string]$Book['parts'][$Part])
    return ,$d    # never unrolled into its child nodes
}

function Resolve-EbiXlsxTarget {
    # PURE. A relationship target, relative to the part that owns it -> part name.
    param([string]$From, [string]$Target)
    if ($Target.StartsWith('/')) { return $Target.TrimStart('/') }
    $i = $From.LastIndexOf('/')
    $dir = if ($i -ge 0) { $From.Substring(0, $i) } else { '' }
    $segs = New-Object System.Collections.ArrayList
    foreach ($s in @(($dir + '/' + $Target) -split '/')) {
        if ($s -eq '' -or $s -eq '.') { continue }
        if ($s -eq '..') { if ($segs.Count -gt 0) { $segs.RemoveAt($segs.Count - 1) }; continue }
        [void]$segs.Add($s)
    }
    return ($segs.ToArray() -join '/')
}

function Get-EbiXlsxRels {
    # PURE. Id -> target part name, for the part's _rels/<name>.rels.
    param($Book, [string]$Part)
    $i = $Part.LastIndexOf('/')
    $rp = if ($i -ge 0) { $Part.Substring(0, $i) + '/_rels/' + $Part.Substring($i + 1) + '.rels' } else { '_rels/' + $Part + '.rels' }
    $out = @{}
    $d = Get-EbiXlsxXml -Book $Book -Part $rp
    if ($null -eq $d) { return $out }
    $ns = Get-EbiXlsxNs -Doc $d
    foreach ($r in $d.SelectNodes('/pr:Relationships/pr:Relationship', $ns)) {
        if ($r.GetAttribute('TargetMode') -eq 'External') { continue }
        $out[$r.GetAttribute('Id')] = Resolve-EbiXlsxTarget -From $Part -Target $r.GetAttribute('Target')
    }
    return $out
}

function Get-EbiXlsxSheets {
    param($Book)
    $d = Get-EbiXlsxXml -Book $Book -Part 'xl/workbook.xml'
    if ($null -eq $d) { return @() }
    $ns = Get-EbiXlsxNs -Doc $d
    $rel = Get-EbiXlsxRels -Book $Book -Part 'xl/workbook.xml'
    $rNs = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
    $out = New-Object System.Collections.ArrayList
    foreach ($s in $d.SelectNodes('/m:workbook/m:sheets/m:sheet', $ns)) {
        $id = $s.GetAttribute('id', $rNs)
        $state = $s.GetAttribute('state'); if ($state -eq '') { $state = 'visible' }
        [void]$out.Add(@{ name = $s.GetAttribute('name'); part = $(if ($rel.Contains($id)) { $rel[$id] } else { '' }); state = $state })
    }
    return $out.ToArray()
}

function Test-EbiXlsxOn {
    # PURE. A boolean property element: <b/>, <b val="1"/>, <b val="true"/>.
    param($Node)
    if ($null -eq $Node) { return $false }
    $v = $Node.GetAttribute('val')
    return ($v -eq '' -or $v -eq '1' -or $v -eq 'true')
}

function Get-EbiXlsxStyles {
    # PURE. xf[i] = @{ font; fill } indexes; fonts[i] = @{ name; bold };
    # fills[i] = @{ solid; rgb } (rgb '' for theme / indexed colours).
    param($Book)
    $out = @{ xf = @(); fonts = @(); fills = @() }
    $d = Get-EbiXlsxXml -Book $Book -Part 'xl/styles.xml'
    if ($null -eq $d) { return $out }
    $ns = Get-EbiXlsxNs -Doc $d
    $out['fonts'] = @(foreach ($f in $d.SelectNodes('/m:styleSheet/m:fonts/m:font', $ns)) {
        $n = $f.SelectSingleNode('m:name', $ns)
        @{ name = $(if ($null -ne $n) { $n.GetAttribute('val') } else { '' }); bold = (Test-EbiXlsxOn -Node $f.SelectSingleNode('m:b', $ns)) }
    })
    $out['fills'] = @(foreach ($f in $d.SelectNodes('/m:styleSheet/m:fills/m:fill', $ns)) {
        $p = $f.SelectSingleNode('m:patternFill', $ns)
        $solid = ($null -ne $p -and $p.GetAttribute('patternType') -eq 'solid')
        $fg = if ($null -ne $p) { $p.SelectSingleNode('m:fgColor', $ns) } else { $null }
        @{ solid = $solid; rgb = $(if ($null -ne $fg) { $fg.GetAttribute('rgb') } else { '' }) }
    })
    $out['xf'] = @(foreach ($x in $d.SelectNodes('/m:styleSheet/m:cellXfs/m:xf', $ns)) {
        @{ font = [int]('0' + $x.GetAttribute('fontId')); fill = [int]('0' + $x.GetAttribute('fillId')) }
    })
    return $out
}

function ConvertFrom-EbiCellRef {
    # PURE. 'AB12' -> @{ row = 12; col = 28 }.
    param([string]$Ref)
    $m = [regex]::Match($Ref, '^([A-Za-z]+)(\d+)$')
    if (-not $m.Success) { return @{ row = 0; col = 0 } }
    $c = 0
    foreach ($ch in $m.Groups[1].Value.ToUpperInvariant().ToCharArray()) { $c = $c * 26 + ([int]$ch - 64) }
    return @{ row = [int]$m.Groups[2].Value; col = $c }
}

function ConvertTo-EbiColumnName {
    # PURE. 1 -> 'A', 28 -> 'AB'.
    param([int]$Col)
    $s = ''
    $c = $Col
    while ($c -gt 0) { $r = ($c - 1) % 26; $s = [string][char](65 + $r) + $s; $c = [int][math]::Floor(($c - 1) / 26) }
    return $s
}

function Get-EbiXlsxSharedStrings {
    # PURE. The shared string table; phonetic runs (rPh) are not text.
    param($Book)
    $d = Get-EbiXlsxXml -Book $Book -Part 'xl/sharedStrings.xml'
    if ($null -eq $d) { return @() }
    $ns = Get-EbiXlsxNs -Doc $d
    return @(foreach ($si in $d.SelectNodes('/m:sst/m:si', $ns)) {
        $sb = New-Object System.Text.StringBuilder
        foreach ($t in $si.SelectNodes('m:t | m:r/m:t', $ns)) { [void]$sb.Append($t.InnerText) }
        $sb.ToString()
    })
}

function Get-EbiXlsxCells {
    <#
      PURE. Every <c> of the sheet part: @{ ref; row; col; text; hasValue;
      font; bold; yellow }. text is the shown string (shared / inline /
      cached value; a formula with no cached value shows its formula, as
      openpyxl does); hasValue is false for a cell that only carries a
      style. yellow = a solid fill whose rgb ends in FFFF00.
    #>
    param($Book, [string]$Part, $Styles, $Strings)
    $d = Get-EbiXlsxXml -Book $Book -Part $Part
    if ($null -eq $d) { return @() }
    if ($null -eq $Styles) { $Styles = Get-EbiXlsxStyles -Book $Book }
    if ($null -eq $Strings) { $Strings = @(Get-EbiXlsxSharedStrings -Book $Book) }
    $ns = Get-EbiXlsxNs -Doc $d
    $xf = @($Styles['xf']); $fonts = @($Styles['fonts']); $fills = @($Styles['fills'])
    $out = New-Object System.Collections.ArrayList
    foreach ($c in $d.SelectNodes('/m:worksheet/m:sheetData/m:row/m:c', $ns)) {
        $ref = $c.GetAttribute('r')
        $rc = ConvertFrom-EbiCellRef -Ref $ref
        $t = $c.GetAttribute('t')
        $v = $c.SelectSingleNode('m:v', $ns)
        $f = $c.SelectSingleNode('m:f', $ns)
        $text = $null
        if ($t -eq 'inlineStr') {
            $sb = New-Object System.Text.StringBuilder
            foreach ($x in $c.SelectNodes('m:is/m:t | m:is/m:r/m:t', $ns)) { [void]$sb.Append($x.InnerText) }
            $text = $sb.ToString()
        } elseif ($null -ne $v) {
            if ($t -eq 's') { $i = [int]$v.InnerText; $text = $(if ($i -ge 0 -and $i -lt $Strings.Count) { [string]$Strings[$i] } else { '' }) }
            else { $text = $v.InnerText }
        } elseif ($null -ne $f) {
            $text = '=' + $f.InnerText
        }
        $s = [int]('0' + $c.GetAttribute('s'))
        $fontName = ''; $bold = $false; $yellow = $false
        if ($s -lt $xf.Count) {
            $fi = [int]$xf[$s]['font']; $li = [int]$xf[$s]['fill']
            if ($fi -lt $fonts.Count) { $fontName = [string]$fonts[$fi]['name']; $bold = [bool]$fonts[$fi]['bold'] }
            if ($li -lt $fills.Count) { $yellow = ([bool]$fills[$li]['solid'] -and ([string]$fills[$li]['rgb']).ToUpperInvariant().EndsWith('FFFF00')) }
        }
        [void]$out.Add(@{ ref = $ref; row = [int]$rc['row']; col = [int]$rc['col']; text = $(if ($null -eq $text) { '' } else { $text }); hasValue = ($null -ne $text); font = $fontName; bold = $bold; yellow = $yellow })
    }
    return $out.ToArray()
}

function Get-EbiXlsxView {
    # PURE. The sheet's first sheetView: tabSelected, the top-left cell it
    # is scrolled to (the pane's when frozen), the active cell.
    param($Book, [string]$Part)
    $out = @{ tabSelected = $false; topLeft = 'A1'; active = ''; hasView = $false }
    $d = Get-EbiXlsxXml -Book $Book -Part $Part
    if ($null -eq $d) { return $out }
    $ns = Get-EbiXlsxNs -Doc $d
    $v = $d.SelectSingleNode('/m:worksheet/m:sheetViews/m:sheetView', $ns)
    if ($null -eq $v) { return $out }
    $out['hasView'] = $true
    $out['tabSelected'] = ($v.GetAttribute('tabSelected') -eq '1')
    $pane = $v.SelectSingleNode('m:pane', $ns)
    if ($null -ne $pane) { $out['topLeft'] = $pane.GetAttribute('topLeftCell') }
    elseif ($v.GetAttribute('topLeftCell') -ne '') { $out['topLeft'] = $v.GetAttribute('topLeftCell') }
    foreach ($sel in $v.SelectNodes('m:selection', $ns)) {
        $p = $sel.GetAttribute('pane')
        if ($p -eq '' -or $p -eq 'bottomLeft') { $out['active'] = $sel.GetAttribute('activeCell') }
    }
    return $out
}

function Get-EbiXlsxRowTops {
    # PURE. -> scriptblock-free helper data: @{ default; heights } (points).
    param($Book, [string]$Part)
    $d = Get-EbiXlsxXml -Book $Book -Part $Part
    $out = @{ default = 15.0; heights = @{} }
    if ($null -eq $d) { return $out }
    $ns = Get-EbiXlsxNs -Doc $d
    $fmt = $d.SelectSingleNode('/m:worksheet/m:sheetFormatPr', $ns)
    if ($null -ne $fmt -and $fmt.GetAttribute('defaultRowHeight') -ne '') { $out['default'] = [double]::Parse($fmt.GetAttribute('defaultRowHeight'), [System.Globalization.CultureInfo]::InvariantCulture) }
    foreach ($r in $d.SelectNodes('/m:worksheet/m:sheetData/m:row', $ns)) {
        $n = [int]$r.GetAttribute('r')
        if ($r.GetAttribute('hidden') -eq '1') { $out['heights'][$n] = 0.0 }
        elseif ($r.GetAttribute('ht') -ne '') { $out['heights'][$n] = [double]::Parse($r.GetAttribute('ht'), [System.Globalization.CultureInfo]::InvariantCulture) }
    }
    return $out
}

function Get-EbiXlsxRowY {
    # PURE. EMU from the sheet top to the top of 0-based row Row0, plus Off.
    param($Rows, [int]$Row0, [double]$Off)
    $pt = 0.0
    for ($r = 1; $r -le $Row0; $r++) { $pt += $(if ($Rows['heights'].Contains($r)) { [double]$Rows['heights'][$r] } else { [double]$Rows['default'] }) }
    return ($pt * 12700.0 + $Off)
}

function Add-EbiXlsxShapes {
    # PURE. Walk pic / sp / cxnSp / grpSp under Node into Out; Tf maps a
    # child's xfrm coordinates to its parent's: @(ox, oy, sx, sy, chx, chy).
    param($Node, $Ns, $Tf, $Out)
    foreach ($ch in $Node.ChildNodes) {
        if ($ch.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        $t = $ch.LocalName
        if (@('pic', 'sp', 'grpSp', 'cxnSp') -notcontains $t) { continue }
        $pr = if ($t -eq 'grpSp') { $ch.SelectSingleNode('xdr:grpSpPr', $Ns) } else { $ch.SelectSingleNode('xdr:spPr', $Ns) }
        $x = if ($null -ne $pr) { $pr.SelectSingleNode('a:xfrm', $Ns) } else { $null }
        if ($null -eq $x) { continue }
        $o = $x.SelectSingleNode('a:off', $Ns); $e = $x.SelectSingleNode('a:ext', $Ns)
        if ($null -eq $o -or $null -eq $e) { continue }
        $X = [double]$o.GetAttribute('x'); $Y = [double]$o.GetAttribute('y'); $W = [double]$e.GetAttribute('cx'); $H = [double]$e.GetAttribute('cy')
        $x0 = $Tf[0] + ($X - $Tf[4]) * $Tf[2]; $y0 = $Tf[1] + ($Y - $Tf[5]) * $Tf[3]
        $x1 = $Tf[0] + ($X + $W - $Tf[4]) * $Tf[2]; $y1 = $Tf[1] + ($Y + $H - $Tf[5]) * $Tf[3]
        if ($t -eq 'grpSp') {
            $co = $x.SelectSingleNode('a:chOff', $Ns); $ce = $x.SelectSingleNode('a:chExt', $Ns)
            if ($null -eq $co -or $null -eq $ce) { continue }
            $sx = ($x1 - $x0) / [math]::Max(1.0, [double]$ce.GetAttribute('cx'))
            $sy = ($y1 - $y0) / [math]::Max(1.0, [double]$ce.GetAttribute('cy'))
            Add-EbiXlsxShapes -Node $ch -Ns $Ns -Tf @($x0, $y0, $sx, $sy, [double]$co.GetAttribute('x'), [double]$co.GetAttribute('y')) -Out $Out
            continue
        }
        $nv = $ch.SelectSingleNode('.//xdr:cNvPr', $Ns)
        $line = ''
        if ($t -ne 'pic') {
            $ln = $ch.SelectSingleNode('xdr:spPr/a:ln', $Ns)
            if ($null -ne $ln) {
                $c = $ln.SelectSingleNode('.//a:srgbClr', $Ns)
                if ($null -ne $c) { $line = $c.GetAttribute('val') } elseif ($null -ne $ln.SelectSingleNode('a:noFill', $Ns)) { $line = 'none' }
            }
        }
        [void]$Out.Add(@{ kind = $t; name = $(if ($null -ne $nv) { $nv.GetAttribute('name') } else { '' }); box = @($x0, $y0, $x1, $y1); line = $line; col = -1; row = -1; z = 0 })
    }
}

function Get-EbiXlsxDrawingItems {
    # PURE. Every picture / shape on the sheet (see the header).
    param($Book, [string]$Part)
    $rows = Get-EbiXlsxRowTops -Book $Book -Part $Part
    $out = New-Object System.Collections.ArrayList
    $rels = Get-EbiXlsxRels -Book $Book -Part $Part
    foreach ($dp in @($rels.Values | Where-Object { $_ -match '/drawings/[^/]+\.xml$' } | Sort-Object)) {
        $d = Get-EbiXlsxXml -Book $Book -Part $dp
        if ($null -eq $d) { continue }
        $ns = Get-EbiXlsxNs -Doc $d
        $z = -1
        foreach ($anc in $d.DocumentElement.ChildNodes) {
            if ($anc.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            $z++
            $got = New-Object System.Collections.ArrayList
            Add-EbiXlsxShapes -Node $anc -Ns $ns -Tf @(0.0, 0.0, 1.0, 1.0, 0.0, 0.0) -Out $got
            if ($got.Count -eq 0) { continue }
            foreach ($it in $got) { $it['z'] = $z }
            $f = $anc.SelectSingleNode('xdr:from', $ns)
            if ($null -ne $f) {
                # the cell anchor is the truth for y (the xfrm goes stale when rows move)
                $col0 = [int]$f.SelectSingleNode('xdr:col', $ns).InnerText; $row0 = [int]$f.SelectSingleNode('xdr:row', $ns).InnerText
                $ay0 = Get-EbiXlsxRowY -Rows $rows -Row0 $row0 -Off ([double]$f.SelectSingleNode('xdr:rowOff', $ns).InnerText)
                $to = $anc.SelectSingleNode('xdr:to', $ns)
                if ($null -ne $to) { $ay1 = Get-EbiXlsxRowY -Rows $rows -Row0 ([int]$to.SelectSingleNode('xdr:row', $ns).InnerText) -Off ([double]$to.SelectSingleNode('xdr:rowOff', $ns).InnerText) }
                else { $ext = $anc.SelectSingleNode('xdr:ext', $ns); $ay1 = $ay0 + $(if ($null -ne $ext) { [double]$ext.GetAttribute('cy') } else { 0.0 }) }
                $top = ($got | ForEach-Object { [double]$_['box'][1] } | Measure-Object -Minimum).Minimum
                $bot = ($got | ForEach-Object { [double]$_['box'][3] } | Measure-Object -Maximum).Maximum
                $sc = ($ay1 - $ay0) / [math]::Max(1.0, $bot - $top)
                foreach ($it in $got) {
                    $b = $it['box']
                    $it['box'] = @([double]$b[0], ($ay0 + ([double]$b[1] - $top) * $sc), [double]$b[2], ($ay0 + ([double]$b[3] - $top) * $sc))
                    $it['col'] = $col0; $it['row'] = $row0
                }
            }
            foreach ($it in $got) { [void]$out.Add($it) }
        }
    }
    return $out.ToArray()
}

function Add-EbiXlsxPictures {
    # PURE. Walk pic / grpSp under Node into Out: SX / SY scale a child's
    # EMU to the sheet's (a group's ext over its chExt); Group = the
    # outermost group's name ('' at the top).
    param($Node, $Ns, [double]$SX, [double]$SY, [string]$Group, $Rels, $Out)
    foreach ($ch in $Node.ChildNodes) {
        if ($ch.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
        if ($ch.LocalName -eq 'grpSp') {
            $x = $ch.SelectSingleNode('xdr:grpSpPr/a:xfrm', $Ns)
            # gx / gy, not sx / sy: variable names are case-blind, and $sx
            # would BE the parameter $SX -- the later siblings' scale
            $gx = $SX; $gy = $SY
            if ($null -ne $x) {
                $e = $x.SelectSingleNode('a:ext', $Ns); $ce = $x.SelectSingleNode('a:chExt', $Ns)
                if ($null -ne $e -and $null -ne $ce -and [double]$ce.GetAttribute('cx') -gt 0 -and [double]$ce.GetAttribute('cy') -gt 0) {
                    $gx = $SX * [double]$e.GetAttribute('cx') / [double]$ce.GetAttribute('cx')
                    $gy = $SY * [double]$e.GetAttribute('cy') / [double]$ce.GetAttribute('cy')
                }
            }
            $nv = $ch.SelectSingleNode('xdr:nvGrpSpPr/xdr:cNvPr', $Ns)
            $g = if ($Group -ne '') { $Group } elseif ($null -ne $nv) { $nv.GetAttribute('name') } else { '' }
            Add-EbiXlsxPictures -Node $ch -Ns $Ns -SX $gx -SY $gy -Group $g -Rels $Rels -Out $Out
            continue
        }
        if ($ch.LocalName -ne 'pic') { continue }
        $nv = $ch.SelectSingleNode('xdr:nvPicPr/xdr:cNvPr', $Ns)
        $blip = $ch.SelectSingleNode('xdr:blipFill/a:blip', $Ns)
        $emb = if ($null -ne $blip) { $blip.GetAttribute('embed', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships') } else { '' }
        $e = $ch.SelectSingleNode('xdr:spPr/a:xfrm/a:ext', $Ns)
        $src = $ch.SelectSingleNode('xdr:blipFill/a:srcRect', $Ns)
        $crop = @{ l = 0; t = 0; r = 0; b = 0 }
        if ($null -ne $src) { foreach ($k in @('l', 't', 'r', 'b')) { $v = $src.GetAttribute($k); if ($v -ne '') { $crop[$k] = [int]$v } } }
        [void]$Out.Add(@{ name = $(if ($null -ne $nv) { $nv.GetAttribute('name') } else { '' }); group = $Group;
                          media = $(if ($emb -ne '' -and $Rels.Contains($emb)) { [string]$Rels[$emb] } else { '' });
                          shownW = $(if ($null -ne $e) { [Math]::Round([double]$e.GetAttribute('cx') * $SX / 9525.0, 1) } else { 0 });
                          shownH = $(if ($null -ne $e) { [Math]::Round([double]$e.GetAttribute('cy') * $SY / 9525.0, 1) } else { 0 });
                          crop = $crop })
    }
}

function Get-EbiXlsxPictures {
    # PURE. Every picture on the sheet (see the header), drawing order.
    param($Book, [string]$Part)
    $out = New-Object System.Collections.ArrayList
    $rels = Get-EbiXlsxRels -Book $Book -Part $Part
    foreach ($dp in @($rels.Values | Where-Object { $_ -match '/drawings/[^/]+\.xml$' } | Sort-Object)) {
        $d = Get-EbiXlsxXml -Book $Book -Part $dp
        if ($null -eq $d) { continue }
        $ns = Get-EbiXlsxNs -Doc $d
        $dr = Get-EbiXlsxRels -Book $Book -Part $dp
        foreach ($anc in $d.DocumentElement.ChildNodes) {
            if ($anc.NodeType -ne [System.Xml.XmlNodeType]::Element) { continue }
            Add-EbiXlsxPictures -Node $anc -Ns $ns -SX 1.0 -SY 1.0 -Group '' -Rels $dr -Out $out
        }
    }
    return $out.ToArray()
}
