$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'kernel/BookReview.ps1')

Reset-Tests 'BookReview'

# ---------------------------------------------------------------- fixtures
# A tiny evidence book as XML parts (the review reads parts, not Excel).
# ASCII sheet names / font: the checks take every name from $Names.
$M = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main'
$R = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'
$P = 'http://schemas.openxmlformats.org/package/2006/relationships'
$X = 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
$A = 'http://schemas.openxmlformats.org/drawingml/2006/main'

function Esc { param([string]$S) return [System.Security.SecurityElement]::Escape($S) }

function New-SheetXml {
    # Cells: @(@('B5', 'text', styleIndex), ...); View: @{ sel; top; act; pane }
    param($Cells, $View, [bool]$Drawing = $false, $Heights = @{})
    $byRow = @{}
    foreach ($c in $Cells) { $rc = ConvertFrom-EbiCellRef -Ref $c[0]; $rowNo = [int]$rc['row']; if (-not $byRow.Contains($rowNo)) { $byRow[$rowNo] = New-Object System.Collections.ArrayList }; [void]$byRow[$rowNo].Add($c) }
    foreach ($hk in $Heights.Keys) { if (-not $byRow.Contains([int]$hk)) { $byRow[[int]$hk] = New-Object System.Collections.ArrayList } }
    # $rowNo / $hk, never $r: variable names are case-blind, and a local
    # $r hides the script's $R (the relationships namespace) -> xmlns:r=""
    $sb = New-Object System.Text.StringBuilder
    [void]$sb.Append('<worksheet xmlns="' + $M + '" xmlns:r="' + $R + '">')
    [void]$sb.Append('<sheetViews><sheetView workbookViewId="0"' + $(if ($View['sel']) { ' tabSelected="1"' } else { '' }) + $(if ($View['top']) { ' topLeftCell="' + $View['top'] + '"' } else { '' }) + '>')
    if ($View['pane']) { [void]$sb.Append('<pane ySplit="2" topLeftCell="' + $View['pane'] + '" activePane="bottomLeft" state="frozen"/>') }
    [void]$sb.Append('<selection' + $(if ($View['pane']) { ' pane="bottomLeft"' } else { '' }) + ' activeCell="' + $View['act'] + '" sqref="' + $View['act'] + '"/></sheetView></sheetViews>')
    [void]$sb.Append('<sheetFormatPr defaultRowHeight="15"/><sheetData>')
    foreach ($rowNo in @($byRow.Keys | Sort-Object)) {
        [void]$sb.Append('<row r="' + $rowNo + '"' + $(if ($Heights.Contains($rowNo)) { ' ht="' + $Heights[$rowNo] + '" customHeight="1"' } else { '' }) + '>')
        foreach ($c in $byRow[$rowNo]) {
            if ($null -eq $c[1]) { [void]$sb.Append('<c r="' + $c[0] + '" s="' + $c[2] + '"/>') }
            else { [void]$sb.Append('<c r="' + $c[0] + '" s="' + $c[2] + '" t="inlineStr"><is><t xml:space="preserve">' + (Esc $c[1]) + '</t></is></c>') }
        }
        [void]$sb.Append('</row>')
    }
    [void]$sb.Append('</sheetData>')
    if ($Drawing) { [void]$sb.Append('<drawing r:id="rId1"/>') }
    [void]$sb.Append('</worksheet>')
    return $sb.ToString()
}

function New-Anchor {
    # One twoCellAnchor holding a pic or a red sp, rows 0-based like the file.
    param([string]$Kind, [string]$Name, [int]$Col, [int]$Row, [int]$ToRow, [long]$X0, [long]$W, [int]$Id)
    $shape = if ($Kind -eq 'pic') {
        '<xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + $Id + '" name="' + (Esc $Name) + '"/><xdr:cNvPicPr/></xdr:nvPicPr><xdr:blipFill/><xdr:spPr><a:xfrm><a:off x="' + $X0 + '" y="0"/><a:ext cx="' + $W + '" cy="1000"/></a:xfrm></xdr:spPr></xdr:pic>'
    } else {
        '<xdr:sp><xdr:nvSpPr><xdr:cNvPr id="' + $Id + '" name="' + (Esc $Name) + '"/><xdr:cNvSpPr/></xdr:nvSpPr><xdr:spPr><a:xfrm><a:off x="' + $X0 + '" y="0"/><a:ext cx="' + $W + '" cy="1000"/></a:xfrm><a:ln><a:solidFill><a:srgbClr val="FF0000"/></a:solidFill></a:ln></xdr:spPr></xdr:sp>'
    }
    return '<xdr:twoCellAnchor><xdr:from><xdr:col>' + $Col + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $Row + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from><xdr:to><xdr:col>' + ($Col + 10) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + $ToRow + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>' + $shape + '<xdr:clientData/></xdr:twoCellAnchor>'
}

$names = @{ send = 'SEND'; giftRecv = 'GIFTRECV'; gfixRecv = 'GFIXRECV'; compare = 'CMP'; font = 'FontX'; jobLog = 'JOBLOG' }

$parts = @{}
$parts['xl/workbook.xml'] = '<workbook xmlns="' + $M + '" xmlns:r="' + $R + '"><sheets>' +
    '<sheet name="SEND" sheetId="1" r:id="rId1"/><sheet name="GIFTRECV" sheetId="2" r:id="rId2"/>' +
    '<sheet name="GFIXRECV" sheetId="3" r:id="rId3"/><sheet name="CMP" sheetId="4" r:id="rId4"/>' +
    '<sheet name="OLD" sheetId="5" state="hidden" r:id="rId5"/></sheets></workbook>'
$parts['xl/_rels/workbook.xml.rels'] = '<Relationships xmlns="' + $P + '">' +
    '<Relationship Id="rId1" Type="t" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="t" Target="worksheets/sheet2.xml"/>' +
    '<Relationship Id="rId3" Type="t" Target="/xl/worksheets/sheet3.xml"/><Relationship Id="rId4" Type="t" Target="worksheets/sheet4.xml"/>' +
    '<Relationship Id="rId5" Type="t" Target="worksheets/sheet5.xml"/><Relationship Id="rId9" Type="t" Target="styles.xml"/></Relationships>'
# fonts: 0 FontX, 1 FontX bold, 2 Arial; fills: 2 = solid yellow; xf: 0 plain, 1 bold, 2 Arial, 3 yellow
$parts['xl/styles.xml'] = '<styleSheet xmlns="' + $M + '">' +
    '<fonts count="3"><font><name val="FontX"/></font><font><b/><name val="FontX"/></font><font><name val="Arial"/></font></fonts>' +
    '<fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FFFFFF00"/></patternFill></fill></fills>' +
    '<cellXfs count="4"><xf fontId="0" fillId="0"/><xf fontId="1" fillId="0"/><xf fontId="2" fillId="0"/><xf fontId="0" fillId="2"/></cellXfs></styleSheet>'

# SEND: first sheet, selected, at the top; two file groups
$parts['xl/worksheets/sheet1.xml'] = New-SheetXml -Cells @(@('A3', 'JJPCRS12', 0), @('B3', 'STEP:MQFX0020', 0), @('A30', 'JJPCRS12', 0), @('B30', 'STEP:MQFX0030', 0)) -View @{ sel = $true; act = 'A3' }
# GIFTRECV: also selected (NG); a good block, then a block with an extra
# line after its END + date line, an END with no START; yellow on an empty row
$parts['xl/worksheets/sheet2.xml'] = New-SheetXml -View @{ sel = $true; act = 'A3' } -Cells @(
    @('B5', 'GIFTReceive.log', 1), @('B6', 'GIFTReceiver START F1', 0), @('B7', 'file stored : [/r/F1.csv]', 3), @('C7', $null, 3), @('D7', $null, 3), @('E7', $null, 3), @('F7', $null, 3), @('G7', $null, 3), @('H7', $null, 3), @('I7', $null, 3), @('J7', $null, 3),
    @('B8', 'GIFTReceiver END F1', 0), @('B9', '2026/09/28', 0),
    @('B12', 'GIFTReceive.log', 1), @('B13', 'GIFTReceiver START F2', 0), @('B14', 'GIFTReceiver END F2', 0), @('B15', 'GIFTReceiver END F3', 0), @('B16', '2026/09/28', 0), @('B17', 'stray line', 0),
    @('C20', $null, 3), @('D20', $null, 3))
# GFIXRECV: scrolled down (pane), the receive label not bold and naming the
# wrong log, an Arial cell, two job logs (second Renaming not yellow)
$parts['xl/worksheets/sheet3.xml'] = New-SheetXml -View @{ pane = 'A40'; act = 'A3' } -Cells @(
    @('B8', 'JOBLOG', 1),
    @('B9', 'Start Date and Time 2026-10-26', 0), @('B10', "INFO      Uploading '/x/a' (107,520 bytes) to '/y'", 0), @('B11', "INFO      Renaming '/y/a'", 3),
    @('B12', 'Start Date and Time 2026-10-26', 0), @('B13', "INFO      Uploading '/x/b' (99 bytes) to '/y'", 0), @('B14', "INFO      Renaming '/y/b'", 0),
    @('B20', 'GFIXReceive.log', 0), @('B21', 'GFIXReceiverUnzip START Z1', 2), @('B22', 'GFIXReceiverUnzip END Z1', 0), @('B23', '2026/10/26', 0))
# CMP: set labels at 3 (G right) and 32 (G wrong); a stale label at 60;
# a hand picture pasted over the tool's picture of set 2; a red box on nothing
$parts['xl/worksheets/sheet4.xml'] = New-SheetXml -View @{ act = 'A1' } -Drawing $true -Cells @(
    @('B3', 'JJPCRS12', 0), @('G3', 'STEP:MQFX0020', 0), @('B32', 'JJPCRS12', 0), @('G32', 'STEP:OTHER', 0), @('B60', 'JJPCRS12', 0), @('G60', 'STEP:MQFX0030', 0))
$parts['xl/worksheets/_rels/sheet4.xml.rels'] = '<Relationships xmlns="' + $P + '"><Relationship Id="rId1" Type="d" Target="../drawings/drawing1.xml"/></Relationships>'
$parts['xl/drawings/drawing1.xml'] = '<xdr:wsDr xmlns:xdr="' + $X + '" xmlns:a="' + $A + '">' +
    (New-Anchor -Kind 'pic' -Name 'verifyPic_df_1' -Col 1 -Row 3 -ToRow 25 -X0 100000 -W 9000000 -Id 2) +
    (New-Anchor -Kind 'sp' -Name 'verifyMark_df_1_1' -Col 3 -Row 20 -ToRow 21 -X0 400000 -W 500000 -Id 3) +
    (New-Anchor -Kind 'pic' -Name 'verifyPic_df_2' -Col 1 -Row 32 -ToRow 50 -X0 100000 -W 9000000 -Id 4) +
    (New-Anchor -Kind 'pic' -Name 'Picture 1' -Col 1 -Row 32 -ToRow 52 -X0 100000 -W 9000000 -Id 5) +
    (New-Anchor -Kind 'sp' -Name 'leftover' -Col 30 -Row 200 -ToRow 201 -X0 90000000 -W 500000 -Id 6) +
    '</xdr:wsDr>'
$parts['xl/worksheets/sheet5.xml'] = New-SheetXml -View @{ sel = $true; act = 'Z99' } -Cells @()
$book = @{ ok = $true; parts = $parts; path = 'fixture.xlsx' }

# ---------------------------------------------------------------- Xlsx reader
Assert-Equal 'xl/drawings/drawing1.xml' (Resolve-EbiXlsxTarget -From 'xl/worksheets/sheet4.xml' -Target '../drawings/drawing1.xml') 'xlsx: a relative target resolves against its part'
Assert-Equal 'xl/worksheets/sheet3.xml' (Resolve-EbiXlsxTarget -From 'xl/workbook.xml' -Target '/xl/worksheets/sheet3.xml') 'xlsx: an absolute target drops the leading slash'
$rc = ConvertFrom-EbiCellRef -Ref 'AB12'
Assert-Equal '12|28' ('' + $rc['row'] + '|' + $rc['col']) 'xlsx: AB12 -> row 12, column 28'
Assert-Equal 'AB' (ConvertTo-EbiColumnName -Col 28) 'xlsx: column 28 -> AB'
$sheets = @(Get-EbiXlsxSheets -Book $book)
Assert-Equal 'SEND,GIFTRECV,GFIXRECV,CMP,OLD' (@($sheets | ForEach-Object { $_['name'] }) -join ',') 'xlsx: the sheets in tab order'
Assert-Equal 'hidden' $sheets[4]['state'] 'xlsx: a hidden sheet says so'
$styles = Get-EbiXlsxStyles -Book $book
$cells = @(Get-EbiXlsxCells -Book $book -Part 'xl/worksheets/sheet2.xml' -Styles $styles -Strings @())
$b5 = @($cells | Where-Object { $_['ref'] -eq 'B5' })[0]
Assert-True ($b5['bold'] -and $b5['font'] -eq 'FontX' -and $b5['text'] -eq 'GIFTReceive.log') 'xlsx: a bold label cell reads with its font and text'
$c20 = @($cells | Where-Object { $_['ref'] -eq 'C20' })[0]
Assert-True ($c20['yellow'] -and -not $c20['hasValue']) 'xlsx: a style-only yellow cell has no value'
$items = @(Get-EbiXlsxDrawingItems -Book $book -Part 'xl/worksheets/sheet4.xml')
Assert-Equal 5 $items.Count 'xlsx: five drawing items'
$p1 = @($items | Where-Object { $_['name'] -eq 'verifyPic_df_1' })[0]
Assert-True ($p1['col'] -eq 1 -and $p1['row'] -eq 3 -and [math]::Abs([double]$p1['box'][1] - 3 * 15 * 12700) -lt 1) 'xlsx: y comes from the cell anchor (3 rows of 15 pt)'

# shared strings: rich runs join, phonetic runs do not
$ss = @{ parts = @{ 'xl/sharedStrings.xml' = '<sst xmlns="' + $M + '"><si><t>plain</t></si><si><r><t>ri</t></r><r><t>ch</t></r><rPh><t>PH</t></rPh></si></sst>' } }
Assert-Equal 'plain|rich' (@(Get-EbiXlsxSharedStrings -Book $ss) -join '|') 'xlsx: shared strings, rich runs joined, phonetic text left out'

# ---------------------------------------------------------------- the review
$model = New-EbiBookModel -Book $book -Names $names
$f = @(Invoke-EbiBookReview -Model $model -Names $names)
$lines = @($f | ForEach-Object { $_['level'] + ' ' + $_['where'] + ': ' + $_['message'] })
$all = $lines -join "`n"
function Has { param([string]$Like) return (@($lines | Where-Object { $_ -like $Like }).Count -gt 0) }

Assert-True (Has 'NG GIFTRECV: this sheet is selected*') 'review view: a second selected sheet is NG'
Assert-True (Has 'NG GFIXRECV: scrolled to A40*') 'review view: a sheet scrolled down (frozen pane) is NG'
Assert-True (-not (Has '*OLD*')) 'review view: a hidden sheet is not checked'
Assert-True (Has 'NG GFIXRECV!B21: font "Arial"*') 'review fonts: a cell in another font is NG'
Assert-True (Has 'NG GIFTRECV!B17: 1 line(s) after the last END + date line*') 'review logs: a line after END + date is NG (K-37)'
Assert-True (Has 'NG GIFTRECV!B15: END F3 has no START line*') 'review logs: an END without its START is NG (K-37)'
Assert-True (-not (Has '*GIFTRECV!B5*')) 'review logs: the good block has no finding'
Assert-True (Has 'NG GFIXRECV!B20: log label "GFIXReceive.log" is not bold*') 'review logs: a label not bold is NG'
Assert-True (Has 'NG GFIXRECV!B20: label "GFIXReceive.log" but the lines are GFIXReceiverUnzip -> "GFIXReceiveUnzip.log"*') 'review logs: a label naming the wrong log is NG (K-36)'
Assert-True (Has 'NG GIFTRECV!C20: yellow fill on an empty row*') 'review yellow: yellow on an empty row is NG (K-38)'
Assert-True (-not (Has '*GIFTRECV!B7*')) 'review yellow: yellow over its text is fine'
Assert-True (Has 'NG CMP!G32: G32 is "STEP:OTHER", SEND says "STEP:MQFX0030"*') 'review labels: a wrong STEP name is NG (K-39)'
Assert-True (Has 'NG CMP!B60: B60/G60: a label on no set''s row*') 'review labels: a label where no set starts is NG (K-48)'
Assert-True (-not (Has '*label row(s) but*')) 'review labels: two sets, two groups'
Assert-True (Has 'NG GFIXRECV!B14: JOBLOG: the Renaming line is not yellow*') 'review job logs: a Renaming line not yellow is NG (K-45)'
Assert-True (-not (Has 'NG GFIXRECV!B11*')) 'review job logs: a yellow Renaming line is not NG'
Assert-True (Has 'WARN GFIXRECV!B11: yellow stops before the text ends*') 'review yellow: yellow on one column of a long line is a WARN, not NG'
Assert-True (Has 'NG CMP!B33: "Picture 1" is pasted over "verifyPic_df_2"*') 'review shapes: a hand picture over a tool picture is NG (K-49)'
Assert-True (Has 'NG CMP: red box "leftover" is not on any picture*') 'review shapes: a red box on nothing is NG'
Assert-True (-not (Has '*verifyMark_df_1_1*')) 'review shapes: a red box on a picture is fine'

# job-log order (K-44): set 2 is the empty one (one red box), the 99-byte log is second -> fine;
# swap the sizes and it is NG
$model2 = New-EbiBookModel -Book $book -Names $names
$jb = @(Get-EbiBookJobLogBlocks -Model $model2 -Names $names)
Assert-Equal '107520,99' (@($jb | ForEach-Object { $_['bytes'] }) -join ',') 'review job logs: the upload sizes are read per job log'

# ---------------------------------------------------------------- the step
. (Join-Path (Split-Path $here -Parent) 'modules/verify/verify.book_review.ps1')
$v = VerifyBookReview-Verdict -Findings $f -MaxAlert 3 -Name 'b.xlsx'
Assert-Equal 'ng' $v['code'] 'step verdict: any NG -> ng'
Assert-Equal 4 @($v['alert']).Count 'step verdict: maxAlert lines + one "... and N more"'
Assert-True (([string]@($v['alert'])[3]).StartsWith('... and ')) 'step verdict: ... counting the rest'
$w = VerifyBookReview-Verdict -Findings @(@{ level = 'WARN'; where = 'S!B1'; message = 'm'; ref = '' }) -MaxAlert 8 -Name 'b.xlsx'
Assert-True ($w['code'] -eq 'ok' -and $w['warn'] -eq 1) 'step verdict: WARN only -> ok'

# Read-EbiXlsx round trip through a real zip
$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-book-' + [guid]::NewGuid().ToString('N').Substring(0, 8) + '.xlsx')
try {
    Add-Type -AssemblyName System.IO.Compression -ErrorAction SilentlyContinue
    $fs = [System.IO.File]::Open($tmp, [System.IO.FileMode]::CreateNew)
    $zip = New-Object System.IO.Compression.ZipArchive($fs, [System.IO.Compression.ZipArchiveMode]::Create)
    foreach ($k in $parts.Keys) {
        $e = $zip.CreateEntry($k)
        $sw = New-Object System.IO.StreamWriter($e.Open(), (New-Object System.Text.UTF8Encoding($false)))
        $sw.Write([string]$parts[$k]); $sw.Dispose()
    }
    $zip.Dispose(); $fs.Dispose()
    $rb = Read-EbiXlsx -Path $tmp
    Assert-True ($rb['ok'] -and $rb['parts'].Count -eq $parts.Count) 'xlsx: a zip on disk reads back every part'
    $log = New-Object PSObject
    $log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) }
    $ctx = @{ WorkDir = [System.IO.Path]::GetTempPath(); DryRun = $false; Log = $log }
    $sheetsIn = @{ sendData = 'SEND'; giftResult = 'GIFTRECV'; gfixResult = 'GFIXRECV'; compare = 'CMP' }
    $r = Invoke-Step -In @{ path = $tmp; sheets = $sheetsIn; font = 'FontX'; jobLog = 'JOBLOG'; maxAlert = 8 } -Ctx $ctx
    Assert-True ($r['ok'] -and $r['code'] -eq 'ng' -and $r['ng'] -eq @($f | Where-Object { $_['level'] -eq 'NG' }).Count) ('step: the zip reviewed like the parts (' + $r['ng'] + ' NG)')
    $r2 = Invoke-Step -In @{ path = ($tmp + '.none'); sheets = $sheetsIn; font = 'FontX'; jobLog = 'JOBLOG'; maxAlert = 8 } -Ctx $ctx
    Assert-True ((-not $r2['ok']) -and $r2['failure'] -eq 'file_not_found') 'step: no book -> file_not_found'
    $r3 = Invoke-Step -In @{ path = $tmp; sheets = @{ sendData = 'SEND' }; font = 'FontX'; jobLog = 'JOBLOG'; maxAlert = 8 } -Ctx $ctx
    Assert-True ((-not $r3['ok']) -and $r3['failure'] -eq 'input_invalid') 'step: a sheets map without giftResult -> input_invalid'
} finally {
    Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue
}

Complete-Tests
