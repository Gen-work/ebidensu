#Requires -Version 5.1
# ============================================================
#  kernel/Layout.ps1
#
#  PURE geometry shared by the screen.* / excel.* steps: where a run of
#  list rows sits in a window, where a list entry sits in a screenshot of
#  a list scrolled to its end, which pixel rows of a strip carry ink, how
#  a rectangle drawn in a picture's pixels lands on an Excel sheet, and how
#  a set of captures stacks onto a sheet. Dot-source only (no param(),
#  ASCII source). No Windows type is named here; the one impure need
#  (decoding a PNG into pixels) lives in kernel/Image.ps1.
#
#    Get-EbiRowRegion        header + rows [first..first+count) at a pitch
#    Get-EbiListEntryRects   list entries -> rects, counted from the last row
#    Get-EbiInkBands         per-row ink counts -> contiguous bands
#    Resolve-EbiAnchoredRect a rect anchored to an image edge -> absolute px
#    ConvertTo-EbiSheetRect  image px rect -> sheet points over a picture
#    New-EbiStackPlan        capture pairs -> an excel.insert_pictures list
# ============================================================

function Get-EbiRowRegion {
    <#
      PURE. The rectangle covering a fixed header band plus a run of list
      rows. All numbers are pixels relative to the window's top-left, then
      offset by the window's own position (OriginX/OriginY) so the result is
      in screen pixels.
        Top:        top edge of the region (above the header, e.g. the panel
                    title)
        FirstRowTop:top edge of list row 1
        RowHeight:  row pitch
        FirstIndex: 1-based index of the first row to include
        Count:      rows to include
      Rows above FirstIndex are NOT included -- the region is header +
      the chosen rows, so when the chosen rows are not at the top of the
      list the region is cut into: header band [Top..FirstRowTop) and the
      rows band. A single rectangle cannot skip rows, so in that case the
      caller gets skippedRows > 0 and the region still spans from Top down
      to the last chosen row (rows in between are visible).
      Returns @{ x; y; width; height; skippedRows }.
    #>
    param([int]$OriginX = 0, [int]$OriginY = 0, [int]$Left, [int]$Right, [int]$Top, [int]$FirstRowTop, [double]$RowHeight, [int]$FirstIndex = 1, [int]$Count = 1, [int]$BottomPad = 0)
    if ($FirstIndex -lt 1) { $FirstIndex = 1 }
    if ($Count -lt 1) { $Count = 1 }
    $lastBottom = $FirstRowTop + [int][Math]::Round(($FirstIndex - 1 + $Count) * $RowHeight) + $BottomPad
    return @{ x = $OriginX + $Left; y = $OriginY + $Top; width = ($Right - $Left); height = ($lastBottom - $Top); skippedRows = ($FirstIndex - 1) }
}

function Get-EbiListEntryRects {
    <#
      PURE. A list page scrolled to its END shows its last row at a fixed
      place (LastRowCenterY), and each earlier row RowPitch above it. Given
      the page's entries in order (names) and the targets, return one rect
      per CONTIGUOUS run of target rows (three files in a row get one box,
      the way the operator draws it):
        @{ x; y; width; height; first; last; names }  (px, first/last are
      0-based entry indexes). Targets not on the page -> missing.
    #>
    param([string[]]$Names, [string[]]$Targets, [double]$LastRowCenterY, [double]$RowPitch, [int]$X, [int]$Width, [int]$Height)
    $all = @($Names)
    $idx = New-Object System.Collections.ArrayList
    $missing = New-Object System.Collections.ArrayList
    foreach ($t in @($Targets)) {
        $i = [array]::IndexOf($all, [string]$t)
        if ($i -lt 0) { [void]$missing.Add([string]$t) } else { [void]$idx.Add($i) }
    }
    $sorted = @($idx | Sort-Object -Unique)
    $rects = New-Object System.Collections.ArrayList
    $n = $all.Count
    $k = 0
    while ($k -lt $sorted.Count) {
        $a = [int]$sorted[$k]; $b = $a
        while (($k + 1) -lt $sorted.Count -and [int]$sorted[$k + 1] -eq ($b + 1)) { $k++; $b = [int]$sorted[$k] }
        $centerA = $LastRowCenterY - ($n - 1 - $a) * $RowPitch
        $centerB = $LastRowCenterY - ($n - 1 - $b) * $RowPitch
        $top = [int][Math]::Round($centerA - $Height / 2.0, [System.MidpointRounding]::AwayFromZero)
        $bottom = [int][Math]::Round($centerB + $Height / 2.0, [System.MidpointRounding]::AwayFromZero)
        [void]$rects.Add(@{ x = $X; y = $top; width = $Width; height = ($bottom - $top); first = $a; last = $b; names = @($all[$a..$b]) })
        $k++
    }
    return @{ rects = $rects.ToArray(); missing = $missing.ToArray() }
}

function Get-EbiInkBands {
    <#
      PURE. Counts[i] = how many ink pixels pixel row i of a strip holds.
      Rows with at least MinInk ink pixels form bands; bands closer than
      MergeGap rows merge (anti-aliasing leaves a blank row inside a
      digit). Returns @(@{ top; bottom; height }) top to bottom (0-based,
      bottom inclusive, relative to the strip).
    #>
    param([int[]]$Counts, [int]$MinInk = 1, [int]$MergeGap = 1)
    $bands = New-Object System.Collections.ArrayList
    $start = -1; $lastInk = -1
    for ($i = 0; $i -lt @($Counts).Count; $i++) {
        if ([int]$Counts[$i] -ge $MinInk) {
            if ($start -lt 0) { $start = $i }
            elseif (($i - $lastInk - 1) -gt $MergeGap) {
                [void]$bands.Add(@{ top = $start; bottom = $lastInk; height = ($lastInk - $start + 1) })
                $start = $i
            }
            $lastInk = $i
        }
    }
    if ($start -ge 0) { [void]$bands.Add(@{ top = $start; bottom = $lastInk; height = ($lastInk - $start + 1) }) }
    return $bands.ToArray()
}

function Resolve-EbiAnchoredRect {
    <#
      PURE. A rect given relative to one corner of an image -> absolute
      top-left px. Rect: @{ x; y; w; h; anchor } where anchor is
      tl (default) | tr | bl | br; for tr/br x is the distance from the
      RIGHT edge to the rect's left side, for bl/br y the distance from the
      BOTTOM edge to the rect's top (so a status-bar cell stays put when the
      window is resized). Returns @{ x; y; w; h }.
    #>
    param($Rect, [int]$ImageWidth, [int]$ImageHeight)
    $a = if ($Rect -is [System.Collections.IDictionary] -and $Rect.Contains('anchor') -and $null -ne $Rect['anchor']) { ([string]$Rect['anchor']).ToLowerInvariant() } else { 'tl' }
    $x = [double]$Rect['x']; $y = [double]$Rect['y']
    if ($a -eq 'tr' -or $a -eq 'br') { $x = $ImageWidth - $x }
    if ($a -eq 'bl' -or $a -eq 'br') { $y = $ImageHeight - $y }
    return @{ x = $x; y = $y; w = [double]$Rect['w']; h = [double]$Rect['h'] }
}

function ConvertTo-EbiSheetRect {
    <#
      PURE. A rect in a picture's IMAGE pixels -> sheet points, given where
      the picture shape sits (points) and the image's pixel size. Scaling
      is taken from the shape (shape width / image width), so a picture
      Excel shrank still gets its box on the right pixels.
      Returns @{ left; top; width; height } in points.
    #>
    param([double]$PicLeft, [double]$PicTop, [double]$PicWidth, [double]$PicHeight, [int]$ImageWidth, [int]$ImageHeight, $Rect)
    $sx = if ($ImageWidth -gt 0) { $PicWidth / $ImageWidth } else { 0.75 }
    $sy = if ($ImageHeight -gt 0) { $PicHeight / $ImageHeight } else { 0.75 }
    return @{ left = $PicLeft + [double]$Rect['x'] * $sx; top = $PicTop + [double]$Rect['y'] * $sy; width = [double]$Rect['w'] * $sx; height = [double]$Rect['h'] * $sy }
}

function New-EbiStackPlan {
    <#
      PURE. Capture sets -> the picture list excel.insert_pictures stacks
      down a sheet. Each set is @{ first; last? } (paths): first alone when
      there is no last; first, then the Separator picture (optional; placed
      at SeparatorColumn, SeparatorRows tall), then last, when there is.
      MarkRects go on the LAST picture of each set (the one that shows the
      end of the data). Sets are separated by GapRows blank rows.
      Returns an array of @{ path; column; gapRows; rects; role }.
    #>
    param([object[]]$Sets, [string]$Separator = '', [string]$SeparatorColumn = '', [int]$SeparatorRows = 4, [object[]]$MarkRects = @(), [int]$GapRows = 1, [string]$Column = 'B')
    $plan = New-Object System.Collections.ArrayList
    $first = $true
    foreach ($s in @($Sets)) {
        if ($null -eq $s) { continue }
        $p1 = [string]$s['first']
        $p2 = if ($s.Contains('last') -and -not [string]::IsNullOrWhiteSpace([string]$s['last'])) { [string]$s['last'] } else { '' }
        $gap = if ($first) { 0 } else { $GapRows }
        $first = $false
        if ($p2 -eq '') {
            [void]$plan.Add(@{ path = $p1; column = $Column; gapRows = $gap; rects = @($MarkRects); role = 'only' })
            continue
        }
        [void]$plan.Add(@{ path = $p1; column = $Column; gapRows = $gap; rects = @(); role = 'first' })
        if (-not [string]::IsNullOrWhiteSpace($Separator)) {
            $col = if ([string]::IsNullOrWhiteSpace($SeparatorColumn)) { $Column } else { $SeparatorColumn }
            [void]$plan.Add(@{ path = $Separator; column = $col; gapRows = 0; rects = @(); role = 'separator'; rows = $SeparatorRows })
        }
        [void]$plan.Add(@{ path = $p2; column = $Column; gapRows = 0; rects = @($MarkRects); role = 'last' })
    }
    return $plan.ToArray()
}
