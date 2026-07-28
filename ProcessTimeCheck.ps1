# ============================================================
#  ProcessTimeCheck.ps1
#
#  PURE library for the ProcessTime phase's on-sheet audit ("check")
#  columns -- NO Excel COM, NO OCR, NO mapping I/O. Dot-source only
#  (no param() block), so ProcessTime.ps1 can dot-source it safely per
#  CLAUDE.md's dot-source rule. ASCII source; Japanese headers built from
#  [char] code points so the file is codepage-agnostic.
#
#  The ProcessTime output workbook lays each result out as A..H data
#  columns (No. / GIFT-GFIX / correl / start / end / duration / count /
#  job). This module owns the worksheet-side verification columns appended
#  after the data (I onward): it returns a DATA-DRIVEN column spec and a
#  pure formula-templating helper. ProcessTime.ps1's COM-side
#  Set-ProcessTimeCheckColumns walks the same spec to write the headers,
#  per-row formulas and number formats, so the formulas live in exactly
#  one place, are unit-testable, and are decoupled from the data-row write
#  loop (previously they were inlined per row, un-testable and coupled).
#
#  Convention: functions return plain arrays -- never return ,@(...)
#  because callers wrap calls in @() and that nests in PS 5.1 (see
#  ProcessTimeParse.ps1's identical note).
# ============================================================

# ---------------------------------------------------------------------------
# Get-ProcessTimeCheckColumnSpec
#   The ordered spec for the audit columns appended after the A..H data
#   columns. Each entry is a hashtable:
#     Col          worksheet column letter (I / J / K)
#     ColIndex     1-based worksheet column index (9 / 10 / 11)
#     Header       Japanese column header (built from [char], ASCII source)
#     NumberFormat cell number format ('' = leave the workbook default)
#     Width        column width the shared formatting block should apply
#     Formula      formula TEMPLATE; '{0}' is the data row number (fed to
#                  New-ProcessTimeCheckFormula). Each template is
#                  SELF-GUARDING (blank source cells -> the formula yields
#                  "" itself) so a partial row -- e.g. an OCR read that got
#                  the start but not the end time -- leaves the check cell
#                  blank instead of showing a spurious value or a #VALUE
#                  error. This preserves the old inline behavior (I/J were
#                  only written when start+end+duration were all real) with
#                  no per-row COM inspection needed.
#
#   Columns:
#     I  shori-jikan (kensan) -- worksheet re-derivation of the duration
#        (=E-D), a real time serial ('[h]:mm:ss'). Guarded on both D and E
#        being real numbers (ISNUMBER), so a text fallback in either never
#        produces a #VALUE error.
#     J  chekku -- T/F compare of the written duration (F) against the
#        re-derived one (I), to the second. Guarded on I and F being real.
#     K  kensu (sanshou) -- the EXPECTED record count for this row's job,
#        pulled out of the project's monthly reference workbook with the
#        operator's own lookup:
#          INDEX(<value column>, MATCH(LEFT(C{0},<KeyLength>)&"*", <key column>, 0))
#        A miss yields "" (IFERROR), never #N/A. When no reference is
#        configured the column is still emitted, but as an inert PLACEHOLDER
#        (the same formula with <DIR>/<BOOK>/<SHEET> tokens, written as TEXT
#        so Excel never tries to resolve a bogus external link): the operator
#        can fill the real path in later with a search-and-replace instead of
#        rebuilding the formula from scratch. -CountReference @{...} with
#        PlaceholderWhenUnset = $false leaves the cells blank instead.
#     L  kensu-chekku -- the count check itself, "OK" / "NG" / blank:
#          * against K when K has a value (the reference is authoritative);
#          * otherwise against the SAME correl's other side -- the GIFT row is
#            compared with its GFIX row and vice versa. This is the check that
#            always works, with or without a reference workbook.
#          * EQUAL counts read OK -- including 0 vs 0. A zero count is a
#            legitimate result (an interface with no data that day); only a
#            DISAGREEMENT between the two sides is a finding.
#          * blank whenever the value being compared against is missing (a
#            partial row, an unlisted job, or a correl with only one side),
#            so "not checkable" never reads as a failure.
#        The paired row is not at a fixed offset -- the layout groups all GIFT
#        rows then all GFIX rows per job -- so the COM writer resolves it per
#        row (Get-ProcessTimeCountPairMap) and fills the template's {1}.
#
# ---------------------------------------------------------------------------
function Get-ProcessTimeCheckColumnSpec {
    param([hashtable]$CountReference = $null)
    # shori-jikan (kensan) -- "processing time (recheck)"
    $hCalc  = [string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593 + '(' + [char]0x691C + [char]0x7B97 + ')'
    # chekku -- "check"
    $hCheck = [string][char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF
    # kensu-chekku -- "record-count check"
    $hCount = [string][char]0x4EF6 + [char]0x6570 + [char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF
    # kensu (sanshou) -- "record count (reference)"
    $hRef   = [string][char]0x4EF6 + [char]0x6570 + '(' + [char]0x53C2 + [char]0x7167 + ')'

    $spec = @(
        @{
            Col = 'I'; ColIndex = 9; Header = $hCalc; NumberFormat = '[h]:mm:ss'; Width = 20.0
            Formula = '=IF(AND(ISNUMBER(D{0}),ISNUMBER(E{0})),E{0}-D{0},"")'
        },
        @{
            Col = 'J'; ColIndex = 10; Header = $hCheck; NumberFormat = ''; Width = 8.0
            Formula = '=IF(AND(ISNUMBER(I{0}),ISNUMBER(F{0})),IF(ROUND(F{0}*86400,0)=ROUND(I{0}*86400,0),"T","F"),"")'
        }
    )

    $refOk = Test-ProcessTimeCountReference $CountReference
    $emitPlaceholder = $true
    if ($null -ne $CountReference -and $CountReference -is [hashtable] -and
        $CountReference.ContainsKey('PlaceholderWhenUnset')) {
        $emitPlaceholder = [bool]$CountReference['PlaceholderWhenUnset']
    }

    # K -- the reference lookup. Live formula when configured; otherwise an
    # inert text placeholder (or nothing, when the operator turned that off).
    $kEntry = @{ Col = 'K'; ColIndex = 11; Header = $hRef; NumberFormat = ''; Width = 14.0; Formula = ''; IsText = $false }
    if ($refOk) {
        $kEntry.Formula = (New-ProcessTimeCountLookupFormula -Reference $CountReference)
    } elseif ($emitPlaceholder) {
        $kEntry.Formula = (New-ProcessTimeCountPlaceholderFormula)
        $kEntry.IsText  = $true          # written as text, never evaluated
        $kEntry.NumberFormat = '@'
    }
    $spec += $kEntry

    # L -- the count check. Uses K when it holds a value, else the same
    # correl's other side ({1} = that row, filled in by the COM writer).
    $spec += @{
        Col = 'L'; ColIndex = 12; Header = $hCount; NumberFormat = ''; Width = 12.0
        NeedsPair = $true
        Formula = '=IF(TRIM(G{0})="","",IF(K{0}<>"",IFERROR(IF(VALUE(SUBSTITUTE(G{0},",",""))=VALUE(K{0}),"OK","NG"),"NG"),IF(TRIM(G{1})="","",IFERROR(IF(VALUE(SUBSTITUTE(G{0},",",""))=VALUE(SUBSTITUTE(G{1},",","")),"OK","NG"),"NG"))))'
        # No paired row for this correl (only one side present): the only
        # comparison left is the reference one.
        FormulaNoPair = '=IF(OR(TRIM(G{0})="",K{0}=""),"",IFERROR(IF(VALUE(SUBSTITUTE(G{0},",",""))=VALUE(K{0}),"OK","NG"),"NG"))'
    }
    return $spec
}

# ---------------------------------------------------------------------------
# New-ProcessTimeCheckFormula
#   Fills a check-column formula TEMPLATE for one data row: replaces every
#   '{0}' with the row number and returns the concrete formula string. Pure
#   (string only) so it is directly unit-testable, e.g.
#   New-ProcessTimeCheckFormula -Template '=E{0}-D{0}' -Row 5 -> '=E5-D5'.
# ---------------------------------------------------------------------------
function New-ProcessTimeCheckFormula {
    param(
        [Parameter(Mandatory = $true)][string]$Template,
        [Parameter(Mandatory = $true)][int]$Row,
        # '{1}' -- the row holding the SAME correl's other side (GIFT <-> GFIX).
        # Only the count-check template uses it; 0 means "no paired row", and
        # the caller should have picked FormulaNoPair instead.
        [int]$PairRow = 0
    )
    return ($Template -f $Row, $PairRow)
}

# ---------------------------------------------------------------------------
# Get-ProcessTimeCountPairMap
#   Maps each data row to the row holding the SAME correl's other side, so the
#   count check can compare GIFT against GFIX. The output layout groups all
#   GIFT rows then all GFIX rows per job, so the pair is NOT at a fixed
#   offset -- and rows retained from an earlier run sit wherever that run left
#   them. The COM caller therefore reads (row, side, correl) off the sheet and
#   hands it here; this stays pure and unit-testable.
#
#   -Rows: objects/hashtables with Row (int), Side ('GIFT'/'GFIX') and
#   CorrelId. Returns a hashtable rowNumber -> paired rowNumber, containing
#   only rows that actually HAVE a partner. A correl with two rows on the same
#   side (a duplicate) is left unpaired rather than guessing.
# ---------------------------------------------------------------------------
function Get-ProcessTimeCountPairMap {
    param($Rows)
    $map = @{}
    if ($null -eq $Rows) { return $map }
    $bySideCorrel = @{}
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $side = ([string]$r.Side).Trim().ToUpperInvariant()
        $correl = ([string]$r.CorrelId).Trim()
        if ([string]::IsNullOrWhiteSpace($correl)) { continue }
        if ($side -ne 'GIFT' -and $side -ne 'GFIX') { continue }
        $key = ('{0}|{1}' -f $side, $correl)
        if ($bySideCorrel.ContainsKey($key)) { $bySideCorrel[$key] = 0 }   # duplicate -> unusable
        else { $bySideCorrel[$key] = [int]$r.Row }
    }
    foreach ($r in @($Rows)) {
        if ($null -eq $r) { continue }
        $side = ([string]$r.Side).Trim().ToUpperInvariant()
        $correl = ([string]$r.CorrelId).Trim()
        if ([string]::IsNullOrWhiteSpace($correl)) { continue }
        $other = if ($side -eq 'GIFT') { 'GFIX' } elseif ($side -eq 'GFIX') { 'GIFT' } else { continue }
        # This row's own side must itself be unambiguous: a correl listed
        # twice on one side gives no single "the other row" to compare with.
        $selfKey = ('{0}|{1}' -f $side, $correl)
        if (-not $bySideCorrel.ContainsKey($selfKey) -or [int]$bySideCorrel[$selfKey] -le 0) { continue }
        $key = ('{0}|{1}' -f $other, $correl)
        if (-not $bySideCorrel.ContainsKey($key)) { continue }
        $pair = [int]$bySideCorrel[$key]
        if ($pair -le 0) { continue }
        $self = [int]$r.Row
        if ($self -le 0 -or $self -eq $pair) { continue }
        $map[$self] = $pair
    }
    return $map
}

# ---------------------------------------------------------------------------
# New-ProcessTimeCountPlaceholderFormula
#   The K-column lookup with the reference workbook left as obvious tokens,
#   for the case where no reference is configured yet. It is written into the
#   sheet as TEXT (never evaluated), so a bogus external link can neither
#   prompt for updates nor show #REF -- the operator replaces <DIR>, <BOOK>
#   and <SHEET>, then converts the column back to formulas (select column ->
#   Data -> Text to Columns -> Finish) once the real path is known.
#   ASCII tokens on purpose: they survive any codepage and are unambiguous
#   targets for a search-and-replace.
# ---------------------------------------------------------------------------
function New-ProcessTimeCountPlaceholderFormula {
    param([int]$KeyLength = 7)
    $ref = @{
        Enabled = $true; Directory = '<DIR>'; FileName = '<BOOK>.xlsx'; SheetName = '<SHEET>'
        KeyColumn = 'G'; ValueColumn = 'O'; KeyLength = $KeyLength; FirstRow = 1; LastRow = 20000
    }
    return (New-ProcessTimeCountLookupFormula -Reference $ref)
}

# ---------------------------------------------------------------------------
# Test-ProcessTimeCountReference
#   $true when the hashtable is a RESOLVED, usable count reference (produced
#   by Resolve-ProcessTimeCountReference): enabled and carrying both a file
#   name and a sheet name. Anything else -- $null, disabled, half-filled --
#   is $false, so the caller silently keeps the self-contained K check.
# ---------------------------------------------------------------------------
function Test-ProcessTimeCountReference {
    param($Reference)
    if ($null -eq $Reference) { return $false }
    if ($Reference -isnot [hashtable]) { return $false }
    if (-not $Reference.ContainsKey('Enabled') -or -not [bool]$Reference['Enabled']) { return $false }
    if ([string]::IsNullOrWhiteSpace([string]$Reference['FileName'])) { return $false }
    if ([string]::IsNullOrWhiteSpace([string]$Reference['SheetName'])) { return $false }
    return $true
}

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeCountReference
#   Turns the raw ProcessTime.CountReference config block into the resolved
#   reference the spec builder consumes, expanding the two filename/sheet
#   tokens against THIS output workbook's tag:
#     {Tag}    the output tag ('JOD', 'JRV', ...) -- the same classification
#              that picks the output workbook, so each tag looks its counts
#              up in its own reference file.
#     {Month}  -Month verbatim (the caller passes the config's Month, or the
#              current month number when that is blank). Any Japanese suffix
#              (the 'gatsu' month character) lives in the config string
#              itself, keeping this file ASCII.
#   Returns @{ Enabled = $false } (never $null) when the feature is off, the
#   config is incomplete, or a token cannot be expanded -- e.g. a '{Tag}'
#   template under OutputMode 'Single', which has no tag. Failing to a
#   disabled reference is deliberate: a half-expanded path would silently
#   point Excel at a workbook that does not exist.
# ---------------------------------------------------------------------------
function Resolve-ProcessTimeCountReference {
    param(
        $Reference,
        [string]$Tag = '',
        [string]$Month = ''
    )
    # A disabled result still carries PlaceholderWhenUnset: the K column is
    # emitted either way, and that flag decides whether it gets the inert
    # placeholder lookup or stays empty.
    $off = @{ Enabled = $false; PlaceholderWhenUnset = $true }
    if ($null -ne $Reference -and $Reference -is [hashtable] -and $Reference.ContainsKey('PlaceholderWhenUnset')) {
        $off['PlaceholderWhenUnset'] = [bool]$Reference['PlaceholderWhenUnset']
    }
    if ($null -eq $Reference -or $Reference -isnot [hashtable]) { return $off }
    if (-not $Reference.ContainsKey('Enabled') -or -not [bool]$Reference['Enabled']) { return $off }

    $fileName  = [string]$Reference['FileName']
    $sheetName = [string]$Reference['SheetName']
    if ([string]::IsNullOrWhiteSpace($fileName) -or [string]::IsNullOrWhiteSpace($sheetName)) { return $off }

    $tag = ([string]$Tag).Trim()
    foreach ($t in @($fileName, $sheetName)) {
        if ($t -match '\{Tag\}' -and [string]::IsNullOrWhiteSpace($tag)) { return $off }
    }
    $expand = {
        param([string]$Text)
        $out = $Text.Replace('{Tag}', $tag)
        return $out.Replace('{Month}', ([string]$Month))
    }
    $fileName  = & $expand $fileName
    $sheetName = & $expand $sheetName

    $keyCol = [string]$Reference['KeyColumn']
    $valCol = [string]$Reference['ValueColumn']
    if ([string]::IsNullOrWhiteSpace($keyCol)) { $keyCol = 'G' }
    if ([string]::IsNullOrWhiteSpace($valCol)) { $valCol = 'O' }

    $keyLen = 0
    if ($Reference.ContainsKey('KeyLength') -and $null -ne $Reference['KeyLength']) {
        try { $keyLen = [int]$Reference['KeyLength'] } catch { $keyLen = 0 }
    }
    $firstRow = 0; $lastRow = 0
    if ($Reference.ContainsKey('FirstRow') -and $null -ne $Reference['FirstRow']) {
        try { $firstRow = [int]$Reference['FirstRow'] } catch { $firstRow = 0 }
    }
    if ($Reference.ContainsKey('LastRow') -and $null -ne $Reference['LastRow']) {
        try { $lastRow = [int]$Reference['LastRow'] } catch { $lastRow = 0 }
    }

    return @{
        Enabled     = $true
        PlaceholderWhenUnset = [bool]$off['PlaceholderWhenUnset']
        Directory   = [string]$Reference['Directory']
        FileName    = $fileName
        SheetName   = $sheetName
        KeyColumn   = $keyCol.Trim().ToUpperInvariant()
        ValueColumn = $valCol.Trim().ToUpperInvariant()
        KeyLength   = $keyLen
        FirstRow    = $firstRow
        LastRow     = $lastRow
    }
}

# ---------------------------------------------------------------------------
# New-ProcessTimeExternalRange
#   Builds one Excel external-workbook range reference:
#     'C:\dir\[Book.xlsx]Sheet'!$G$1:$G$20000
#   Notes that matter on the office PC:
#     * The DIRECTORY is included whenever configured. The short form Excel
#       shows while the source workbook is open ('[Book.xlsx]Sheet'!...) only
#       resolves for an OPEN workbook; the full-path form also works against a
#       closed one, which is the normal case here. A blank Directory still
#       emits the short form (for a reference workbook the operator keeps
#       open) rather than guessing a path.
#     * The range is BOUNDED by default (FirstRow/LastRow). Whole-column
#       external references to a CLOSED workbook are unreliable in Excel;
#       pass 0/0 only when the reference workbook is always open.
#     * A path or sheet name containing an apostrophe has it doubled, per
#       Excel's quoting rule.
# ---------------------------------------------------------------------------
function New-ProcessTimeExternalRange {
    param(
        [string]$Directory,
        [string]$FileName,
        [string]$SheetName,
        [string]$Column,
        [int]$FirstRow = 0,
        [int]$LastRow = 0
    )
    $dir = ([string]$Directory).Trim()
    if ($dir.EndsWith('\')) { $dir = $dir.TrimEnd('\') }
    $book = if ([string]::IsNullOrWhiteSpace($dir)) { ('[{0}]' -f $FileName) }
            else { ('{0}\[{1}]' -f $dir, $FileName) }
    $qualifier = ($book + $SheetName).Replace("'", "''")
    $col = ([string]$Column).Trim().ToUpperInvariant()
    $range = if ($FirstRow -gt 0 -and $LastRow -ge $FirstRow) {
        ('${0}${1}:${0}${2}' -f $col, $FirstRow, $LastRow)
    } else {
        ('${0}:${0}' -f $col)
    }
    return ("'{0}'!{1}" -f $qualifier, $range)
}

# ---------------------------------------------------------------------------
# New-ProcessTimeCountLookupFormula
#   The K-column formula TEMPLATE ('{0}' = data row number) for the expected
#   record count, i.e. the operator's own hand-written lookup:
#     =IFERROR(INDEX(<value range>,MATCH(LEFT(C{0},7)&"*",<key range>,0)),"")
#   LEFT(...)&"*" is a PREFIX match: the reference sheet keys on the job code
#   while column C holds the longer correl id, so only the first KeyLength
#   characters are compared (MATCH's wildcard form, hence match type 0).
#   KeyLength <= 0 compares the whole cell instead. IFERROR turns an unlisted
#   job's #N/A into "", which the L compare then reads as "nothing to check".
# ---------------------------------------------------------------------------
function New-ProcessTimeCountLookupFormula {
    param([hashtable]$Reference)
    if (-not (Test-ProcessTimeCountReference $Reference)) { return '' }
    # NOT named $args -- that is a PowerShell automatic variable.
    $rangeArgs = @{
        Directory = [string]$Reference['Directory']
        FileName  = [string]$Reference['FileName']
        SheetName = [string]$Reference['SheetName']
        FirstRow  = [int]$Reference['FirstRow']
        LastRow   = [int]$Reference['LastRow']
    }
    $valRange = New-ProcessTimeExternalRange @rangeArgs -Column ([string]$Reference['ValueColumn'])
    $keyRange = New-ProcessTimeExternalRange @rangeArgs -Column ([string]$Reference['KeyColumn'])
    $keyLen = [int]$Reference['KeyLength']
    $lookup = if ($keyLen -gt 0) { ('LEFT(C{{0}},{0})&"*"' -f $keyLen) } else { 'C{0}' }
    return ('=IFERROR(INDEX({0},MATCH({1},{2},0)),"")' -f $valRange, $lookup, $keyRange)
}
