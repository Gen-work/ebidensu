#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ProcessTimeCheck.ps1')

Reset-Tests 'ProcessTimeCheck'

# ---------------------------------------------------------------------------
# New-ProcessTimeCheckFormula : pure {0} -> row-number substitution.
# ---------------------------------------------------------------------------
Assert-Equal '=E5-D5' (New-ProcessTimeCheckFormula -Template '=E{0}-D{0}' -Row 5) 'row 5 fills the duration re-derivation template'
Assert-Equal '=E27-D27' (New-ProcessTimeCheckFormula -Template '=E{0}-D{0}' -Row 27) 'a two-digit row fills both placeholders'
Assert-Equal '=A2+1' (New-ProcessTimeCheckFormula -Template '=A{0}+1' -Row 2) 'a template with a single placeholder still fills'

# ---------------------------------------------------------------------------
# Get-ProcessTimeCheckColumnSpec : shape / columns / headers.
# ---------------------------------------------------------------------------
$spec = @(Get-ProcessTimeCheckColumnSpec)
Assert-Equal 4 $spec.Count 'four audit columns (I/J/K/L)'

Assert-Equal 'I' $spec[0].Col 'first audit column is I'
Assert-Equal 'J' $spec[1].Col 'second audit column is J'
Assert-Equal 'K' $spec[2].Col 'third audit column is K'
Assert-Equal 'L' $spec[3].Col 'fourth audit column is L'
Assert-Equal 9  $spec[0].ColIndex 'I is column index 9 (after the A..H data columns)'
Assert-Equal 10 $spec[1].ColIndex 'J is column index 10'
Assert-Equal 11 $spec[2].ColIndex 'K is column index 11'
Assert-Equal 12 $spec[3].ColIndex 'L is column index 12'

# Headers are built from [char] code points (ASCII source, codepage-agnostic).
$hCalc  = [string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593 + '(' + [char]0x691C + [char]0x7B97 + ')'  # shori-jikan (kensan)
$hCheck = [string][char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF                                          # chekku
$hCount = [string][char]0x4EF6 + [char]0x6570 + [char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF            # kensu-chekku
$hRefHdr = [string][char]0x4EF6 + [char]0x6570 + '(' + [char]0x53C2 + [char]0x7167 + ')'                             # kensu (sanshou)
Assert-Equal $hCalc  $spec[0].Header 'I header is shori-jikan (kensan)'
Assert-Equal $hCheck $spec[1].Header 'J header is chekku'
Assert-Equal $hRefHdr $spec[2].Header 'K header is kensu (sanshou)'
Assert-Equal $hCount  $spec[3].Header 'L header is kensu-chekku'

Assert-Equal '[h]:mm:ss' $spec[0].NumberFormat 'I column carries a duration number format'
Assert-Equal '' $spec[1].NumberFormat 'J column leaves the default format'
Assert-Equal '@' $spec[2].NumberFormat 'K holds a text placeholder when no reference is configured'
Assert-Equal '' $spec[3].NumberFormat 'L column leaves the default format'

# ---------------------------------------------------------------------------
# Filling each column's real template through the pure helper.
# ---------------------------------------------------------------------------
$iFilled = New-ProcessTimeCheckFormula -Template $spec[0].Formula -Row 5
Assert-True ($iFilled -notmatch '\{0\}') 'I formula has no leftover placeholder after filling'
Assert-True ($iFilled -match 'E5') 'I formula references E5 for row 5'
Assert-True ($iFilled -match 'D5') 'I formula references D5 for row 5'

$jFilled = New-ProcessTimeCheckFormula -Template $spec[1].Formula -Row 5
Assert-True ($jFilled -notmatch '\{0\}') 'J formula has no leftover placeholder after filling'
Assert-True ($jFilled -match 'F5') 'J formula compares the written duration F5'
Assert-True ($jFilled -match 'I5') 'J formula compares against the re-derived duration I5'
Assert-True ($jFilled -match '"T","F"') 'J formula yields T/F'

# K with no configured reference: the inert placeholder lookup.
$kFilled = New-ProcessTimeCheckFormula -Template $spec[2].Formula -Row 5
Assert-True ($kFilled -notmatch '\{0\}') 'K placeholder has no leftover placeholder after filling'
Assert-True ([bool]$spec[2].IsText) 'K placeholder is written as text, never evaluated as a formula'
Assert-True ($kFilled -match '<DIR>') 'K placeholder carries a directory token to replace'
Assert-True ($kFilled -match '<BOOK>') 'K placeholder carries a workbook token to replace'
Assert-True ($kFilled -match '<SHEET>') 'K placeholder carries a sheet token to replace'
Assert-True ($kFilled -match 'LEFT\(C5,7\)') 'K placeholder already has the row-correct lookup key'

# The column COUNT never varies -- I/J/K/L indexes are fixed, so an unwanted
# placeholder empties K's cells rather than removing the column.
$specNoPh = @(Get-ProcessTimeCheckColumnSpec -CountReference @{ PlaceholderWhenUnset = $false })
Assert-Equal 4 $specNoPh.Count 'PlaceholderWhenUnset = $false keeps the four-column layout'
Assert-Equal '' $specNoPh[2].Formula 'PlaceholderWhenUnset = $false leaves K''s cells empty'

# L (count check) with no reference: GIFT vs GFIX, OK/NG, 0 vs 0 is OK.
Assert-True ([bool]$spec[3].NeedsPair) 'L needs the paired row resolved by the COM writer'
$lPair = New-ProcessTimeCheckFormula -Template $spec[3].Formula -Row 5 -PairRow 12
Assert-True ($lPair -notmatch '\{0\}') 'L formula has no leftover row placeholder'
Assert-True ($lPair -notmatch '\{1\}') 'L formula has no leftover pair placeholder'
Assert-True ($lPair -match 'G5') 'L compares this row''s count'
Assert-True ($lPair -match 'G12') 'L compares against the paired row''s count'
Assert-True ($lPair -match 'K5') 'L prefers the reference count when K has a value'
Assert-True ($lPair -match '"OK","NG"') 'L reports OK/NG'
Assert-True ($lPair -notmatch '>0') 'L no longer requires a POSITIVE count -- 0 vs 0 is a match, not NG'
Assert-True ($lPair -match 'TRIM\(G5\)=""') 'L stays blank when this row has no count'
Assert-True ($lPair -match 'TRIM\(G12\)=""') 'L stays blank when the other side has no count'

$lNoPair = New-ProcessTimeCheckFormula -Template $spec[3].FormulaNoPair -Row 5
Assert-True ($lNoPair -notmatch '\{1\}') 'the no-pair template never references a paired row'
Assert-True ($lNoPair -match 'K5=""') 'with no paired row the check falls back to the reference only'

# ---------------------------------------------------------------------------
# Get-ProcessTimeCountPairMap : GIFT <-> GFIX row pairing.
# ---------------------------------------------------------------------------
$pairRows = @(
    [pscustomobject]@{ Row = 2; Side = 'GIFT'; CorrelId = 'A' },
    [pscustomobject]@{ Row = 3; Side = 'GIFT'; CorrelId = 'B' },
    [pscustomobject]@{ Row = 8; Side = 'GFIX'; CorrelId = 'A' },
    [pscustomobject]@{ Row = 9; Side = 'GFIX'; CorrelId = 'B' }
)
$pm = Get-ProcessTimeCountPairMap -Rows $pairRows
Assert-Equal 8 $pm[2] 'a GIFT row pairs with its correl''s GFIX row (not a fixed offset)'
Assert-Equal 2 $pm[8] 'the pairing is symmetric'
Assert-Equal 9 $pm[3] 'the second correl pairs independently'
Assert-Equal 4 $pm.Count 'every row with a partner is mapped'

$pmOne = Get-ProcessTimeCountPairMap -Rows @([pscustomobject]@{ Row = 2; Side = 'GIFT'; CorrelId = 'A' })
Assert-Equal 0 $pmOne.Count 'a correl with only one side has no pair'
$pmDup = Get-ProcessTimeCountPairMap -Rows @(
    [pscustomobject]@{ Row = 2; Side = 'GIFT'; CorrelId = 'A' },
    [pscustomobject]@{ Row = 3; Side = 'GIFT'; CorrelId = 'A' },
    [pscustomobject]@{ Row = 8; Side = 'GFIX'; CorrelId = 'A' }
)
Assert-Equal 0 $pmDup.Count 'a duplicated side is left unpaired rather than guessed'
Assert-Equal 0 (Get-ProcessTimeCountPairMap -Rows $null).Count 'null input -> empty map'
Assert-Equal 0 (Get-ProcessTimeCountPairMap -Rows @(
    [pscustomobject]@{ Row = 2; Side = 'DF'; CorrelId = 'A' },
    [pscustomobject]@{ Row = 3; Side = ''; CorrelId = '' })).Count 'unknown sides / blank correls are ignored'

# ---------------------------------------------------------------------------
# Resolve-ProcessTimeCountReference : token expansion + fail-safe disabling.
# ---------------------------------------------------------------------------
$gatsu = [string][char]0x6708   # 'gatsu' (month) -- keeps this file ASCII
$refCfg = @{
    Enabled     = $true
    Directory   = 'C:\ref'
    FileName    = ('GPCS({Tag})_{Month}' + $gatsu + '.xlsx')
    SheetName   = '{Tag}'
    KeyColumn   = 'g'
    ValueColumn = 'o'
    KeyLength   = 7
    FirstRow    = 1
    LastRow     = 20000
}

$r = Resolve-ProcessTimeCountReference -Reference $refCfg -Tag 'JOD' -Month '7'
Assert-True ([bool]$r.Enabled) 'a complete, enabled reference resolves'
Assert-Equal ('GPCS(JOD)_7' + $gatsu + '.xlsx') $r.FileName '{Tag}/{Month} both expand in the file name'
Assert-Equal 'JOD' $r.SheetName '{Tag} expands in the sheet name'
Assert-Equal 'G' $r.KeyColumn 'key column is upper-cased'
Assert-Equal 'O' $r.ValueColumn 'value column is upper-cased'
Assert-Equal 7 $r.KeyLength 'key length carries through'

Assert-True (-not (Resolve-ProcessTimeCountReference -Reference $refCfg -Tag '' -Month '7').Enabled) `
    'a {Tag} template with no tag (OutputMode Single) disables rather than half-expanding'
Assert-True (-not (Resolve-ProcessTimeCountReference -Reference $null -Tag 'JOD').Enabled) `
    'a null config resolves to disabled'
Assert-True (-not (Resolve-ProcessTimeCountReference -Reference @{ Enabled = $false; FileName = 'x.xlsx'; SheetName = 'S' } -Tag 'JOD').Enabled) `
    'Enabled=$false resolves to disabled'
Assert-True (-not (Resolve-ProcessTimeCountReference -Reference @{ Enabled = $true; FileName = ''; SheetName = 'S' } -Tag 'JOD').Enabled) `
    'a blank file name resolves to disabled'

$rNoTag = Resolve-ProcessTimeCountReference -Reference @{ Enabled = $true; FileName = 'fixed.xlsx'; SheetName = 'Sheet1' } -Tag ''
Assert-True ([bool]$rNoTag.Enabled) 'a token-free template needs no tag'
Assert-Equal 'G' $rNoTag.KeyColumn 'key column defaults to G'
Assert-Equal 'O' $rNoTag.ValueColumn 'value column defaults to O'

# ---------------------------------------------------------------------------
# New-ProcessTimeExternalRange : Excel external-reference syntax.
# ---------------------------------------------------------------------------
Assert-Equal "'C:\ref\[B.xlsx]JOD'!`$O`$1:`$O`$20000" `
    (New-ProcessTimeExternalRange -Directory 'C:\ref' -FileName 'B.xlsx' -SheetName 'JOD' -Column 'O' -FirstRow 1 -LastRow 20000) `
    'full-path bounded range (works against a CLOSED reference workbook)'
Assert-Equal "'C:\ref\[B.xlsx]JOD'!`$G:`$G" `
    (New-ProcessTimeExternalRange -Directory 'C:\ref\' -FileName 'B.xlsx' -SheetName 'JOD' -Column 'g') `
    'a trailing backslash is trimmed and 0 rows means the whole column'
Assert-Equal "'[B.xlsx]JOD'!`$O`$1:`$O`$9" `
    (New-ProcessTimeExternalRange -Directory '' -FileName 'B.xlsx' -SheetName 'JOD' -Column 'O' -FirstRow 1 -LastRow 9) `
    'a blank directory emits the short (open-workbook) form'
Assert-True ((New-ProcessTimeExternalRange -Directory "C:\o'brien" -FileName 'B.xlsx' -SheetName 'S' -Column 'O') -match "o''brien") `
    'an apostrophe in the path is doubled per Excel quoting'

# ---------------------------------------------------------------------------
# Get-ProcessTimeCheckColumnSpec -CountReference : the 4-column layout.
# ---------------------------------------------------------------------------
Assert-Equal 4 (@(Get-ProcessTimeCheckColumnSpec -CountReference @{ Enabled = $false }).Count) `
    'a disabled reference still emits all four columns (K as a placeholder)'
Assert-Equal 4 (@(Get-ProcessTimeCheckColumnSpec -CountReference $null).Count) `
    'a null reference still emits all four columns'

$refSpec = @(Get-ProcessTimeCheckColumnSpec -CountReference $r)
Assert-Equal 4 $refSpec.Count 'an enabled reference keeps the four-column layout'
Assert-Equal 'K' $refSpec[2].Col 'K is the reference lookup'
Assert-Equal 'L' $refSpec[3].Col 'L is the count check'
Assert-Equal 11 $refSpec[2].ColIndex 'K is column index 11'
Assert-Equal 12 $refSpec[3].ColIndex 'L is column index 12'
Assert-True (-not [bool]$refSpec[2].IsText) 'a configured reference makes K a live formula, not text'
Assert-Equal $hRefHdr $refSpec[2].Header 'K header is kensu (sanshou)'
Assert-Equal $hCount  $refSpec[3].Header 'L header is kensu-chekku'
Assert-Equal 'I' $refSpec[0].Col 'the duration columns are untouched by the reference'
Assert-Equal 'J' $refSpec[1].Col 'the duration columns are untouched by the reference'

$kRef = New-ProcessTimeCheckFormula -Template $refSpec[2].Formula -Row 3
Assert-True ($kRef -notmatch '\{0\}') 'K lookup formula has no leftover placeholder after filling'
Assert-True ($kRef -match 'LEFT\(C3,7\)&"\*"') 'K lookup prefix-matches the first 7 chars of the correl id'
Assert-True ($kRef -match 'INDEX\(') 'K lookup uses INDEX'
Assert-True ($kRef -match 'MATCH\(') 'K lookup uses MATCH'
Assert-True ($kRef -match 'IFERROR\(') 'K lookup turns an unlisted job into "" rather than #N/A'
Assert-True ($kRef -match [regex]::Escape("'C:\ref\[GPCS(JOD)_7" + $gatsu + ".xlsx]JOD'!`$O`$1:`$O`$20000")) `
    'K lookup returns the value column of the resolved reference workbook'
Assert-True ($kRef -match [regex]::Escape("`$G`$1:`$G`$20000")) 'K lookup matches against the key column'

$lRef = New-ProcessTimeCheckFormula -Template $refSpec[3].Formula -Row 3 -PairRow 10
Assert-True ($lRef -notmatch '\{0\}') 'L formula has no leftover placeholder after filling'
Assert-True ($lRef -match 'G3') 'L compares the OCR-read count in G'
Assert-True ($lRef -match 'K3') 'L compares against the reference count in K'
Assert-True ($lRef -match 'G10') 'L falls back to the paired row when K is blank'
Assert-True ($lRef -match '"OK","NG"') 'L yields OK/NG'
Assert-True ($lRef -match 'TRIM\(G3\)=""') 'L stays blank when the row has no count'

# KeyLength 0 compares the whole correl-id cell instead of a prefix.
$rWhole = Resolve-ProcessTimeCountReference -Reference @{ Enabled = $true; FileName = 'B.xlsx'; SheetName = 'S'; KeyLength = 0 } -Tag 'JOD'
$kWhole = New-ProcessTimeCheckFormula -Template (@(Get-ProcessTimeCheckColumnSpec -CountReference $rWhole)[2].Formula) -Row 4
Assert-True ($kWhole -notmatch 'LEFT\(') 'KeyLength 0 drops the LEFT() prefix match'
Assert-True ($kWhole -match 'MATCH\(C4,') 'KeyLength 0 matches the whole correl-id cell'

exit (Complete-Tests)
