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
Assert-Equal 3 $spec.Count 'three audit columns (I/J/K)'

Assert-Equal 'I' $spec[0].Col 'first audit column is I'
Assert-Equal 'J' $spec[1].Col 'second audit column is J'
Assert-Equal 'K' $spec[2].Col 'third audit column is K'
Assert-Equal 9  $spec[0].ColIndex 'I is column index 9 (after the A..H data columns)'
Assert-Equal 10 $spec[1].ColIndex 'J is column index 10'
Assert-Equal 11 $spec[2].ColIndex 'K is column index 11'

# Headers are built from [char] code points (ASCII source, codepage-agnostic).
$hCalc  = [string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593 + '(' + [char]0x691C + [char]0x7B97 + ')'  # shori-jikan (kensan)
$hCheck = [string][char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF                                          # chekku
$hCount = [string][char]0x4EF6 + [char]0x6570 + [char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF            # kensu-chekku
Assert-Equal $hCalc  $spec[0].Header 'I header is shori-jikan (kensan)'
Assert-Equal $hCheck $spec[1].Header 'J header is chekku'
Assert-Equal $hCount $spec[2].Header 'K header is kensu-chekku'

Assert-Equal '[h]:mm:ss' $spec[0].NumberFormat 'I column carries a duration number format'
Assert-Equal '' $spec[1].NumberFormat 'J column leaves the default format'
Assert-Equal '' $spec[2].NumberFormat 'K column leaves the default format'

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

$kFilled = New-ProcessTimeCheckFormula -Template $spec[2].Formula -Row 5
Assert-True ($kFilled -notmatch '\{0\}') 'K (count) formula has no leftover placeholder after filling'
Assert-True ($kFilled -match 'G5') 'K formula checks the record-count cell G5'
Assert-True ($kFilled -match '"OK"') 'K formula reports OK for a valid count'
Assert-True ($kFilled -match '"NG"') 'K formula reports NG for a missing/zero count'

exit (Complete-Tests)
