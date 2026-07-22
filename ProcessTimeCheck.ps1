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
#     K  kensu-check -- formalizes the operator's manual "count check". FIRST
#        VERSION criterion (documented so the office-PC pass can validate
#        it): the per-row record count (col G) parses to a POSITIVE number
#        once any thousands commas are stripped. Blank G leaves the cell
#        blank (partial row); a zero or non-numeric count reads "NG". A
#        stricter GIFT-vs-GFIX cross-row equality check is a deliberate
#        follow-up, NOT done here: the output layout groups all GIFT rows
#        then all GFIX rows per job, so a GIFT row's paired GFIX row is not
#        at a fixed offset a single-row formula template can reference.
# ---------------------------------------------------------------------------
function Get-ProcessTimeCheckColumnSpec {
    # shori-jikan (kensan) -- "processing time (recheck)"
    $hCalc  = [string][char]0x51E6 + [char]0x7406 + [char]0x6642 + [char]0x9593 + '(' + [char]0x691C + [char]0x7B97 + ')'
    # chekku -- "check"
    $hCheck = [string][char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF
    # kensu-chekku -- "record-count check"
    $hCount = [string][char]0x4EF6 + [char]0x6570 + [char]0x30C1 + [char]0x30A7 + [char]0x30C3 + [char]0x30AF

    $spec = @(
        @{
            Col = 'I'; ColIndex = 9; Header = $hCalc; NumberFormat = '[h]:mm:ss'; Width = 20.0
            Formula = '=IF(AND(ISNUMBER(D{0}),ISNUMBER(E{0})),E{0}-D{0},"")'
        },
        @{
            Col = 'J'; ColIndex = 10; Header = $hCheck; NumberFormat = ''; Width = 8.0
            Formula = '=IF(AND(ISNUMBER(I{0}),ISNUMBER(F{0})),IF(ROUND(F{0}*86400,0)=ROUND(I{0}*86400,0),"T","F"),"")'
        },
        @{
            Col = 'K'; ColIndex = 11; Header = $hCount; NumberFormat = ''; Width = 12.0
            Formula = '=IF(TRIM(G{0})="","",IF(ISNUMBER(VALUE(SUBSTITUTE(G{0},",",""))),IF(VALUE(SUBSTITUTE(G{0},",",""))>0,"OK","NG"),"NG"))'
        }
    )
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
        [Parameter(Mandatory = $true)][int]$Row
    )
    return ($Template -f $Row)
}
