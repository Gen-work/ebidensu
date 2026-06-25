# ============================================================
#  ProjectLabels.ps1
#
#  Every Japanese sheet name / label used at RUNTIME is built here from
#  [char] code points, so the .ps1 files that consume them stay pure
#  ASCII. That is what lets us keep .ps1 BOM-less (per project decision)
#  without risking mojibake: PS 5.1 on a JP-locale host mis-decodes raw
#  UTF-8 in a BOM-less script, but it cannot mis-decode [char]0x9001.
#
#  This source file is intentionally ASCII-only (no raw Japanese, even in
#  comments). Run Check-Encoding.ps1 to print the constructed labels and
#  eyeball them on a real console.
#
#  Dot-source only (no param() block). Returns a hashtable of labels.
# ============================================================

function Get-ProjectLabels {
    $L = @{}

    # ---- sheet names (romaji gloss : code points) ----
    # 'Soushin data'        : U+9001 U+4FE1 + katakana de-(-)-ta
    $L['SheetSoshinData']  = [char]0x9001 + [char]0x4FE1 + [char]0x30C7 + [char]0x30FC + [char]0x30BF
    # 'Soushin =>'          : U+9001 U+4FE1 + U+21D2
    $L['SheetSoshinArrow'] = [char]0x9001 + [char]0x4FE1 + [char]0x21D2
    # 'Jushin =>'           : U+53D7 U+4FE1 + U+21D2
    $L['SheetJushinArrow'] = [char]0x53D7 + [char]0x4FE1 + [char]0x21D2
    # 'GIFT soushin kekka'  : GIFT + U+9001 U+4FE1 U+7D50 U+679C
    $L['SheetGiftSend']    = 'GIFT' + [char]0x9001 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C
    # 'GFIX soushin kekka'
    $L['SheetGfixSend']    = 'GFIX' + [char]0x9001 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C
    # 'GIFT jushin kekka'   : GIFT + U+53D7 U+4FE1 U+7D50 U+679C
    $L['SheetGiftRecv']    = 'GIFT' + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C
    # 'GFIX jushin kekka'
    $L['SheetGfixRecv']    = 'GFIX' + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C
    # 'GIFT de-ta vs GFIX de-ta' : GIFT + de(-)ta + 'vs' + GFIX + de(-)ta
    $dchars = [char]0x30C7 + [char]0x30FC + [char]0x30BF   # katakana 'de-ta'
    $L['SheetDfCompare']   = 'GIFT' + $dchars + 'vs' + 'GFIX' + $dchars

    # ---- inline labels ----
    # bold header above the pasted GFIX log: U+25BC (black down triangle) +
    # 'GFIX' + katakana 'rogu' (U+30ED U+30B0)
    $L['GfixLogLabel']     = [char]0x25BC + 'GFIX' + [char]0x30ED + [char]0x30B0
    # NoGfix section header (spec 8.9 default): 'GFIX' + U+53D7 U+4FE1 +
    # katakana 'fairu' (U+30D5 U+30A1 U+30A4 U+30EB) + hiragana 'nashi'
    # (U+306A U+3057)
    $L['GiftNoGfixHeader'] = 'GFIX' + [char]0x53D7 + [char]0x4FE1 +
        [char]0x30D5 + [char]0x30A1 + [char]0x30A4 + [char]0x30EB +
        [char]0x306A + [char]0x3057
    # NoGfix past-data note (spec 2.5 / F4): kanji 'kako-bun' (U+904E U+53BB
    # U+5206) + katakana 'de-ta-' (U+30C7 U+30FC U+30BF U+30FC). Written to the
    # NoGfixNoteColumn cell next to a past-data screenshot by the Mark phase.
    $L['NoGfixPastData']   = [char]0x904E + [char]0x53BB + [char]0x5206 +
        [char]0x30C7 + [char]0x30FC + [char]0x30BF + [char]0x30FC

    return $L
}

# The five "send-side" sheets that the Align/Precheck phase (spec 6)
# compares against the J4 baseline workbook, in canonical order.
function Get-AlignSendSheets {
    $L = Get-ProjectLabels
    return @(
        $L['SheetSoshinData'],   # Soushin data
        $L['SheetSoshinArrow'],  # Soushin =>
        $L['SheetGiftSend'],     # GIFT soushin kekka
        $L['SheetGfixSend'],     # GFIX soushin kekka
        $L['SheetJushinArrow']   # Jushin =>
    )
}

# The three "receive-side" sheets currently in scope for local testing
# (spec 6: only the last 3 sheets are edited in the local test).
function Get-AlignRecvSheets {
    $L = Get-ProjectLabels
    return @(
        $L['SheetGiftRecv'],   # GIFT jushin kekka
        $L['SheetGfixRecv'],   # GFIX jushin kekka
        $L['SheetDfCompare']   # GIFT de-ta vs GFIX de-ta
    )
}
