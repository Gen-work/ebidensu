#Requires -Version 5.1
# ============================================================
# PomBook.ps1 - pure sheet-layout plan for POM evidence workbooks
#
# A POM book is an evidence workbook for a job that failed on its first run,
# was fixed, and was run again. Its delivered layout (the JJPCRR12_JJDLWA09 /
# JJPCRR12_JJDSWPA1 samples) carries each GFIX result sheet twice:
#   GFIX<send>result(<before>)  = the first run
#   GFIX<send>result(<after>)   = the rerun after the fix
# plus GFIX<data>_<before-after compare>.
#
# Get-PomBookPlan turns a book's current sheet names into the list of
# actions that bring an old-layout book (YJDSWM39 etc.) into that shape:
#   - GFIX send / recv result -> renamed (<before>), then copied as (<after>)
#     right after it (the copy is a placeholder until the rerun evidence
#     goes in);
#   - GIFT send result is deleted (the POM layout has no GIFT send sheet;
#     operator 10/27, second answer);
#   - GIFT recv result and GIFT-vs-GFIX compare stay where they are;
#   - a new, empty GFIX before/after compare sheet goes at the end.
# The JJPCRR12 sheets are NOT added (only books named after JJPCRR12 have
# them).
#
# Dot-source only (no param()). No COM. ASCII source; Japanese via [char].
# Convert-PomBook.ps1 performs the plan with Excel COM.
# ============================================================

function Get-PomBookLabels {
    $L = Get-ProjectLabels
    # U+5909 U+66F4 = 'henkou' (change); U+524D = mae (before), U+5F8C = go (after)
    $henkou = [char]0x5909 + [char]0x66F4
    $before = $henkou + [char]0x524D
    $after  = $henkou + [char]0x5F8C
    $dchars = [char]0x30C7 + [char]0x30FC + [char]0x30BF   # katakana 'de-ta'
    @{
        GfixSend       = $L['SheetGfixSend']
        GfixRecv       = $L['SheetGfixRecv']
        GiftSend       = $L['SheetGiftSend']
        GiftCompare    = $L['SheetDfCompare']
        GfixSendBefore = $L['SheetGfixSend'] + '(' + $before + ')'
        GfixSendAfter  = $L['SheetGfixSend'] + '(' + $after + ')'
        GfixRecvBefore = $L['SheetGfixRecv'] + '(' + $before + ')'
        GfixRecvAfter  = $L['SheetGfixRecv'] + '(' + $after + ')'
        # 'GFIX de-ta _ henkou mae go hikaku' : U+6BD4 U+8F03 = hikaku (compare)
        PomCompare     = 'GFIX' + $dchars + '_' + $henkou + [char]0x524D + [char]0x5F8C +
                         [char]0x6BD4 + [char]0x8F03
    }
}

# Returns @{ status; actions; missing }
#   status  'convert' : actions brings the book into the POM layout
#           'already' : it is already in the POM layout (nothing to do)
#           'unknown' : neither layout; missing lists the sheets not found
#   actions ordered list of @{ op = 'delete'|'rename'|'copyAfter'|'newCompare'; ... }
#           delete     : sheet
#           rename     : sheet -> name
#           copyAfter  : copy sheet, place right after it, name the copy
#           newCompare : name; template = sheet to copy the look from ('' = blank)
function Get-PomBookPlan {
    param([string[]]$SheetNames)
    $P = Get-PomBookLabels
    $names = @($SheetNames)
    $has = { param($n) $names -contains $n }

    if ((& $has $P.GfixSendBefore) -or (& $has $P.GfixRecvBefore)) {
        return @{ status = 'already'; actions = @(); missing = @() }
    }
    $missing = @(@($P.GfixSend, $P.GfixRecv) | Where-Object { -not (& $has $_) })
    if ($missing.Count -gt 0) {
        return @{ status = 'unknown'; actions = @(); missing = $missing }
    }

    $actions = New-Object System.Collections.ArrayList
    if (& $has $P.GiftSend) {
        [void]$actions.Add(@{ op = 'delete'; sheet = $P.GiftSend })
    }
    foreach ($pair in @(
            @($P.GfixSend, $P.GfixSendBefore, $P.GfixSendAfter),
            @($P.GfixRecv, $P.GfixRecvBefore, $P.GfixRecvAfter))) {
        [void]$actions.Add(@{ op = 'rename';    sheet = $pair[0]; name = $pair[1] })
        [void]$actions.Add(@{ op = 'copyAfter'; sheet = $pair[1]; name = $pair[2] })
    }
    if (-not (& $has $P.PomCompare)) {
        $tpl = ''
        if (& $has $P.GiftCompare) { $tpl = $P.GiftCompare }
        [void]$actions.Add(@{ op = 'newCompare'; name = $P.PomCompare; template = $tpl })
    }
    return @{ status = 'convert'; actions = @($actions); missing = @() }
}

# Sheet order after the plan, for the dry-run listing and the tests.
function Get-PomBookResultOrder {
    param([string[]]$SheetNames, [object[]]$Actions)
    $order = New-Object System.Collections.ArrayList
    foreach ($n in @($SheetNames)) { [void]$order.Add($n) }
    foreach ($a in @($Actions)) {
        switch ($a.op) {
            'delete'     { $order.Remove($a.sheet) }
            'rename'     { $i = $order.IndexOf($a.sheet); if ($i -ge 0) { $order[$i] = $a.name } }
            'copyAfter'  { $i = $order.IndexOf($a.sheet); if ($i -ge 0) { $order.Insert($i + 1, $a.name) } }
            'newCompare' { [void]$order.Add($a.name) }
        }
    }
    return @($order)
}
