# ============================================================
#  kernel/Key.ps1
#
#  Key normalization shared by every place that turns an item's key into
#  text: the template scope (kernel/Context.ps1, {{item.key}} /
#  {{item.keySafe}}), table.load's collision check (P1-24), and the
#  comparison/candidate logic P1-27 adds here later. Dot-source only
#  (no param()), ASCII source, no class.
#
#  Rules are spec/PROFILE-SCHEMA.md 6.6 (P0-R4):
#    display form  : column values joined with " / "
#    file-safe form: each value full-width -> half-width, Windows-illegal
#                    and control characters -> "_", values joined with "_"
#
#  This is the seed of P1-27 (kernel/Key.ps1 + table.key): normalization
#  lives in ONE file so no step ever writes its own -eq. The full-width
#  mapping is the one WorkbookResolver.ps1's class uses (U+FF01..U+FF5E ->
#  U+0021..U+007E, U+3000 -> space), re-stated here as functions because
#  ebi-dance libraries do not use PowerShell classes (STEP-CONTRACT 1.1).
# ============================================================

function ConvertTo-EbiHalfWidth {
    # Full-width ASCII (U+FF01..U+FF5E) and the ideographic space (U+3000)
    # to their half-width equivalents. Everything else is untouched, so
    # Japanese text survives; only the look-alike ASCII is folded.
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $Value.ToCharArray()) {
        $code = [int]$ch
        if ($code -eq 0x3000) {
            [void]$sb.Append(' ')
        } elseif ($code -ge 0xFF01 -and $code -le 0xFF5E) {
            [void]$sb.Append([char]($code - 0xFEE0))
        } else {
            [void]$sb.Append($ch)
        }
    }
    return $sb.ToString()
}

function Test-EbiContainsFullWidth {
    param([string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return $false }
    foreach ($ch in $Value.ToCharArray()) {
        $code = [int]$ch
        if ($code -eq 0x3000 -or ($code -ge 0xFF01 -and $code -le 0xFF5E)) { return $true }
    }
    return $false
}

function ConvertTo-EbiKeySafeSegment {
    # One column value -> its file-name-safe form: half-width first, then
    # every Windows-illegal path character and every control character
    # becomes "_". Nothing is trimmed or lower-cased: keySafe must stay
    # distinguishable from the display form only by these substitutions.
    param([string]$Value)
    $half = ConvertTo-EbiHalfWidth -Value $Value
    if ($half.Length -eq 0) { return '' }
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $half.ToCharArray()) {
        $code = [int]$ch
        $illegal = ($code -lt 0x20) -or ($code -eq 0x7F) -or ('\/:*?"<>|'.IndexOf($ch) -ge 0)
        if ($illegal) { [void]$sb.Append('_') } else { [void]$sb.Append($ch) }
    }
    return $sb.ToString()
}

function Get-EbiKeyValues {
    # The item's key column values in declared order. A missing column
    # reads as '' -- table.load (P1-24) is where a missing key column is
    # an error; the template layer must not throw.
    param($Item, $KeyColumns)
    $out = New-Object System.Collections.ArrayList
    foreach ($col in @($KeyColumns)) {
        $name = [string]$col
        $v = ''
        if ($null -ne $Item -and ($Item -is [System.Collections.IDictionary]) -and $Item.Contains($name) -and $null -ne $Item[$name]) {
            $v = [string]$Item[$name]
        }
        [void]$out.Add($v)
    }
    return $out.ToArray()
}

function Get-EbiKeyDisplay {
    # {{item.key}}: single column -> its raw value; composite -> joined
    # with " / " in declared order (PROFILE-SCHEMA 6.6 a).
    param($Item, $KeyColumns)
    $vals = @(Get-EbiKeyValues -Item $Item -KeyColumns $KeyColumns)
    return ($vals -join ' / ')
}

function ConvertTo-EbiKeySafe {
    # {{item.keySafe}}: each column value made file-safe, joined with "_".
    # NOT unique across rows by construction ("A_B"+"C" and "A"+"B_C"
    # collide) -- table.load checks the whole table for that.
    param($Item, $KeyColumns)
    $vals = @(Get-EbiKeyValues -Item $Item -KeyColumns $KeyColumns)
    $safe = New-Object System.Collections.ArrayList
    foreach ($v in $vals) { [void]$safe.Add((ConvertTo-EbiKeySafeSegment -Value $v)) }
    return ($safe.ToArray() -join '_')
}
