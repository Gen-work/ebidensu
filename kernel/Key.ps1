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

# ============================================================
#  P1-27: matching. Everything below is the ONE comparison rule set.
#  No step under modules/ writes its own -eq against a key: file.find,
#  table.key, table.set, flow.checkpoint and verify.match_record all call
#  Get-EbiKeyMatchTier / Find-EbiKeyMatches, and the ambiguity shape they
#  hand to human.choose comes from New-EbiCandidateList.
#
#  Rules are PROFILE-SCHEMA.md 6.2's confirmedRules, grown at runtime:
#      { kind = 'suffix'; pattern = '<regex>' }   strip a suffix before comparing
#      { kind = 'fullwidth' }                     fold full-width ASCII
#      { kind = 'case-insensitive' }              ignore case
#      { kind = 'prefix'; pattern = '<regex>' }   strip a prefix
#  Tiers, best first: exact > suffix/prefix stripped > full-width folded >
#  case-folded (each later tier applies the earlier normalizations too).
#  The stamped batch-run suffix '.yymmdd.hhmmssff' is the one rule the
#  project always had (MappingStore.ps1 Get-CorrelIdAliases), so it is
#  in the default rule set whenever a profile declares none.
# ============================================================

function Get-EbiKeyDefaultRules {
    return @(
        @{ kind = 'suffix'; pattern = '\.\d{6}\.\d{8}$'; note = 'transfer batch timestamp suffix' }
        @{ kind = 'fullwidth' }
        @{ kind = 'case-insensitive' }
    )
}

function Get-EbiKeyRules {
    # profile.worklist.key.confirmedRules, or the defaults when absent.
    param($Profile)
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary])) {
        $wl = $Profile['worklist']
        if ($wl.Contains('key') -and ($wl['key'] -is [System.Collections.IDictionary]) -and $wl['key'].Contains('confirmedRules') -and $null -ne $wl['key']['confirmedRules']) {
            $out = New-Object System.Collections.ArrayList
            foreach ($r in $wl['key']['confirmedRules']) { if ($r -is [System.Collections.IDictionary] -and $r.Contains('kind')) { [void]$out.Add($r) } }
            return $out.ToArray()
        }
    }
    return (Get-EbiKeyDefaultRules)
}

function Get-EbiKeyTierOrder { return @('exact', 'stripped', 'fullwidth', 'case') }

function ConvertTo-EbiKeyForm {
    <#
      PURE. A value normalized up to a tier:
        exact      trimmed only
        stripped   suffix / prefix rules applied
        fullwidth  + full-width folded (only when a fullwidth rule exists)
        case       + lower-cased (only when a case-insensitive rule exists)
      A tier whose rule is not declared yields the same text as the tier
      before it, so it can never produce a new match.
    #>
    param([string]$Value, [string]$Tier, $Rules)
    $v = if ($null -eq $Value) { '' } else { $Value.Trim() }
    if ($Tier -eq 'exact') { return $v }
    foreach ($r in @($Rules)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $kind = [string]$r['kind']
        if ($kind -eq 'suffix' -and $r.Contains('pattern') -and [string]$r['pattern'] -ne '') { $v = [regex]::Replace($v, [string]$r['pattern'], '') }
        if ($kind -eq 'prefix' -and $r.Contains('pattern') -and [string]$r['pattern'] -ne '') { $v = [regex]::Replace($v, [string]$r['pattern'], '') }
    }
    if ($Tier -eq 'stripped') { return $v }
    $hasFull = $false; $hasCase = $false
    foreach ($r in @($Rules)) { if ($r -is [System.Collections.IDictionary]) { if ([string]$r['kind'] -eq 'fullwidth') { $hasFull = $true }; if ([string]$r['kind'] -eq 'case-insensitive') { $hasCase = $true } } }
    if ($hasFull) { $v = ConvertTo-EbiHalfWidth -Value $v }
    if ($Tier -eq 'fullwidth') { return $v }
    if ($hasCase) { $v = $v.ToLowerInvariant() }
    return $v
}

function Get-EbiKeyMatchTier {
    # PURE. '' when the value is not the key under any rule, else the best
    # tier ('exact' | 'stripped' | 'fullwidth' | 'case'). Ordinal compare.
    param([string]$Value, [string]$Key, $Rules = $null)
    if ($null -eq $Rules) { $Rules = Get-EbiKeyDefaultRules }
    if ([string]::IsNullOrWhiteSpace($Key) -or [string]::IsNullOrWhiteSpace($Value)) { return '' }
    foreach ($tier in (Get-EbiKeyTierOrder)) {
        $a = ConvertTo-EbiKeyForm -Value $Value -Tier $tier -Rules $Rules
        $b = ConvertTo-EbiKeyForm -Value $Key -Tier $tier -Rules $Rules
        if ([string]::Equals($a, $b, [System.StringComparison]::Ordinal)) { return $tier }
    }
    return ''
}

function Get-EbiKeyOfRecord {
    # PURE. The comparable key text of a record (hashtable) under the key
    # columns: one column -> its value; several -> joined with ' / ' (the
    # display form). A record that is itself a string is its own key.
    param($Record, $KeyColumns)
    if ($Record -is [string]) { return $Record }
    return (Get-EbiKeyDisplay -Item $Record -KeyColumns $KeyColumns)
}

function Get-EbiKeyPartsTier {
    <#
      PURE. A composite key compares COLUMN BY COLUMN: each value against
      its part of the key at the same tier (a suffix rule anchored with $
      must see the column's end, not the ' / '-joined text). The row's
      tier is the worst tier any column needed; '' if any column fails.
    #>
    param($Values, $Parts, $Rules)
    $vals = @($Values); $parts = @($Parts)
    if ($vals.Count -ne $parts.Count -or $vals.Count -eq 0) { return '' }
    $order = Get-EbiKeyTierOrder
    $worst = 0
    for ($i = 0; $i -lt $vals.Count; $i++) {
        $t = Get-EbiKeyMatchTier -Value ([string]$vals[$i]) -Key ([string]$parts[$i]) -Rules $Rules
        if ($t -eq '') { return '' }
        $n = [array]::IndexOf($order, $t)
        if ($n -gt $worst) { $worst = $n }
    }
    return $order[$worst]
}

function Find-EbiKeyMatches {
    <#
      PURE. Records (hashtables, or plain strings) + a key -> the best tier
      that has hits and every record in it:
        @{ tier; indexes = int[] (0-based); records = object[] }
      Composite keys compare column by column at the same tier, through
      the same normalization (' / ' joined form), never by a step's own
      -eq. tier '' and empty arrays when nothing matches.
    #>
    param($Records, [string]$Key, $KeyColumns = @(), $Rules = $null)
    if ($null -eq $Rules) { $Rules = Get-EbiKeyDefaultRules }
    $best = ''
    $byTier = @{}
    foreach ($t in (Get-EbiKeyTierOrder)) { $byTier[$t] = New-Object System.Collections.ArrayList }
    $cols = @($KeyColumns)
    $parts = @($Key -split ' / ')
    $i = -1
    foreach ($rec in @($Records)) {
        $i++
        if ($null -eq $rec) { continue }
        $tier = ''
        if (($rec -is [System.Collections.IDictionary]) -and $cols.Count -gt 0 -and $parts.Count -eq $cols.Count) {
            $tier = Get-EbiKeyPartsTier -Values (Get-EbiKeyValues -Item $rec -KeyColumns $cols) -Parts $parts -Rules $Rules
        } else {
            $tier = Get-EbiKeyMatchTier -Value (Get-EbiKeyOfRecord -Record $rec -KeyColumns $cols) -Key $Key -Rules $Rules
        }
        if ($tier -ne '') { [void]$byTier[$tier].Add($i) }
    }
    foreach ($t in (Get-EbiKeyTierOrder)) { if ($byTier[$t].Count -gt 0) { $best = $t; break } }
    $idx = New-Object System.Collections.ArrayList
    $recs = New-Object System.Collections.ArrayList
    if ($best -ne '') {
        $all = @($Records)
        foreach ($n in $byTier[$best]) { [void]$idx.Add([int]$n); [void]$recs.Add($all[$n]) }
    }
    return @{ tier = $best; indexes = $idx.ToArray(); records = $recs.ToArray() }
}

function Find-EbiKeySafeCollisions {
    <#
      PURE. Rows + key columns -> @{ ok; collisions = @(@{ keySafe; keys }) }.
      The file-safe form is not unique by construction (PROFILE-SCHEMA 6.6
      a); table.load refuses a table where two rows share one, because
      capture/<page>/<keySafe>.png would then overwrite silently.
    #>
    param($Rows, $KeyColumns)
    $seen = @{}
    $order = New-Object System.Collections.ArrayList
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $safe = ConvertTo-EbiKeySafe -Item $row -KeyColumns $KeyColumns
        if (-not $seen.Contains($safe)) { $seen[$safe] = New-Object System.Collections.ArrayList; [void]$order.Add($safe) }
        [void]$seen[$safe].Add((Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns))
    }
    $out = New-Object System.Collections.ArrayList
    foreach ($safe in $order) { if ($seen[$safe].Count -gt 1) { [void]$out.Add(@{ keySafe = $safe; keys = $seen[$safe].ToArray() }) } }
    return @{ ok = ($out.Count -eq 0); collisions = $out.ToArray() }
}

function New-EbiCandidateList {
    <#
      PURE. The P0-R4 standard candidate shape (PROFILE-SCHEMA 6.6 c):
        @{ candidates = @(@{ id='c1'; candidate; evidence=@{...} }, ...);
           suggestion = @{ id; reason } | $null; doubts = '' }
      Items are @{ candidate = <text>; evidence = <map> } in the order the
      caller ranked them; SuggestIndex (0-based) names the one the caller
      would pick, -1 for none. Evidence is whatever the caller really
      knows -- never a made-up value.
    #>
    param($Items, [int]$SuggestIndex = -1, [string]$Reason = '', [string]$Doubts = '')
    $list = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($it in @($Items)) {
        $i++
        $ev = if ($it -is [System.Collections.IDictionary] -and $it.Contains('evidence') -and ($it['evidence'] -is [System.Collections.IDictionary])) { $it['evidence'] } else { @{} }
        $cand = if ($it -is [System.Collections.IDictionary] -and $it.Contains('candidate')) { [string]$it['candidate'] } else { [string]$it }
        [void]$list.Add(@{ id = ('c' + $i); candidate = $cand; evidence = $ev })
    }
    $sugg = $null
    if ($SuggestIndex -ge 0 -and $SuggestIndex -lt $list.Count) { $sugg = @{ id = ('c' + ($SuggestIndex + 1)); reason = $Reason } }
    return @{ candidates = $list.ToArray(); suggestion = $sugg; doubts = $Doubts }
}
