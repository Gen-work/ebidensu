#Requires -Version 5.1
# ============================================================
#  kernel/Rules.ps1
#
#  The rule-table engine of PROFILE-SCHEMA.md section 5 (verify.assert,
#  P1-33; moved out of the step in P2-09 so `ebi profile check` runs
#  fixtures through the same code). Dot-source only (no param(), ASCII
#  source). Pure.
#
#    Get-EbiRuleOps        the twelve ops, no more
#    Test-EbiRuleTable     -> @{ ok; message; rules; default }  (else never ok)
#    Test-EbiRuleHolds     one rule against one record; unreadable -> $false,
#                          except a within rule with no window at all -> $true
#    Invoke-EbiRuleTable   -> @{ code; message; rule; field }
# ============================================================

. (Join-Path $PSScriptRoot 'Parse.ps1')   # ConvertTo-EbiDateTime

function Get-EbiRuleOps { return @('equals', 'notEquals', 'in', 'notIn', 'matches', 'present', 'empty', 'within', 'gt', 'lt', 'gte', 'lte') }

function Test-EbiRuleTable {
    # PURE. -> @{ ok; message; rules (array); default }
    param($Rules)
    if ($null -eq $Rules -or -not ($Rules -is [System.Collections.IDictionary])) { return @{ ok = $false; message = 'rules must be a map { rules: [...], default: ok }'; rules = @(); default = 'ok' } }
    $list = @(if ($Rules.Contains('rules') -and $null -ne $Rules['rules']) { $Rules['rules'] })
    $default = if ($Rules.Contains('default') -and $null -ne $Rules['default']) { [string]$Rules['default'] } else { 'ok' }
    if (-not ($default -in @('ok', 'ng', 'unknown'))) { return @{ ok = $false; message = ('default "' + $default + '" must be ok, ng or unknown'); rules = @(); default = $default } }
    $i = 0
    foreach ($r in $list) {
        $i++
        if (-not ($r -is [System.Collections.IDictionary])) { return @{ ok = $false; message = ('rule ' + $i + ' is not a map'); rules = @(); default = $default } }
        $op = if ($r.Contains('op')) { [string]$r['op'] } else { '' }
        if (-not ((Get-EbiRuleOps) -contains $op)) { return @{ ok = $false; message = ('rule ' + $i + ': op "' + $op + '" is not one of ' + ((Get-EbiRuleOps) -join ', ')); rules = @(); default = $default } }
        if (-not $r.Contains('field') -or [string]::IsNullOrWhiteSpace([string]$r['field'])) { return @{ ok = $false; message = ('rule ' + $i + ' has no field'); rules = @(); default = $default } }
        $else = if ($r.Contains('else')) { [string]$r['else'] } else { '' }
        if (-not ($else -in @('ng', 'unknown'))) { return @{ ok = $false; message = ('rule ' + $i + ' (' + [string]$r['field'] + '): else must be ng or unknown, never ok (PROFILE-SCHEMA 5.2); got "' + $else + '"'); rules = @(); default = $default } }
        if ($op -in @('equals', 'notEquals', 'in', 'notIn', 'matches', 'within', 'gt', 'lt', 'gte', 'lte') -and -not $r.Contains('value')) { return @{ ok = $false; message = ('rule ' + $i + ' (' + $op + ') needs a value'); rules = @(); default = $default } }
    }
    return @{ ok = $true; message = ''; rules = $list; default = $default }
}

function ConvertTo-EbiRuleNumber {
    param([string]$Text)
    $t = if ($null -eq $Text) { '' } else { $Text.Trim().Replace(',', '') }
    $d = 0.0
    if ($t -ne '' -and [double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return @{ ok = $true; value = $d } }
    return @{ ok = $false; value = 0.0 }
}

function Test-EbiRuleHolds {
    # PURE. Does one rule hold for the record? Unreadable -> $false.
    param($Record, $Rule)
    $field = [string]$Rule['field']
    $raw = if ($Record -is [System.Collections.IDictionary] -and $Record.Contains($field) -and $null -ne $Record[$field]) { [string]$Record[$field] } else { '' }
    $v = if ($Rule.Contains('value')) { $Rule['value'] } else { $null }
    switch ([string]$Rule['op']) {
        'present'   { return (-not [string]::IsNullOrWhiteSpace($raw)) }
        'empty'     { return ([string]::IsNullOrWhiteSpace($raw)) }
        'equals'    { return ([string]::Equals($raw.Trim(), ([string]$v).Trim(), [System.StringComparison]::Ordinal)) }
        'notEquals' { return (-not [string]::Equals($raw.Trim(), ([string]$v).Trim(), [System.StringComparison]::Ordinal)) }
        'in'        { foreach ($x in @($v)) { if ([string]::Equals($raw.Trim(), ([string]$x).Trim(), [System.StringComparison]::Ordinal)) { return $true } }; return $false }
        'notIn'     { foreach ($x in @($v)) { if ([string]::Equals($raw.Trim(), ([string]$x).Trim(), [System.StringComparison]::Ordinal)) { return $false } }; return $true }
        'matches'   { try { return [regex]::IsMatch($raw, [string]$v) } catch { return $false } }
        'within'    {
            # No window declared (run.timeWindow still null: nobody gave one)
            # is not a failed check but an absent one -- the old Test-MqRecord
            # skipped its time check when Expected was null. An unreadable
            # FIELD value never passes; an absent WINDOW passes.
            if ($null -eq $v -or -not ($v -is [System.Collections.IDictionary]) -or $v.Count -eq 0) { return $true }
            $t = ConvertTo-EbiDateTime -Text $raw
            if (-not $t['ok']) { return $false }
            $from = $null; $to = $null
            if ($v.Contains('from') -and [string]$v['from'] -ne '') { $p = ConvertTo-EbiDateTime -Text ([string]$v['from']); if (-not $p['ok']) { return $false }; $from = $p['value'] }
            if ($v.Contains('to') -and [string]$v['to'] -ne '') { $p = ConvertTo-EbiDateTime -Text ([string]$v['to']); if (-not $p['ok']) { return $false }; $to = $p['value'] }
            if ($null -eq $from -and $null -eq $to) { return $false }
            return (($null -eq $from -or $t['value'] -ge $from) -and ($null -eq $to -or $t['value'] -le $to))
        }
        default {
            $a = ConvertTo-EbiRuleNumber -Text $raw; $b = ConvertTo-EbiRuleNumber -Text ([string]$v)
            if (-not $a['ok'] -or -not $b['ok']) { return $false }
            switch ([string]$Rule['op']) {
                'gt'  { return ($a['value'] -gt $b['value']) }
                'lt'  { return ($a['value'] -lt $b['value']) }
                'gte' { return ($a['value'] -ge $b['value']) }
                'lte' { return ($a['value'] -le $b['value']) }
            }
            return $false
        }
    }
}

function Invoke-EbiRuleTable {
    # PURE. -> @{ code; message; rule; field }
    param($Record, $Rules, [string]$Default)
    $i = 0
    foreach ($r in @($Rules)) {
        $i++
        if (Test-EbiRuleHolds -Record $Record -Rule $r) { continue }
        $msg = if ($r.Contains('message') -and $null -ne $r['message']) { [string]$r['message'] } else { ([string]$r['field'] + ' ' + [string]$r['op'] + ' failed') }
        return @{ code = [string]$r['else']; message = $msg; rule = $i; field = [string]$r['field'] }
    }
    return @{ code = $Default; message = ''; rule = 0; field = '' }
}

