# modules/verify/verify.assert.ps1
# The rule-table engine (P1-33, PROFILE-SCHEMA.md section 5): rules run in
# order, the first one not satisfied decides, `else` may only be ng or
# unknown (a rule that could conclude ok is refused when the table is
# validated), and a value that cannot be read for a numeric / time op is a
# failed rule, never a pass. Twelve ops, no more: anything else is a new
# verify.* step.

. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')   # ConvertTo-EbiDateTime

$Manifest = @{
  id         = 'verify.assert'
  group      = 'verify'
  summary    = 'Run a rule table over a record: ok / ng / unknown plus the message'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    record = @{ type='map'; required=$true; desc='the record to judge (field -> text)' }
    rules  = @{ type='map'; required=$true; desc='{ rules: [ { field, op, value?, else, message } ], default: ok }' }
  }
  outputs    = @{
    code    = @{ type='string'; desc='ok | ng | unknown' }
    reason  = @{ type='string'; desc='the failing rule''s message, for people (named reason: message is a reserved return key)' }
    rule    = @{ type='int';    desc='1-based index of the rule that decided; 0 when every rule passed' }
    field   = @{ type='string'; desc='that rule''s field' }
  }
  failures   = @(
    @{ id = 'rules_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.assert'; with = @{ record = '{{steps.row.out.record}}'; rules = '{{page.rules}}' } }
  notes      = 'A verdict is never a failure: ng and unknown are outputs (code), and human.gate decides what to ask. Only a malformed rule table fails.'
}

function VerifyAssert-Ops { return @('equals', 'notEquals', 'in', 'notIn', 'matches', 'present', 'empty', 'within', 'gt', 'lt', 'gte', 'lte') }

function VerifyAssert-Validate {
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
        if (-not ((VerifyAssert-Ops) -contains $op)) { return @{ ok = $false; message = ('rule ' + $i + ': op "' + $op + '" is not one of ' + ((VerifyAssert-Ops) -join ', ')); rules = @(); default = $default } }
        if (-not $r.Contains('field') -or [string]::IsNullOrWhiteSpace([string]$r['field'])) { return @{ ok = $false; message = ('rule ' + $i + ' has no field'); rules = @(); default = $default } }
        $else = if ($r.Contains('else')) { [string]$r['else'] } else { '' }
        if (-not ($else -in @('ng', 'unknown'))) { return @{ ok = $false; message = ('rule ' + $i + ' (' + [string]$r['field'] + '): else must be ng or unknown, never ok (PROFILE-SCHEMA 5.2); got "' + $else + '"'); rules = @(); default = $default } }
        if ($op -in @('equals', 'notEquals', 'in', 'notIn', 'matches', 'within', 'gt', 'lt', 'gte', 'lte') -and -not $r.Contains('value')) { return @{ ok = $false; message = ('rule ' + $i + ' (' + $op + ') needs a value'); rules = @(); default = $default } }
    }
    return @{ ok = $true; message = ''; rules = $list; default = $default }
}

function VerifyAssert-Number {
    param([string]$Text)
    $t = if ($null -eq $Text) { '' } else { $Text.Trim().Replace(',', '') }
    $d = 0.0
    if ($t -ne '' -and [double]::TryParse($t, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$d)) { return @{ ok = $true; value = $d } }
    return @{ ok = $false; value = 0.0 }
}

function VerifyAssert-Holds {
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
            $t = ConvertTo-EbiDateTime -Text $raw
            if (-not $t['ok'] -or $null -eq $v -or -not ($v -is [System.Collections.IDictionary])) { return $false }
            $from = $null; $to = $null
            if ($v.Contains('from') -and [string]$v['from'] -ne '') { $p = ConvertTo-EbiDateTime -Text ([string]$v['from']); if (-not $p['ok']) { return $false }; $from = $p['value'] }
            if ($v.Contains('to') -and [string]$v['to'] -ne '') { $p = ConvertTo-EbiDateTime -Text ([string]$v['to']); if (-not $p['ok']) { return $false }; $to = $p['value'] }
            if ($null -eq $from -and $null -eq $to) { return $false }
            return (($null -eq $from -or $t['value'] -ge $from) -and ($null -eq $to -or $t['value'] -le $to))
        }
        default {
            $a = VerifyAssert-Number -Text $raw; $b = VerifyAssert-Number -Text ([string]$v)
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

function VerifyAssert-Run {
    # PURE. -> @{ code; message; rule; field }
    param($Record, $Rules, [string]$Default)
    $i = 0
    foreach ($r in @($Rules)) {
        $i++
        if (VerifyAssert-Holds -Record $Record -Rule $r) { continue }
        $msg = if ($r.Contains('message') -and $null -ne $r['message']) { [string]$r['message'] } else { ([string]$r['field'] + ' ' + [string]$r['op'] + ' failed') }
        return @{ code = [string]$r['else']; message = $msg; rule = $i; field = [string]$r['field'] }
    }
    return @{ code = $Default; message = ''; rule = 0; field = '' }
}

function Invoke-Step {
    param($In, $Ctx)
    $v = VerifyAssert-Validate -Rules $In['rules']
    if (-not $v['ok']) { return @{ ok = $false; failure = 'rules_invalid'; message = $v['message']; code = ''; reason = ''; rule = 0; field = '' } }
    $r = VerifyAssert-Run -Record $In['record'] -Rules $v['rules'] -Default $v['default']
    return @{ ok = $true; code = $r['code']; reason = $r['message']; rule = $r['rule']; field = $r['field'] }
}
