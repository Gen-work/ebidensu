# modules/verify/verify.assert.ps1
# The rule-table engine (P1-33, PROFILE-SCHEMA.md section 5): rules run in
# order, the first one not satisfied decides, `else` may only be ng or
# unknown (a rule that could conclude ok is refused when the table is
# validated), and a value that cannot be read for a numeric / time op is a
# failed rule, never a pass. Twelve ops, no more: anything else is a new
# verify.* step. The engine lives in kernel/Rules.ps1 (P2-09 moved it) so
# `ebi profile check` judges fixtures with the same code.

. (Join-Path $PSScriptRoot '..\..\kernel\Rules.ps1')   # the rule-table engine (Test-EbiRuleTable / Invoke-EbiRuleTable)

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
  notes      = 'A verdict is never a failure: ng and unknown are outputs (code), and human.gate decides what to ask. Only a malformed rule table fails. A within rule whose value is null or an empty map (no run.timeWindow given) holds: an absent window is not a failed check.'
}

function VerifyAssert-Validate { param($Rules) return (Test-EbiRuleTable -Rules $Rules) }
function VerifyAssert-Run { param($Record, $Rules, [string]$Default) return (Invoke-EbiRuleTable -Record $Record -Rules $Rules -Default $Default) }

function Invoke-Step {
    param($In, $Ctx)
    $v = VerifyAssert-Validate -Rules $In['rules']
    if (-not $v['ok']) { return @{ ok = $false; failure = 'rules_invalid'; message = $v['message']; code = ''; reason = ''; rule = 0; field = '' } }
    $r = VerifyAssert-Run -Record $In['record'] -Rules $v['rules'] -Default $v['default']
    return @{ ok = $true; code = $r['code']; reason = $r['message']; rule = $r['rule']; field = $r['field'] }
}
