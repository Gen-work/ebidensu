# modules/verify/verify.assert_each.ps1
# verify.assert over a LIST: the same rule table (PROFILE-SCHEMA.md section
# 5, kernel/Rules.ps1) run on every record, and the worst verdict wins
# (ng > unknown > ok). For "did any row of this job fail?" -- 2026-10-13
# JJDSWM41: GoAnywhere showed the Send row with a red status icon; the icon
# is not in the page text, but the view's resolved column (added by the
# operator) reads "No" on exactly that row and blank on every other.

. (Join-Path $PSScriptRoot '..\..\kernel\Rules.ps1')

$Manifest = @{
  id         = 'verify.assert_each'
  group      = 'verify'
  summary    = 'Run a rule table over every record; the worst verdict wins (ng > unknown > ok)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    records  = @{ type='list';   required=$true; desc='the records to judge (verify.filter_records / parse_text output)' }
    rules    = @{ type='map';    required=$true; desc='{ rules: [ { field, op, value?, else, message } ], default: ok } -- the verify.assert shape' }
    keyField = @{ type='string'; default='key'; desc='field named in the reason for a record that failed' }
  }
  outputs    = @{
    code        = @{ type='string'; desc='ok | ng | unknown: the worst verdict over the records (ok when there are none)' }
    reason      = @{ type='string'; desc='the first failing record''s key and its rule message; "" when ok' }
    failed      = @{ type='list';   desc='@{ key; code; reason; record } per record that did not pass, in list order' }
    judged      = @{ type='int';    desc='records judged' }
    failedCount = @{ type='int' }
  }
  failures   = @(
    @{ id = 'rules_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.assert_each'; with = @{ records = '{{steps.mine.out.records}}'; rules = '{{profile.rules.goAnywhereJob}}' } }
  notes      = 'An empty list is ok (nothing failed) -- pair it with a filter step that fails not_found when the rows are missing. The same rule table as verify.assert, so a profile keeps one vocabulary of rules.'
}

function VerifyAssertEach-Worst {
    # PURE. Two codes -> the worse one (ng > unknown > ok > '').
    param([string]$A, [string]$B)
    $rank = @{ '' = 0; 'ok' = 1; 'unknown' = 2; 'ng' = 3 }
    $ra = if ($rank.Contains($A)) { $rank[$A] } else { 0 }
    $rb = if ($rank.Contains($B)) { $rank[$B] } else { 0 }
    return $(if ($rb -gt $ra) { $B } else { $A })
}

function Invoke-Step {
    param($In, $Ctx)
    $v = Test-EbiRuleTable -Rules $In['rules']
    if (-not $v['ok']) { return @{ ok = $false; failure = 'rules_invalid'; message = $v['message']; code = ''; reason = ''; failed = @(); judged = 0; failedCount = 0 } }
    $recs = @(@($In['records']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $kf = [string]$In['keyField']
    $code = 'ok'; $reason = ''
    $failed = New-Object System.Collections.ArrayList
    foreach ($r in $recs) {
        $x = Invoke-EbiRuleTable -Record $r -Rules $v['rules'] -Default $v['default']
        if ([string]$x['code'] -eq 'ok') { continue }
        $k = if ($kf -ne '' -and $r.Contains($kf)) { [string]$r[$kf] } else { '' }
        [void]$failed.Add(@{ key = $k; code = [string]$x['code']; reason = [string]$x['message']; record = $r })
        $code = VerifyAssertEach-Worst -A $code -B ([string]$x['code'])
        if ($reason -eq '') { $reason = $(if ($k -ne '') { $k + ': ' } else { '' }) + [string]$x['message'] }
    }
    if ($failed.Count -gt 1) { $reason = $reason + ' (+' + ($failed.Count - 1) + ' more)' }
    return @{ ok = $true; code = $code; reason = $reason; failed = $failed.ToArray(); judged = $recs.Count; failedCount = $failed.Count }
}
