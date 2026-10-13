# modules/verify/verify.nearest_time.ps1
# The start time the operator gave was not on the page (no row in its
# window by the deadline), but rows ARE there: pick the one nearest in
# time and offer it, with its offset, for a human.gate to confirm --
# instead of polling the same page again and again (2026-10-13: a time
# typed 30 minutes off waited the full timeout, twice). Pure: records in,
# a clock out. When the window was found (found = true) nothing changes.

. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # Get-EbiTimeAround, ConvertTo-EbiClockText
. (Join-Path $PSScriptRoot '..\..\kernel\Rules.ps1')     # ConvertTo-EbiDateTime

$Manifest = @{
  id         = 'verify.nearest_time'
  group      = 'verify'
  summary    = 'When the expected time is not on the page, propose the nearest row''s time (a gate confirms)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    found     = @{ type='bool';   default=$true; desc='browser.wait_for found: true = the window matched, keep the clock as is' }
    date      = @{ type='string'; required=$true; desc='yyyy-MM-dd of the clock' }
    clock     = @{ type='string'; required=$true; desc='the start time given (HH:mm[:ss])' }
    records   = @{ type='list';   default=@(); desc='the page rows (verify.parse_text records) to look through' }
    timeField = @{ type='string'; default='startTime' }
    withinMinutes = @{ type='int'; default=5; desc='a row this close counts as "near"; farther rows are still offered, flagged' }
  }
  outputs    = @{
    clock   = @{ type='string'; desc='HH:mm:ss to go on with: the given clock, or the nearest row''s time' }
    changed = @{ type='bool';   desc='true when the nearest row''s time replaced the given clock' }
    code    = @{ type='string'; desc='ok (kept) | unknown (changed: a person confirms)' }
    reason  = @{ type='string'; desc='what was proposed and how far it is from the time given' }
    nearest = @{ type='map';    desc='the nearest record ({} when none)' }
    offsetMinutes = @{ type='int'; desc='nearest - given, in whole minutes (0 when none)' }
    near    = @{ type='bool';   desc='true when the nearest row is within withinMinutes' }
  }
  failures   = @(
    @{ id = 'parse_error'; transient = $false }
  )
  example    = @{ use = 'verify.nearest_time'; with = @{ found = '{{steps.wait.out.found}}'; date = '{{item.GFIX_DATE}}'; clock = '{{steps.start.out.clock}}'; records = '{{steps.rec.out.records}}' } }
  notes      = 'Never silent: a changed clock comes back as code unknown so the workflow''s human.gate asks (Enter = use it, s = skip and rerun with the right time). Rows whose time field does not parse are ignored.'
}

function VerifyNearestTime-Pick {
    <#
      PURE. Records + the given moment -> @{ nearest; offsetMinutes; seconds }
      (nearest = $null when no record has a readable time).
    #>
    param($Records, [datetime]$At, [string]$TimeField)
    $best = $null; $bestSec = [double]::MaxValue
    foreach ($r in @($Records)) {
        if (-not ($r -is [System.Collections.IDictionary]) -or -not $r.Contains($TimeField)) { continue }
        $t = ConvertTo-EbiDateTime -Text ([string]$r[$TimeField]) -Date $At.Date
        if (-not $t['ok']) { continue }
        $d = [Math]::Abs(($t['value'] - $At).TotalSeconds)
        if ($d -lt $bestSec) { $bestSec = $d; $best = @{ record = $r; time = $t['value'] } }
    }
    if ($null -eq $best) { return @{ nearest = $null; offsetMinutes = 0; seconds = 0 } }
    $off = [int][Math]::Round(($best['time'] - $At).TotalMinutes, [System.MidpointRounding]::AwayFromZero)
    return @{ nearest = $best; offsetMinutes = $off; seconds = $bestSec }
}

function Invoke-Step {
    param($In, $Ctx)
    $given = ConvertTo-EbiClockText -Value $In['clock']
    $keep = @{ ok = $true; clock = $given; changed = $false; code = 'ok'; reason = ''; nearest = @{}; offsetMinutes = 0; near = $false }
    if ($given -eq '') {
        if ($Ctx['DryRun']) { return $keep }
        return @{ ok = $false; failure = 'parse_error'; message = ('unreadable time "' + [string]$In['clock'] + '"'); clock = ''; changed = $false; code = 'unknown'; reason = ''; nearest = @{}; offsetMinutes = 0; near = $false }
    }
    $recs = @(@($In['records']) | Where-Object { $null -ne $_ })
    if ([bool]$In['found'] -or $recs.Count -eq 0) { return $keep }
    $w = Get-EbiTimeAround -Date ([string]$In['date']) -Clock $given
    if (-not $w['ok']) { return @{ ok = $false; failure = 'parse_error'; message = $w['message']; clock = $given; changed = $false; code = 'unknown'; reason = ''; nearest = @{}; offsetMinutes = 0; near = $false } }
    $at = [datetime]::ParseExact([string]$w['at'], 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
    $p = VerifyNearestTime-Pick -Records $recs -At $at -TimeField ([string]$In['timeField'])
    if ($null -eq $p['nearest']) { $keep['reason'] = 'no row on the page has a readable ' + [string]$In['timeField']; return $keep }
    $t = [datetime]$p['nearest']['time']
    $near = ($p['seconds'] -le ([int]$In['withinMinutes'] * 60))
    $off = [int]$p['offsetMinutes']
    $offText = if ($off -gt 0) { '+' + $off } else { [string]$off }
    $reason = ('no row at ' + $given + ' by the deadline; the nearest row on the page starts ' + $t.ToString('HH:mm:ss') + ' (' + $offText + ' min' + $(if ($near) { '' } else { ', farther than ' + [int]$In['withinMinutes'] + ' min' }) + '). Enter = go on with ' + $t.ToString('HH:mm:ss') + ' / s = skip this item and rerun with the right time')
    $Ctx.Log.Info($reason)
    return @{ ok = $true; clock = $t.ToString('HH:mm:ss'); changed = $true; code = 'unknown'; reason = $reason; nearest = $p['nearest']['record']; offsetMinutes = $off; near = $near }
}
