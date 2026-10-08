# modules/verify/verify.time_window.ps1
# A time window around a scheduled time: date + clock, minutes before and
# after, as @{ from; to } ISO -- the value a 'within' condition takes.
# The clock may come straight from an Excel cell ("11:44:59.9999999999984025",
# a day fraction), float artefacts rounded to the second (kernel/LogText.ps1).

. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')

$Manifest = @{
  id         = 'verify.time_window'
  group      = 'verify'
  summary    = 'Build a from/to time window around a date and a clock time'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    date          = @{ type='string'; required=$true; desc='yyyy-MM-dd or yyyy/MM/dd (more is ignored)' }
    clock         = @{ type='string'; required=$true; desc='HH:mm[:ss[.fff]] or an Excel day fraction' }
    beforeMinutes = @{ type='int'; default=0 }
    afterMinutes  = @{ type='int'; default=0 }
    format        = @{ type='string'; default='yyyy-MM-dd HH:mm'; desc='.NET format for the *Text outputs (what a page prints, e.g. yyyy/MM/dd HH:mm)' }
  }
  outputs    = @{
    window = @{ type='map';    desc='@{ from; to } ISO yyyy-MM-ddTHH:mm:ss' }
    at     = @{ type='string'; desc='the scheduled moment, ISO' }
    clock  = @{ type='string'; desc='the clock as HH:mm:ss' }
    atText   = @{ type='string'; desc='at, in format' }
    fromText = @{ type='string'; desc='from, in format' }
    toText   = @{ type='string'; desc='to, in format' }
    minuteTexts = @{ type='list'; desc='every minute from..to in format (browser.wait_for containsAny: ready when any of them shows)' }
  }
  failures   = @(
    @{ id = 'parse_error'; transient = $false }
  )
  example    = @{ use = 'verify.time_window'; with = @{ date = '{{run.date}}'; clock = '{{item.GFIX_TIME}}'; beforeMinutes = 2; afterMinutes = 13 } }
}

function Invoke-Step {
    param($In, $Ctx)
    $w = Get-EbiTimeAround -Date ([string]$In['date']) -Clock ([string]$In['clock']) -BeforeMinutes ([int]$In['beforeMinutes']) -AfterMinutes ([int]$In['afterMinutes'])
    if (-not $w['ok']) {
        if ($Ctx['DryRun']) { return @{ ok = $true; window = @{}; at = ''; clock = ''; atText = ''; fromText = ''; toText = ''; minuteTexts = @() } }
        return @{ ok = $false; failure = 'parse_error'; message = $w['message']; window = @{}; at = ''; clock = ''; atText = ''; fromText = ''; toText = ''; minuteTexts = @() }
    }
    $fmt = [string]$In['format']; if ([string]::IsNullOrWhiteSpace($fmt)) { $fmt = 'yyyy-MM-dd HH:mm' }
    $inv = [System.Globalization.CultureInfo]::InvariantCulture
    $f = { param($iso) ([datetime]::ParseExact($iso, 'yyyy-MM-ddTHH:mm:ss', $inv)).ToString($fmt, $inv) }
    $mins = New-Object System.Collections.ArrayList
    $t0 = [datetime]::ParseExact($w['from'], 'yyyy-MM-ddTHH:mm:ss', $inv); $t1 = [datetime]::ParseExact($w['to'], 'yyyy-MM-ddTHH:mm:ss', $inv)
    for ($t = $t0; $t -le $t1 -and $mins.Count -lt 240; $t = $t.AddMinutes(1)) { [void]$mins.Add($t.ToString($fmt, $inv)) }
    return @{ ok = $true; window = @{ from = $w['from']; to = $w['to'] }; at = $w['at']; clock = (ConvertTo-EbiClockText -Value ([string]$In['clock'])); atText = (& $f $w['at']); fromText = (& $f $w['from']); toText = (& $f $w['to']); minuteTexts = $mins.ToArray() }
}
