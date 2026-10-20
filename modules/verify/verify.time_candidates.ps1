# modules/verify/verify.time_candidates.ps1
# Which entries of a listing (the Jenkins data/report page) belong to a job
# that ran at a known time, when the names say nothing about the job
# (10/20: the GIFT files of M/SJODJSPQ, QJODJSPP were not in DATA\GIFT;
# they sit on the same report page as F<yyyyMMdd><nnnn> files, and the
# operator finds them by the GIFT run time). With a run time: the entries
# inside [time - before, time + after] are the candidates, and exactly the
# expected number of them (one when the count is unknown) is CERTAIN --
# taken without asking. Otherwise the operator picks: the window first,
# then that day's other entries by hour, as human.pick sets / presets.

. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')     # ConvertTo-EbiDateTime
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # Get-EbiTimeAround, ConvertTo-EbiClockText

$Manifest = @{
  id         = 'verify.time_candidates'
  group      = 'verify'
  summary    = 'Listing entries around a run time: certain when the count fits, else sets to pick from'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    records        = @{ type='list';   required=$true; desc='verify.parse_text records of the listing' }
    date           = @{ type='string'; default=''; desc='the run day (yyyy-MM-dd, yyyy/M/d, ...); blank = no_date' }
    clock          = @{ type='string'; default=''; desc='the run time (HH:mm[:ss] or an Excel day fraction); blank or unreadable = no window, the whole day is offered' }
    beforeMinutes  = @{ type='int';    default=2 }
    afterMinutes   = @{ type='int';    default=13 }
    timeField      = @{ type='string'; default='time' }
    keyField       = @{ type='string'; default='key' }
    namePattern    = @{ type='string'; default=''; desc='regex an entry name must match (blank = any)' }
    expect         = @{ type='int';    default=-1; desc='how many files the job has; -1 = unknown (then one is certain)' }
    expectFallback = @{ type='string'; default=''; desc='used when expect is -1 (a worklist FileCount cell); not a number = none' }
    label          = @{ type='string'; default='run'; desc='what the time is, in the set labels ("GIFT run")' }
  }
  outputs    = @{
    certain  = @{ type='bool';   desc='a run time was given and exactly the expected number of entries are in its window' }
    picked   = @{ type='list';   desc='those names when certain, else empty' }
    window   = @{ type='list';   desc='the names inside the window (empty without a run time)' }
    dayCount = @{ type='int';    desc='entries on that day matching namePattern' }
    sets     = @{ type='list';   desc='human.pick sets: the window, then that day by hour' }
    presets  = @{ type='list';   desc='human.pick presets, the same order' }
    info     = @{ type='list';   desc='lines for the pick panel' }
    reason   = @{ type='string' }
  }
  failures   = @(
    @{ id = 'no_date';    transient = $false }
    @{ id = 'rules_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.time_candidates'; with = @{ records = '{{steps.jkRec.out.records}}'; date = '{{item.GIFT_DATE}}'; clock = '{{item.GIFT_TIME}}'; namePattern = '^F\d{12}\.(csv|txt|zip)$'; expectFallback = '{{item.FileCount}}'; label = 'GIFT run' } }
  notes      = 'no_date: the job has no run day (mapping.xlsx GIFT date blank) -- put the files in by hand, or fill the date and r. Entries are compared by their own time field, to the second.'
}

function VerifyTimeCandidates-Expect {
    param([int]$Expect, [string]$Fallback)
    if ($Expect -ge 0) { return $Expect }
    $t = if ($null -eq $Fallback) { '' } else { $Fallback.Trim() }
    if ($t -match '^\d+$' -and [int]$t -gt 0) { return [int]$t }
    return -1
}

function VerifyTimeCandidates-Find {
    <#
      PURE. The whole decision (unit-tested directly). Returns the outputs
      above plus ok / failure / message.
    #>
    param($Records, [string]$Date, [string]$Clock = '', [int]$BeforeMinutes = 2, [int]$AfterMinutes = 13, [string]$TimeField = 'time', [string]$KeyField = 'key', [string]$NamePattern = '', [int]$Expect = -1, [string]$Label = 'run')
    $none = @{ ok = $true; certain = $false; picked = @(); window = @(); dayCount = 0; sets = @(); presets = @(); info = @(); reason = '' }
    $dm = [regex]::Match(([string]$Date).Trim(), '^(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
    if (-not $dm.Success) {
        $none['ok'] = $false; $none['failure'] = 'no_date'
        $none['message'] = ('no ' + $Label + ' day ("' + $Date + '"): fill it (mapping.xlsx) and r, or put the files in by hand and s')
        return $none
    }
    $day = New-Object DateTime ([int]$dm.Groups[1].Value), ([int]$dm.Groups[2].Value), ([int]$dm.Groups[3].Value)
    $win = $null
    $clockText = ConvertTo-EbiClockText -Value $Clock
    if ($clockText -ne '') {
        $w = Get-EbiTimeAround -Date $day.ToString('yyyy-MM-dd') -Clock $Clock -BeforeMinutes $BeforeMinutes -AfterMinutes $AfterMinutes
        if ($w['ok']) { $win = @{ from = [datetime]::ParseExact($w['from'], 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture); to = [datetime]::ParseExact($w['to'], 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) } }
    }
    # that day's entries, by time
    $list = New-Object System.Collections.ArrayList
    foreach ($r in @($Records)) {
        if (-not ($r -is [System.Collections.IDictionary])) { continue }
        $name = if ($r.Contains($KeyField) -and $null -ne $r[$KeyField]) { ([string]$r[$KeyField]).Trim() } else { '' }
        if ($name -eq '') { continue }
        if ($NamePattern -ne '' -and -not [regex]::IsMatch($name, $NamePattern)) { continue }
        $t = ConvertTo-EbiDateTime -Text ([string]$r[$TimeField])
        if (-not $t['ok'] -or $t['value'].Date -ne $day) { continue }
        [void]$list.Add(@{ name = $name; at = $t['value'] })
    }
    $sorted = @($list.ToArray() | Sort-Object -Property @{ Expression = { $_['at'] } }, @{ Expression = { $_['name'] } })
    $inWin = @(if ($null -ne $win) { foreach ($e in $sorted) { if ($e['at'] -ge $win['from'] -and $e['at'] -le $win['to']) { $e } } })
    $want = if ($Expect -ge 0) { $Expect } else { 1 }
    $certain = ($null -ne $win -and $inWin.Count -gt 0 -and $inWin.Count -eq $want)
    $dayShown = $day.ToString('MM/dd')
    $sets = New-Object System.Collections.ArrayList
    $presets = New-Object System.Collections.ArrayList
    $n = 0
    $winNames = @($inWin | ForEach-Object { [string]$_['name'] })
    if ($null -ne $win) {
        $n++
        $notes = @{}; foreach ($e in $inWin) { $notes[[string]$e['name']] = $e['at'].ToString('HH:mm:ss') }
        $lab = ($Label + ' ' + $dayShown + ' ' + $clockText + ': ' + $win['from'].ToString('HH:mm') + '-' + $win['to'].ToString('HH:mm'))
        [void]$sets.Add(@{ n = $n; id = 'window'; label = $lab; names = $winNames; notes = $notes; default = $true })
        [void]$presets.Add(@{ label = ($lab + $(if ($Expect -ge 0) { ' (expect ' + $Expect + ')' } else { '' })); names = $winNames; hint = '' })
    }
    $byHour = [ordered]@{}
    foreach ($e in $sorted) {
        if ($winNames -contains [string]$e['name']) { continue }
        $h = $e['at'].ToString('HH')
        if (-not $byHour.Contains($h)) { $byHour[$h] = New-Object System.Collections.ArrayList }
        [void]$byHour[$h].Add($e)
    }
    foreach ($h in $byHour.Keys) {
        $n++
        $es = @($byHour[$h])
        $notes = @{}; foreach ($e in $es) { $notes[[string]$e['name']] = $e['at'].ToString('HH:mm:ss') }
        $names = @($es | ForEach-Object { [string]$_['name'] })
        $lab = ($dayShown + ' ' + $h + ':00-' + $h + ':59')
        [void]$sets.Add(@{ n = $n; id = ('hour:' + $h); label = $lab; names = $names; notes = $notes; default = $false })
        [void]$presets.Add(@{ label = $lab; names = $names; hint = (($es | ForEach-Object { $_['at'].ToString('HH:mm:ss') }) -join ' ') })
    }
    $info = New-Object System.Collections.ArrayList
    if ($Clock.Trim() -ne '' -and $clockText -eq '') { [void]$info.Add(($Label + ' time "' + $Clock + '" is not a time: pick by hour')) }
    elseif ($clockText -eq '') { [void]$info.Add(('no ' + $Label + ' time: pick by hour (the Teams start message says when)')) }
    if ($Expect -ge 0) { [void]$info.Add(('the job has ' + $Expect + ' file(s)')) }
    $reason = if ($certain) { ('' + $inWin.Count + ' file(s) in the ' + $Label + ' window: ' + ($winNames -join ', ')) }
        elseif ($null -ne $win) { ('' + $inWin.Count + ' file(s) in the ' + $Label + ' window, ' + $want + ' expected; ' + $sorted.Count + ' on ' + $dayShown) }
        else { ('' + $sorted.Count + ' file(s) on ' + $dayShown + ', no ' + $Label + ' time') }
    return @{ ok = $true; certain = $certain; picked = $(if ($certain) { $winNames } else { @() }); window = $winNames; dayCount = $sorted.Count; sets = $sets.ToArray(); presets = $presets.ToArray(); info = $info.ToArray(); reason = $reason }
}

function Invoke-Step {
    param($In, $Ctx)
    if ([string]$In['namePattern'] -ne '') {
        try { [void][regex]::new([string]$In['namePattern']) } catch { return @{ ok = $false; failure = 'rules_invalid'; message = ('namePattern: ' + $_.Exception.Message); certain = $false; picked = @(); window = @(); dayCount = 0; sets = @(); presets = @(); info = @(); reason = '' } }
    }
    $exp = VerifyTimeCandidates-Expect -Expect ([int]$In['expect']) -Fallback ([string]$In['expectFallback'])
    $r = VerifyTimeCandidates-Find -Records $In['records'] -Date ([string]$In['date']) -Clock ([string]$In['clock']) -BeforeMinutes ([int]$In['beforeMinutes']) -AfterMinutes ([int]$In['afterMinutes']) -TimeField ([string]$In['timeField']) -KeyField ([string]$In['keyField']) -NamePattern ([string]$In['namePattern']) -Expect $exp -Label ([string]$In['label'])
    if (-not $r['ok'] -and $Ctx['DryRun']) { $r['ok'] = $true; $r.Remove('failure'); $r.Remove('message') }
    if ($r['ok']) { $Ctx.Log.Info($r['reason']) }
    return $r
}
