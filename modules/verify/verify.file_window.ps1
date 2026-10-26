# modules/verify/verify.file_window.ps1
# Are the files already in a folder the ones a job made at a known time?
# (10/26: DATA\GIFT\KJODWAF1 and LJODWAF1 held files of another GIFT run --
# L's two files were 7 hours apart -- and gift said ok because the folder
# was not empty; track only stopped at the line counts.) A file.list
# listing is checked against the run day / time the plan carried over
# from mapping.xlsx (GIFT_DATE / GIFT_TIME):
#   - how many: the job's FileCount, when known
#   - together: one job's files arrive within a few minutes of each other
#   - the day: the yyyyMMdd in the name (F202609280003.csv), else the file's
#     own date
#   - the time: the file time inside [time - before, time + after], checked
#     only when the file is dated that day (a copy made days later carries
#     the copy's time, not the run's)
# Anything off is code=ng with one line per problem, for a gate.

. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')     # ConvertTo-EbiDateTime
. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # Get-EbiTimeAround, ConvertTo-EbiClockText

$Manifest = @{
  id         = 'verify.file_window'
  group      = 'verify'
  summary    = 'Check files already in a folder against a run day / time (count, together, day, time)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    files          = @{ type='list';   required=$true; desc='file.list files (name, modified)' }
    date           = @{ type='string'; default=''; desc='the run day (yyyy-MM-dd, yyyy/M/d, ...); blank = only count and together are checked' }
    clock          = @{ type='string'; default=''; desc='the run time (HH:mm[:ss] or an Excel day fraction); blank = the time is not checked' }
    beforeMinutes  = @{ type='int';    default=2 }
    afterMinutes   = @{ type='int';    default=60; desc='wider than the download window: a file saved by hand lands a while after the run' }
    spreadMinutes  = @{ type='int';    default=15; desc='one job''s files further apart than this = not one run' }
    nameDate       = @{ type='string'; default='^F(\d{8})'; desc='regex whose group 1 is the yyyyMMdd in a file name; blank = use the file date only' }
    expect         = @{ type='int';    default=-1; desc='how many files the job has; -1 = unknown (not checked)' }
    expectFallback = @{ type='string'; default=''; desc='used when expect is -1 (a worklist FileCount cell); not a positive number = none' }
    label          = @{ type='string'; default='run'; desc='what the time is, in the messages ("GIFT run")' }
  }
  outputs    = @{
    code     = @{ type='string'; desc='ok | ng (no files: ok, nothing to check)' }
    problems = @{ type='list';   desc='one line per problem (empty when ok)' }
    lines    = @{ type='list';   desc='each file as "name  MM/dd HH:mm:ss"' }
    reason   = @{ type='string' }
  }
  failures   = @(
    @{ id = 'rules_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.file_window'; with = @{ files = '{{steps.have.out.files}}'; date = '{{item.GIFT_DATE}}'; clock = '{{item.GIFT_TIME}}'; expectFallback = '{{item.FileCount}}'; label = 'GIFT run' } }
  notes      = 'Pure: decides from the listing only. The file time is LastWriteTime: a copy keeps the original time, a browser download takes the download time, file.download with serverTime takes the server''s.'
}

function VerifyFileWindow-Expect {
    param([int]$Expect, [string]$Fallback)
    if ($Expect -ge 0) { return $Expect }
    $t = if ($null -eq $Fallback) { '' } else { $Fallback.Trim() }
    if ($t -match '^\d+$' -and [int]$t -gt 0) { return [int]$t }
    return -1
}

function VerifyFileWindow-Check {
    <#
      PURE. The whole decision (unit-tested directly). Returns
      @{ ok; code; problems; lines; reason }.
    #>
    param($Files, [string]$Date = '', [string]$Clock = '', [int]$BeforeMinutes = 2, [int]$AfterMinutes = 60, [int]$SpreadMinutes = 15, [string]$NameDate = '^F(\d{8})', [int]$Expect = -1, [string]$Label = 'run')
    $list = New-Object System.Collections.ArrayList
    foreach ($f in @($Files)) {
        if (-not ($f -is [System.Collections.IDictionary])) { continue }
        $t = ConvertTo-EbiDateTime -Text ([string]$f['modified'])
        $nd = $null
        if ($NameDate -ne '') {
            $m = [regex]::Match([string]$f['name'], $NameDate)
            if ($m.Success -and $m.Groups.Count -gt 1) {
                $p = [datetime]::MinValue
                if ([datetime]::TryParseExact($m.Groups[1].Value, 'yyyyMMdd', [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$p)) { $nd = $p.Date }
            }
        }
        [void]$list.Add(@{ name = [string]$f['name']; at = $(if ($t['ok']) { $t['value'] } else { $null }); nameDay = $nd })
    }
    $lines = @(foreach ($e in $list) { $e['name'] + '  ' + $(if ($null -ne $e['at']) { $e['at'].ToString('MM/dd HH:mm:ss') } else { '(no time)' }) })
    if ($list.Count -eq 0) { return @{ ok = $true; code = 'ok'; problems = @(); lines = @(); reason = 'no files: nothing to check' } }

    $problems = New-Object System.Collections.ArrayList
    if ($Expect -gt 0 -and $list.Count -ne $Expect) { [void]$problems.Add(('' + $list.Count + ' file(s) here, the job has ' + $Expect)) }

    $timed = @($list | Where-Object { $null -ne $_['at'] } | Sort-Object -Property @{ Expression = { $_['at'] } })
    if ($timed.Count -gt 1) {
        $first = $timed[0]; $last = $timed[$timed.Count - 1]
        $gap = ($last['at'] - $first['at']).TotalMinutes
        if ($gap -gt $SpreadMinutes) {
            [void]$problems.Add(('not one run: ' + $first['name'] + ' ' + $first['at'].ToString('MM/dd HH:mm') + ' and ' + $last['name'] + ' ' + $last['at'].ToString('MM/dd HH:mm') + ' are ' + [int][Math]::Round($gap) + ' min apart (one job''s files come within ' + $SpreadMinutes + ')'))
        }
    }

    $checked = @('count', 'together')
    $dm = [regex]::Match(([string]$Date).Trim(), '^(\d{4})[-/](\d{1,2})[-/](\d{1,2})')
    if ($dm.Success) {
        $day = New-Object DateTime ([int]$dm.Groups[1].Value), ([int]$dm.Groups[2].Value), ([int]$dm.Groups[3].Value)
        $checked += 'day'
        $win = $null
        $clockText = ConvertTo-EbiClockText -Value $Clock
        if ($clockText -ne '') {
            $w = Get-EbiTimeAround -Date $day.ToString('yyyy-MM-dd') -Clock $Clock -BeforeMinutes $BeforeMinutes -AfterMinutes $AfterMinutes
            if ($w['ok']) {
                $win = @{ from = [datetime]::ParseExact($w['from'], 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture); to = [datetime]::ParseExact($w['to'], 'yyyy-MM-ddTHH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture) }
                $checked += 'time'
            }
        }
        foreach ($e in $list) {
            # the day: the name says it when it can (a copy keeps the name,
            # not always the time); the file date otherwise
            $fileDay = if ($null -ne $e['nameDay']) { $e['nameDay'] } elseif ($null -ne $e['at']) { $e['at'].Date } else { $null }
            if ($null -eq $fileDay) { [void]$problems.Add(($e['name'] + ': no day in the name and no file time')); continue }
            if ($fileDay -ne $day) { [void]$problems.Add(($e['name'] + ' is from ' + $fileDay.ToString('MM/dd') + ', the ' + $Label + ' was ' + $day.ToString('MM/dd'))); continue }
            if ($null -ne $win -and $null -ne $e['at'] -and $e['at'].Date -eq $day -and ($e['at'] -lt $win['from'] -or $e['at'] -gt $win['to'])) {
                [void]$problems.Add(($e['name'] + ' ' + $e['at'].ToString('HH:mm:ss') + ' is outside the ' + $Label + ' ' + $clockText + ' (' + $win['from'].ToString('HH:mm') + '-' + $win['to'].ToString('HH:mm') + ')'))
            }
        }
    }
    $code = if ($problems.Count -gt 0) { 'ng' } else { 'ok' }
    $reason = if ($problems.Count -gt 0) { ($problems.ToArray() -join '; ') }
        else { ('' + $list.Count + ' file(s) fit (' + ($checked -join ', ') + ')' + $(if (-not $dm.Success) { '; no ' + $Label + ' day, so not checked against it' } else { '' })) }
    return @{ ok = $true; code = $code; problems = $problems.ToArray(); lines = $lines; reason = $reason }
}

function Invoke-Step {
    param($In, $Ctx)
    if ([string]$In['nameDate'] -ne '') {
        try { [void][regex]::new([string]$In['nameDate']) } catch { return @{ ok = $false; failure = 'rules_invalid'; message = ('nameDate: ' + $_.Exception.Message); code = 'ng'; problems = @(); lines = @(); reason = '' } }
    }
    $exp = VerifyFileWindow-Expect -Expect ([int]$In['expect']) -Fallback ([string]$In['expectFallback'])
    $r = VerifyFileWindow-Check -Files $In['files'] -Date ([string]$In['date']) -Clock ([string]$In['clock']) -BeforeMinutes ([int]$In['beforeMinutes']) -AfterMinutes ([int]$In['afterMinutes']) -SpreadMinutes ([int]$In['spreadMinutes']) -NameDate ([string]$In['nameDate']) -Expect $exp -Label ([string]$In['label'])
    $Ctx.Log.Info($r['reason'])
    return $r
}
