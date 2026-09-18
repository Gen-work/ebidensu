# ============================================================
#  GiftMqProcessTime.ps1
#
#  PURE library for the standalone GiftMqProcessTime.ps1 driver -- NO
#  Excel COM, NO SendKeys, NO file I/O. Dot-source only (no param()
#  block; ASCII source, Japanese via [char]). Unit-tested by
#  Tests\Test-GiftMqProcessTime.ps1.
#
#  What it models (the operator's daily hand procedure, see
#  docs/GiftMqProcessTime.md):
#    1. the leader's daily job list lands in the operator's own
#       mapping.xlsx as JOB / owner / GIFT run date / GIFT scheduled time;
#    2. the GIFT MQ "Transfer status inquiry results" LIST page (Ctrl+A
#       text) gives every transfer's Send date = the job's START time;
#    3. each record's Detail page gives INSERTDATETIME = the END time;
#    4. the processed record count only exists in the team's Teams chat
#       ("job XXXXWXXX ... (soushin-yotei: N ken)");
#    5. everything goes into the shori-jikan (<label>(<Tag>).xlsx) sheet:
#       No. / GIFT-GFIX / correl / start / end / duration / count / job.
#
#  The scheduled time is only a window: a job may start up to ~10 min
#  early or late, so matching uses the schedule as a centre and the
#  day's ORDER as the tie-break (Resolve-GiftMqJobMatches).
#
#  Convention: functions return plain arrays -- never ,@(...) (callers
#  wrap calls in @() and that nests in PS 5.1).
# ============================================================

# ---------------------------------------------------------------------------
# Get-GiftMqLabels
#   Japanese literals this feature needs, built from code points so the
#   source stays ASCII (CLAUDE.md encoding rule).
# ---------------------------------------------------------------------------
function Get-GiftMqLabels {
    $L = @{}
    # 'jobu' (job) : katakana ji-yo-bu  U+30B8 U+30E7 U+30D6
    $L['Job']        = [char]0x30B8 + [char]0x30E7 + [char]0x30D6
    # 'tantou' (owner / person in charge) : U+62C5 U+5F53
    $L['Owner']      = [char]0x62C5 + [char]0x5F53
    # 'GIFT jikkou-bi' (GIFT run date) : GIFT + U+5B9F U+884C U+65E5
    $L['GiftDate']   = 'GIFT' + [char]0x5B9F + [char]0x884C + [char]0x65E5
    # 'GIFT TIME' is ASCII in the sheet header.
    $L['GiftTime']   = 'GIFT TIME'
    # 'EXCEL' column (the W-form name Teams messages use).
    $L['Excel']      = 'EXCEL'
    # 'ken' (count unit) : U+4EF6
    $L['Ken']        = [string][char]0x4EF6
    # 'soushin-yotei' (planned send count) : U+9001 U+4FE1 U+4E88 U+5B9A
    $L['SendPlan']   = [char]0x9001 + [char]0x4FE1 + [char]0x4E88 + [char]0x5B9A
    # full-width colon U+FF1A (Teams text often carries it)
    $L['FwColon']    = [string][char]0xFF1A
    # output sheet header H: 'jobu' (same as Job) -- used as a sanity check
    $L['OutJob']     = $L['Job']
    return $L
}

# ---------------------------------------------------------------------------
# ConvertFrom-GiftMqListText
#   Parses the LIST page's Ctrl+A text. Each record is two lines:
#     No  SendNode RecvNode Correlid SendDate Tmode RecvDate Rtncd Rsncd
#     Msgid Reccnt FileSize
#   Fields are TAB separated on a real Ctrl+A capture; whitespace-only
#   separation (a hand-pasted copy) is accepted too. Returns records in
#   page order: No, SendNode, RecvNode, CorrelId, SendDate [datetime],
#   SendDateText, Tmode, RecvDate [datetime], Rtncd, Rsncd, MsgId,
#   RecCnt [int], FileSize [long]. Also NumRecords from the
#   'Number of records N' line (-1 when absent).
# ---------------------------------------------------------------------------
function ConvertFrom-GiftMqListText {
    param([string]$Text)

    $records = New-Object System.Collections.ArrayList
    $numRec  = -1
    if ([string]::IsNullOrWhiteSpace($Text)) {
        return @{ NumRecords = $numRec; Records = @($records.ToArray()) }
    }

    $m = [regex]::Match($Text, 'Number of records\s+(\d+)')
    if ($m.Success) { $numRec = [int]$m.Groups[1].Value }

    $rowRx  = [regex]'^(\d+)\s+(\S+)\s+(\S+)\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{1,2}:\d{2}:\d{2})\s+(\S+)\s+(\d{4}/\d{2}/\d{2}\s+\d{1,2}:\d{2}:\d{2})\s+(\d+)\s+(\d+)\s*$'
    $msgRx  = [regex]'^([A-Za-z0-9]{20,})\s+([\d,]+)\s+([\d,]+)\s*$'
    $inv    = [System.Globalization.CultureInfo]::InvariantCulture
    $lines  = $Text -split "`r?`n"
    $cur    = $null

    foreach ($raw in $lines) {
        $line = ($raw -replace "`t", ' ').Trim()
        if ($line.Length -eq 0) { continue }
        $rm = $rowRx.Match($line)
        if ($rm.Success) {
            $sendDt = $null; $recvDt = $null
            try { $sendDt = [datetime]::ParseExact(($rm.Groups[5].Value -replace '\s+', ' '), 'yyyy/MM/dd H:mm:ss', $inv) } catch {}
            try { $recvDt = [datetime]::ParseExact(($rm.Groups[7].Value -replace '\s+', ' '), 'yyyy/MM/dd H:mm:ss', $inv) } catch {}
            $cur = [PSCustomObject]@{
                No           = [int]$rm.Groups[1].Value
                SendNode     = $rm.Groups[2].Value
                RecvNode     = $rm.Groups[3].Value
                CorrelId     = $rm.Groups[4].Value
                SendDate     = $sendDt
                SendDateText = ($rm.Groups[5].Value -replace '\s+', ' ')
                Tmode        = $rm.Groups[6].Value
                RecvDate     = $recvDt
                Rtncd        = [int]$rm.Groups[8].Value
                Rsncd        = [int]$rm.Groups[9].Value
                MsgId        = ''
                RecCnt       = -1
                FileSize     = [long]-1
            }
            [void]$records.Add($cur)
            continue
        }
        if ($null -ne $cur -and $cur.MsgId -eq '') {
            $mm = $msgRx.Match($line)
            if ($mm.Success) {
                $cur.MsgId    = $mm.Groups[1].Value
                $cur.RecCnt   = [int]($mm.Groups[2].Value -replace ',', '')
                $cur.FileSize = [long]($mm.Groups[3].Value -replace ',', '')
            }
        }
    }
    return @{ NumRecords = $numRec; Records = @($records.ToArray()) }
}

# ---------------------------------------------------------------------------
# ConvertFrom-GiftMqDetailText
#   Parses a Detail page's Ctrl+A text ('KEY<TAB>value' per line; a key
#   with no value keeps ''). Keys are upper-case and may hold parentheses
#   (CORRELID(HEX)). Returns an ordered hashtable KEY -> value; header
#   lines without a key shape are ignored.
# ---------------------------------------------------------------------------
function ConvertFrom-GiftMqDetailText {
    param([string]$Text)

    $d = [ordered]@{}
    if ([string]::IsNullOrWhiteSpace($Text)) { return $d }
    $keyRx = [regex]'^([A-Z][A-Z0-9_()]*)(?:\t+|\s{2,}|\s)(.*)$'
    $bare  = [regex]'^([A-Z][A-Z0-9_()]*)\s*$'
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.TrimEnd()
        if ($line.Length -eq 0) { continue }
        $m = $keyRx.Match($line)
        if ($m.Success) {
            $d[$m.Groups[1].Value] = $m.Groups[2].Value.Trim()
            continue
        }
        $b = $bare.Match($line)
        if ($b.Success) { $d[$b.Groups[1].Value] = '' }
    }
    return $d
}

# ---------------------------------------------------------------------------
# ConvertTo-GiftMqDetailDateTime
#   '2026-09-18 06:54:03.91742' -> [datetime] (fraction kept to ms). Also
#   accepts 'yyyy/MM/dd HH:mm:ss'. $null when unparseable.
# ---------------------------------------------------------------------------
function ConvertTo-GiftMqDetailDateTime {
    param([string]$Value)
    if ([string]::IsNullOrWhiteSpace($Value)) { return $null }
    $s = $Value.Trim() -replace '/', '-'
    $m = [regex]::Match($s, '^(\d{4})-(\d{1,2})-(\d{1,2})[ T](\d{1,2}):(\d{2}):(\d{2})(?:\.(\d+))?')
    if (-not $m.Success) { return $null }
    try {
        $ms = 0
        if ($m.Groups[7].Success) {
            $frac = $m.Groups[7].Value
            if ($frac.Length -gt 3) { $frac = $frac.Substring(0, 3) }
            $ms = [int]($frac.PadRight(3, '0'))
        }
        return New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value), `
            ([int]$m.Groups[4].Value), ([int]$m.Groups[5].Value), ([int]$m.Groups[6].Value), $ms
    } catch { return $null }
}

# ---------------------------------------------------------------------------
# ConvertTo-GiftMqWholeSeconds
#   Truncates a [datetime] to the second (the sheet records whole seconds;
#   06:54:03.917 is written as 06:54:03, matching the HM page convention
#   the earlier rows were typed from). $null passes through.
# ---------------------------------------------------------------------------
function ConvertTo-GiftMqWholeSeconds {
    param($DateTime)
    if ($null -eq $DateTime) { return $null }
    $dt = [datetime]$DateTime
    return New-Object DateTime $dt.Year, $dt.Month, $dt.Day, $dt.Hour, $dt.Minute, $dt.Second
}

# ---------------------------------------------------------------------------
# Get-GiftMqDetailEndTime
#   The END time of a transfer = the Detail page's INSERTDATETIME (the
#   receive-log insertion stamp), truncated to the second. -Key overrides
#   the field. $null when the field is missing/unparseable.
# ---------------------------------------------------------------------------
function Get-GiftMqDetailEndTime {
    param($Detail, [string]$Key = 'INSERTDATETIME')
    if ($null -eq $Detail) { return $null }
    if (-not $Detail.Contains($Key)) { return $null }
    return ConvertTo-GiftMqWholeSeconds (ConvertTo-GiftMqDetailDateTime ([string]$Detail[$Key]))
}

# ---------------------------------------------------------------------------
# Test-GiftMqDetailMatchesRecord
#   The safety net behind every Detail click: the page reached must be
#   THIS record's. CORRELID(CHAR) must equal the record's Correlid and
#   SENDDATETIME must agree with the list's Send date to within
#   -ToleranceSeconds (default 1; the list shows whole seconds, the detail
#   carries milliseconds). Missing fields -> $false.
# ---------------------------------------------------------------------------
function Test-GiftMqDetailMatchesRecord {
    param($Detail, $Record, [int]$ToleranceSeconds = 1)
    if ($null -eq $Detail -or $null -eq $Record) { return $false }
    if (-not $Detail.Contains('CORRELID(CHAR)') -or -not $Detail.Contains('SENDDATETIME')) { return $false }
    if ([string]$Detail['CORRELID(CHAR)'] -ne [string]$Record.CorrelId) { return $false }
    $sd = ConvertTo-GiftMqDetailDateTime ([string]$Detail['SENDDATETIME'])
    if ($null -eq $sd -or $null -eq $Record.SendDate) { return $false }
    $delta = [Math]::Abs(((ConvertTo-GiftMqWholeSeconds $sd) - [datetime]$Record.SendDate).TotalSeconds)
    return ($delta -le $ToleranceSeconds)
}

# ---------------------------------------------------------------------------
# ConvertTo-GiftMqScheduledTime
#   Combines the mapping's GIFT run date + GIFT TIME cells into one
#   [datetime]. Either may arrive as text ('2026-09-17',
#   '10:14:59.9999999999984025', '10:30:00.000', '-') or as an Excel
#   serial (double) read via COM. The time is ROUNDED to the nearest
#   second: a time serial typed as 10:15 comes back as 10:14:59.99999.
#   Returns $null when the date is unusable; a usable date with no time
#   returns the date at 00:00 with .HasTime = $false via the -Detail
#   switch (the plain call returns just the [datetime] or $null).
# ---------------------------------------------------------------------------
function ConvertTo-GiftMqScheduledTime {
    param($Date, $Time, [switch]$Detail)

    $day = $null
    if ($Date -is [datetime]) {
        $day = ([datetime]$Date).Date
    } elseif ($Date -is [double] -or $Date -is [int] -or $Date -is [long] -or $Date -is [decimal]) {
        try { $day = [datetime]::FromOADate([double]$Date).Date } catch {}
    } elseif ($null -ne $Date) {
        $s = ([string]$Date).Trim()
        $m = [regex]::Match($s, '^(\d{4})[-/.](\d{1,2})[-/.](\d{1,2})')
        if ($m.Success) {
            try { $day = New-Object DateTime ([int]$m.Groups[1].Value), ([int]$m.Groups[2].Value), ([int]$m.Groups[3].Value) } catch {}
        }
    }
    if ($null -eq $day) { return $null }

    $secs = -1
    if ($Time -is [datetime]) {
        $secs = [int][Math]::Round(([datetime]$Time).TimeOfDay.TotalSeconds)
    } elseif ($Time -is [double] -or $Time -is [int] -or $Time -is [long] -or $Time -is [decimal]) {
        $frac = [double]$Time
        $frac = $frac - [Math]::Floor($frac)
        $secs = [int][Math]::Round($frac * 86400.0)
    } elseif ($null -ne $Time) {
        $t = ([string]$Time).Trim()
        $m = [regex]::Match($t, '^(\d{1,2}):(\d{2})(?::(\d{2})(?:\.(\d+))?)?')
        if ($m.Success) {
            $total = ([int]$m.Groups[1].Value) * 3600 + ([int]$m.Groups[2].Value) * 60
            if ($m.Groups[3].Success) { $total += [int]$m.Groups[3].Value }
            if ($m.Groups[4].Success) {
                $f = '0.' + $m.Groups[4].Value
                if ([double]$f -ge 0.5) { $total += 1 }
            }
            $secs = $total
        }
    }
    if ($secs -ge 86400) { $secs = $secs % 86400 }

    $hasTime = ($secs -ge 0)
    $dt = if ($hasTime) { $day.AddSeconds($secs) } else { $day }
    if ($Detail) { return @{ DateTime = $dt; HasTime = $hasTime } }
    return $dt
}

# ---------------------------------------------------------------------------
# Get-GiftMqMappingColumns
#   Locates the columns this feature reads in the operator's mapping
#   sheet from its header row (1-based indices; -1 when absent). Headers
#   are compared after stripping whitespace/newlines, by prefix, so
#   'GIFT<jikkou-bi>' with a trailing note still matches.
# ---------------------------------------------------------------------------
function Get-GiftMqMappingColumns {
    param([object[]]$Headers)
    $L = Get-GiftMqLabels
    $want = [ordered]@{
        Job      = $L.Job
        Excel    = $L.Excel
        Owner    = $L.Owner
        GiftDate = $L.GiftDate
        GiftTime = $L.GiftTime
    }
    $cols = @{}
    foreach ($k in $want.Keys) { $cols[$k] = -1 }
    if ($null -eq $Headers) { return $cols }
    for ($i = 0; $i -lt $Headers.Count; $i++) {
        $h = [string]$Headers[$i]
        if ([string]::IsNullOrWhiteSpace($h)) { continue }
        $norm = ($h -replace '\s+', '')
        foreach ($k in $want.Keys) {
            if ($cols[$k] -ne -1) { continue }
            $label = ($want[$k] -replace '\s+', '')
            if ($norm.StartsWith($label, [System.StringComparison]::OrdinalIgnoreCase)) {
                $cols[$k] = $i + 1
                break
            }
        }
    }
    return $cols
}

# ---------------------------------------------------------------------------
# ConvertTo-GiftMqSchedules
#   Turns raw mapping rows (hashtables/objects with Job, Excel, Owner,
#   GiftDate, GiftTime, Row) into schedule entries with a [datetime]
#   Scheduled + HasTime. Rows with no usable GIFT run date are dropped
#   (that column is what "the leader gave us this job" means). Optional
#   -Owner keeps one owner's rows; -Jobs keeps listed jobs; -FromDate /
#   -ToDate ([datetime] dates, inclusive) window the run date.
# ---------------------------------------------------------------------------
function ConvertTo-GiftMqSchedules {
    param(
        [object[]]$Rows,
        [string]$Owner = '',
        [string[]]$Jobs = @(),
        $FromDate = $null,
        $ToDate = $null
    )
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Rows) { return @($out.ToArray()) }
    $jobSet = @{}
    foreach ($j in $Jobs) { if (-not [string]::IsNullOrWhiteSpace($j)) { $jobSet[$j.Trim().ToUpperInvariant()] = $true } }

    foreach ($r in $Rows) {
        $job = [string](Get-GiftMqProp $r 'Job')
        if ([string]::IsNullOrWhiteSpace($job)) { continue }
        $job = $job.Trim().ToUpperInvariant()
        if ($jobSet.Count -gt 0 -and -not $jobSet.ContainsKey($job)) { continue }
        $own = [string](Get-GiftMqProp $r 'Owner')
        if (-not [string]::IsNullOrWhiteSpace($Owner) -and ($own.Trim() -ne $Owner.Trim())) { continue }
        $sch = ConvertTo-GiftMqScheduledTime -Date (Get-GiftMqProp $r 'GiftDate') -Time (Get-GiftMqProp $r 'GiftTime') -Detail
        if ($null -eq $sch) { continue }
        $dt = [datetime]$sch.DateTime
        if ($null -ne $FromDate -and $dt.Date -lt ([datetime]$FromDate).Date) { continue }
        if ($null -ne $ToDate   -and $dt.Date -gt ([datetime]$ToDate).Date)   { continue }
        [void]$out.Add([PSCustomObject]@{
            Job       = $job
            Excel     = [string](Get-GiftMqProp $r 'Excel')
            Owner     = $own.Trim()
            Scheduled = $dt
            HasTime   = [bool]$sch.HasTime
            Row       = Get-GiftMqProp $r 'Row'
        })
    }
    return @($out.ToArray())
}

function Get-GiftMqProp($obj, [string]$name) {
    if ($null -eq $obj) { return $null }
    if ($obj -is [System.Collections.IDictionary]) {
        if ($obj.Contains($name)) { return $obj[$name] }
        return $null
    }
    $p = $obj.PSObject.Properties[$name]
    if ($null -ne $p) { return $p.Value }
    return $null
}

# ---------------------------------------------------------------------------
# Resolve-GiftMqJobMatches
#   Assigns LIST-page records to scheduled jobs. Per day, schedules are
#   walked in scheduled order and each takes the nearest still-unassigned
#   record whose Send date lies within +-ToleranceMinutes of the schedule
#   and is not earlier than the record the previous schedule took (the
#   day's order rule). Records with a different Correlid than -CorrelId
#   (when given) are ignored. Result per schedule:
#     Job, Scheduled, HasTime, Owner, Excel, Row,
#     Record   (the chosen list record or $null),
#     Status   'ok' | 'ambiguous' (2+ candidates; nearest chosen) |
#              'notime' (date only; candidates listed, none chosen) |
#              'none' (nothing within tolerance),
#     Candidates (records considered, nearest first), DeltaSec.
#   Unassigned page records are returned separately (Unmatched) so the
#   driver can show "page rows with no job".
# ---------------------------------------------------------------------------
function Resolve-GiftMqJobMatches {
    param(
        [object[]]$Schedules,
        [object[]]$Records,
        [int]$ToleranceMinutes = 10,
        [string]$CorrelId = ''
    )
    $results = New-Object System.Collections.ArrayList
    $recs = @()
    if ($null -ne $Records) {
        $recs = @($Records | Where-Object { $null -ne $_.SendDate -and ([string]::IsNullOrWhiteSpace($CorrelId) -or [string]$_.CorrelId -eq $CorrelId) } | Sort-Object SendDate, No)
    }
    $assigned = @{}   # record No -> job
    $tol = [TimeSpan]::FromMinutes([Math]::Max(0, $ToleranceMinutes))

    $scheds = @()
    if ($null -ne $Schedules) { $scheds = @($Schedules | Sort-Object Scheduled, Job) }

    $lastTakenPerDay = @{}
    foreach ($s in $scheds) {
        $day = ([datetime]$s.Scheduled).Date
        $dayKey = $day.ToString('yyyy-MM-dd')
        $floor = $null
        if ($lastTakenPerDay.ContainsKey($dayKey)) { $floor = $lastTakenPerDay[$dayKey] }

        $cands = New-Object System.Collections.ArrayList
        foreach ($r in $recs) {
            if (([datetime]$r.SendDate).Date -ne $day) { continue }
            if ($assigned.ContainsKey([int]$r.No)) { continue }
            if (-not $s.HasTime) { [void]$cands.Add($r); continue }
            $delta = ([datetime]$r.SendDate) - ([datetime]$s.Scheduled)
            if ($delta.Duration() -gt $tol) { continue }
            if ($null -ne $floor -and ([datetime]$r.SendDate) -lt $floor) { continue }
            [void]$cands.Add($r)
        }

        $status = 'none'; $chosen = $null; $deltaSec = $null
        $ordered = @()
        if ($cands.Count -gt 0) {
            if ($s.HasTime) {
                $ordered = @($cands | Sort-Object @{ Expression = { ([datetime]$_.SendDate - [datetime]$s.Scheduled).Duration() } }, SendDate)
                $chosen = $ordered[0]
                $deltaSec = [int]([datetime]$chosen.SendDate - [datetime]$s.Scheduled).TotalSeconds
                $status = if ($cands.Count -ge 2) { 'ambiguous' } else { 'ok' }
                $assigned[[int]$chosen.No] = $s.Job
                $lastTakenPerDay[$dayKey] = [datetime]$chosen.SendDate
            } else {
                $ordered = @($cands | Sort-Object SendDate)
                $status = 'notime'
            }
        }
        [void]$results.Add([PSCustomObject]@{
            Job        = $s.Job
            Excel      = $s.Excel
            Owner      = $s.Owner
            Row        = $s.Row
            Scheduled  = $s.Scheduled
            HasTime    = $s.HasTime
            Record     = $chosen
            Status     = $status
            Candidates = $ordered
            DeltaSec   = $deltaSec
        })
    }
    $unmatched = @($recs | Where-Object { -not $assigned.ContainsKey([int]$_.No) })
    return @{ Matches = @($results.ToArray()); Unmatched = $unmatched }
}

# ---------------------------------------------------------------------------
# Get-GiftMqJobFromExcelName / Get-GiftMqExcelNameFromJob
#   The project's naming: the JOB is the EXCEL name with the 5th character
#   'W' replaced by 'J' (CJODWCP1 <-> CJODJCP1). Anything not shaped
#   ?XXXW??? is returned unchanged (upper-cased, trimmed).
# ---------------------------------------------------------------------------
function Get-GiftMqJobFromExcelName {
    param([string]$Name)
    if ([string]::IsNullOrWhiteSpace($Name)) { return '' }
    $n = $Name.Trim().ToUpperInvariant()
    if ($n.Length -ge 5 -and $n[4] -eq 'W') { return $n.Substring(0, 4) + 'J' + $n.Substring(5) }
    return $n
}

function Get-GiftMqExcelNameFromJob {
    param([string]$Job)
    if ([string]::IsNullOrWhiteSpace($Job)) { return '' }
    $n = $Job.Trim().ToUpperInvariant()
    if ($n.Length -ge 5 -and $n[4] -eq 'J') { return $n.Substring(0, 4) + 'W' + $n.Substring(5) }
    return $n
}

# ---------------------------------------------------------------------------
# ConvertFrom-GiftMqTeamsText
#   Pulls processed counts out of pasted Teams chat text. The pattern the
#   team writes is one line per event:
#     jobu:QJODWCP1 wo jisshi shimasu. (soushin-yotei:338 ken)
#   i.e. '<jobu>:<name>' and '(<soushin-yotei>:<n><ken>)' on the SAME line (a count
#   on a later line with no job name attaches to the last job named).
#   Both ASCII ':' and full-width colon are accepted; thousands commas
#   are stripped. Returns entries { ExcelName, Job (J-form), Count, Line }
#   in text order. Get-GiftMqTeamsCountMap folds them into Job -> Count
#   (LAST mention wins, so a corrected re-post overrides).
# ---------------------------------------------------------------------------
function ConvertFrom-GiftMqTeamsText {
    param([string]$Text)
    $out = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Text)) { return @($out.ToArray()) }
    $L = Get-GiftMqLabels
    $colon = '[:' + $L.FwColon + ']'
    $jobRx   = [regex]($L.Job + '\s*' + $colon + '\s*([A-Za-z0-9]{8})')
    $countRx = [regex]($L.SendPlan + '\s*' + $colon + '\s*([\d,]+)\s*' + $L.Ken)
    $lastJob = ''
    foreach ($raw in ($Text -split "`r?`n")) {
        $line = $raw.Trim()
        if ($line.Length -eq 0) { continue }
        $jm = $jobRx.Match($line)
        $name = ''
        if ($jm.Success) { $name = $jm.Groups[1].Value.ToUpperInvariant(); $lastJob = $name }
        $cm = $countRx.Match($line)
        if (-not $cm.Success) { continue }
        $who = if ($name -ne '') { $name } else { $lastJob }
        if ($who -eq '') { continue }
        [void]$out.Add([PSCustomObject]@{
            ExcelName = $who
            Job       = Get-GiftMqJobFromExcelName $who
            Count     = [int]($cm.Groups[1].Value -replace ',', '')
            Line      = $line
        })
    }
    return @($out.ToArray())
}

function Get-GiftMqTeamsCountMap {
    param([string]$Text)
    $map = @{}
    foreach ($e in @(ConvertFrom-GiftMqTeamsText -Text $Text)) { $map[$e.Job] = [int]$e.Count }
    return $map
}

# ---------------------------------------------------------------------------
# Format-GiftMqStamp / Format-GiftMqCount
#   The sheet's text conventions: 'yyyy/MM/dd HH:mm:ss' and '<n><ken>'
#   (thousands comma from 10,000 up, matching the hand-typed rows).
# ---------------------------------------------------------------------------
function Format-GiftMqStamp {
    param($DateTime)
    if ($null -eq $DateTime) { return '' }
    return ([datetime]$DateTime).ToString('yyyy/MM/dd HH:mm:ss', [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-GiftMqCount {
    param($Count)
    if ($null -eq $Count -or ([string]$Count).Trim() -eq '') { return '' }
    $n = 0
    if (-not [int]::TryParse((([string]$Count) -replace '[,\s]', ''), [ref]$n)) { return '' }
    $L = Get-GiftMqLabels
    $text = if ($n -ge 10000) { $n.ToString('#,0', [System.Globalization.CultureInfo]::InvariantCulture) } else { [string]$n }
    return $text + $L.Ken
}

# ---------------------------------------------------------------------------
# Get-GiftMqOutputPlan
#   Decides, per matched job, WHERE in the output sheet its GIFT row goes
#   and WHAT still needs writing. -SheetRows are the sheet's existing rows
#   as { Row, Side, Job, Start, End, Count } (Start/End/Count = $null or
#   '' when blank). The n-th match for a job pairs with the n-th existing
#   GIFT row for that job (a job may legitimately run twice a day, e.g.
#   two mapping rows). Result per match:
#     Row        existing GIFT row number, or 0 when it must be appended
#     Action     'update' | 'append'
#     NeedStart / NeedEnd / NeedCount   $true when the cell is blank or
#                                       -Force is set
#   Matches with no Record (Status none/notime) are skipped -- there is
#   nothing to write for them.
# ---------------------------------------------------------------------------
function Get-GiftMqOutputPlan {
    param([object[]]$SheetRows, [object[]]$Matches, [switch]$Force)
    $plan = New-Object System.Collections.ArrayList
    if ($null -eq $Matches) { return @($plan.ToArray()) }
    $rowsByJob = @{}
    if ($null -ne $SheetRows) {
        foreach ($sr in ($SheetRows | Sort-Object Row)) {
            if ([string]$sr.Side -ne 'GIFT') { continue }
            $j = ([string]$sr.Job).Trim().ToUpperInvariant()
            if ($j -eq '') { continue }
            if (-not $rowsByJob.ContainsKey($j)) { $rowsByJob[$j] = New-Object System.Collections.ArrayList }
            [void]$rowsByJob[$j].Add($sr)
        }
    }
    $seen = @{}
    foreach ($m in ($Matches | Sort-Object Scheduled, Job)) {
        if ($null -eq $m.Record) { continue }
        $j = ([string]$m.Job).Trim().ToUpperInvariant()
        $n = 0
        if ($seen.ContainsKey($j)) { $n = [int]$seen[$j] }
        $seen[$j] = $n + 1
        $existing = $null
        if ($rowsByJob.ContainsKey($j) -and $rowsByJob[$j].Count -gt $n) { $existing = $rowsByJob[$j][$n] }
        $blankStart = $true; $blankEnd = $true; $blankCount = $true
        $row = 0
        if ($null -ne $existing) {
            $row = [int]$existing.Row
            $blankStart = [string]::IsNullOrWhiteSpace([string]$existing.Start)
            $blankEnd   = [string]::IsNullOrWhiteSpace([string]$existing.End)
            $blankCount = [string]::IsNullOrWhiteSpace([string]$existing.Count)
        }
        [void]$plan.Add([PSCustomObject]@{
            Job       = $j
            Match     = $m
            Row       = $row
            Action    = if ($row -gt 0) { 'update' } else { 'append' }
            NeedStart = ($Force -or $blankStart)
            NeedEnd   = ($Force -or $blankEnd)
            NeedCount = ($Force -or $blankCount)
        })
    }
    return @($plan.ToArray())
}

# ---------------------------------------------------------------------------
# Get-GiftMqDetailTabCount
#   Tab presses from the page's focus anchor (the title line, reached via
#   Ctrl+F) to record No. N's Detail button: the Detail buttons are the
#   only focusable elements in the list, in page order, after
#   -TabsBeforeFirstDetail other controls (a 'Back' button etc.).
# ---------------------------------------------------------------------------
function Get-GiftMqDetailTabCount {
    param([int]$No, [int]$TabsBeforeFirstDetail = 0)
    if ($No -lt 1) { return 0 }
    return [Math]::Max(0, $TabsBeforeFirstDetail) + $No
}
