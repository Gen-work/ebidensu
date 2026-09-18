#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'modules/verify/GiftMqProcessTime.ps1')

Reset-Tests 'GiftMqProcessTime'

$t = "`t"
$L = Get-GiftMqLabels

# ---- fixtures: a slice of a real LIST page Ctrl+A capture (2026-09-18) ----
$listText = @"
Transfer status inquiry results
Number of records 51
No${t}Send node${t}Recv node${t}Correlid${t}Send date${t}Tmode${t}Recv date${t}Rtncd${t}Rsncd${t}
Msgid${t}Reccnt${t}File size
17${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/14 10:15:21${t}TXT${t}2026/09/14 10:15:21${t}0${t}0${t}
A2009999A000000000000001412026091401152157400001${t}5${t}497
18${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/14 10:30:22${t}TXT${t}2026/09/14 10:30:22${t}0${t}0${t}
A2009999A000000000000001412026091401302216800001${t}5${t}499
23${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/15 11:15:38${t}BIN${t}2026/09/15 11:15:38${t}0${t}0${t}
A2009999A000000000000001292026091502153816600001${t}0${t}1913
24${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/15 11:30:23${t}BIN${t}2026/09/15 11:30:23${t}0${t}0${t}
A2009999A000000000000001292026091502302343300001${t}0${t}606
25${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/15 11:45:29${t}TXT${t}2026/09/15 11:45:29${t}0${t}0${t}
A2009999A000000000000001292026091502452906200001${t}5${t}497
38${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/17 11:30:22${t}TXT${t}2026/09/17 11:30:22${t}0${t}0${t}
A2009999A000000000000001592026091702302190400001${t}28305${t}6213382
39${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/18 06:53:56${t}BIN${t}2026/09/18 06:53:56${t}0${t}0${t}
A2009999A000000000000000542026091721535646200001${t}0${t}82003
49${t}JHM076R${t}JHM102R${t}JJPCRS12${t}2026/09/18 11:00:29${t}TXT${t}2026/09/18 11:00:29${t}0${t}0${t}
A2009999A000000000000001572026091802002982500001${t}368${t}32616
"@

$detailText = @"
Transfer status inquiry results

INSERTDATETIME${t}2026-09-18 06:54:03.91742
LOGKIND${t}SENDLOG
SENDQMGRNAME${t}ATT1
SENDFILENAME${t}LJOD.L.VER.LJODWWBV.ZIP
SENDFILEPATH${t}
RECVFILENAME${t}JJPCRS1220260918065355544789
PROCPARAM1${t}-p 31789 /appl/JPC/shell/GIFTReceiverUnzip.sh JJPCRS1220260918065355544789
CORRELID(HEX)${t}000000000000000000000000000000000000000000000000
CORRELID(CHAR)${t}JJPCRS12
SENDNODE${t}JHM076R
SENDDATETIME${t}2026-09-18 06:53:56.491
RECVDATETIME${t}2026-09-18 06:53:56.521
MSGTOTALLENGTH${t}82003
RETURNCODE${t}0
REASONCODE${t}0
"@

# ---- ConvertFrom-GiftMqListText ----
$list = ConvertFrom-GiftMqListText -Text $listText
Assert-Equal 51 $list.NumRecords                       'list: Number of records read'
Assert-Equal 8  $list.Records.Count                    'list: eight two-line records parsed'
$r49 = $list.Records | Where-Object { $_.No -eq 49 }
Assert-Equal 'JJPCRS12' $r49.CorrelId                  'list: correl id'
Assert-Equal '2026/09/18 11:00:29' (Format-GiftMqStamp $r49.SendDate) 'list: send date parsed as datetime'
Assert-Equal 'TXT' $r49.Tmode                          'list: tmode'
Assert-Equal 368 $r49.RecCnt                           'list: second line Reccnt'
Assert-Equal 32616 $r49.FileSize                       'list: second line file size'
Assert-Equal 'A2009999A000000000000001572026091802002982500001' $r49.MsgId 'list: msgid'
$r38 = $list.Records | Where-Object { $_.No -eq 38 }
Assert-Equal 28305 $r38.RecCnt                         'list: large count'
Assert-Equal 0 (ConvertFrom-GiftMqListText -Text '').Records.Count 'list: empty text -> no records'
$spaceOnly = "1 JHM076R JHM102R JJPCRS12 2026/09/10 10:15:29 TXT 2026/09/10 10:15:29 0 0`nA2009999A000000000000000792026091001152922400001 5 552"
$sp = ConvertFrom-GiftMqListText -Text $spaceOnly
Assert-Equal 1 $sp.Records.Count                       'list: whitespace-separated copy also parses'
Assert-Equal 5 $sp.Records[0].RecCnt                   'list: whitespace-separated second line'
Assert-Equal -1 $sp.NumRecords                         'list: no header line -> NumRecords -1'
$oneDigitHour = "2 JHM076R JHM102R JJPCRS12 2026/09/10 9:05:01 TXT 2026/09/10 9:05:01 0 0"
Assert-Equal '2026/09/10 09:05:01' (Format-GiftMqStamp (ConvertFrom-GiftMqListText -Text $oneDigitHour).Records[0].SendDate) 'list: single-digit hour accepted'

# ---- ConvertFrom-GiftMqDetailText ----
$d = ConvertFrom-GiftMqDetailText -Text $detailText
Assert-Equal '2026-09-18 06:54:03.91742' $d['INSERTDATETIME'] 'detail: INSERTDATETIME'
Assert-Equal 'JJPCRS12' $d['CORRELID(CHAR)']            'detail: key with parentheses'
Assert-Equal '' $d['SENDFILEPATH']                     'detail: empty value kept as empty string'
Assert-True  $d.Contains('SENDFILEPATH')               'detail: empty-value key still present'
Assert-Equal '-p 31789 /appl/JPC/shell/GIFTReceiverUnzip.sh JJPCRS1220260918065355544789' $d['PROCPARAM1'] 'detail: value with spaces kept whole'
Assert-True  (-not $d.Contains('Transfer'))            'detail: title line ignored'

# ---- ConvertTo-GiftMqDetailDateTime / end time ----
$ins = ConvertTo-GiftMqDetailDateTime '2026-09-18 06:54:03.91742'
Assert-Equal 917 $ins.Millisecond                      'detail dt: fraction truncated to ms'
Assert-Equal '2026/09/18 06:54:03' (Format-GiftMqStamp (ConvertTo-GiftMqWholeSeconds $ins)) 'detail dt: whole seconds truncate (not round)'
Assert-Equal '2026/09/18 06:54:03' (Format-GiftMqStamp (Get-GiftMqDetailEndTime -Detail $d)) 'end time = INSERTDATETIME to the second'
Assert-True  ($null -eq (Get-GiftMqDetailEndTime -Detail $d -Key 'NOPE')) 'end time: missing key -> null'
Assert-True  ($null -eq (ConvertTo-GiftMqDetailDateTime 'garbage'))    'detail dt: garbage -> null'
Assert-Equal '2026/09/18 11:30:31' (Format-GiftMqStamp (ConvertTo-GiftMqDetailDateTime '2026/09/18 11:30:31')) 'detail dt: slash form accepted'

# ---- Test-GiftMqDetailMatchesRecord ----
$r39 = $list.Records | Where-Object { $_.No -eq 39 }
Assert-True  (Test-GiftMqDetailMatchesRecord -Detail $d -Record $r39)  'verify: detail belongs to record 39'
Assert-True  (-not (Test-GiftMqDetailMatchesRecord -Detail $d -Record $r49)) 'verify: detail rejected for another record'
$dOther = ConvertFrom-GiftMqDetailText -Text ($detailText -replace 'JJPCRS12', 'JJPCRS13')
Assert-True  (-not (Test-GiftMqDetailMatchesRecord -Detail $dOther -Record $r39)) 'verify: correl mismatch rejected'
Assert-True  (-not (Test-GiftMqDetailMatchesRecord -Detail @{} -Record $r39)) 'verify: missing fields rejected'

# ---- ConvertTo-GiftMqScheduledTime ----
Assert-Equal '2026/09/17 10:15:00' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date '2026-09-17' -Time '10:14:59.9999999999984025')) 'sched: float artefact rounds up to 10:15:00'
Assert-Equal '2026/09/15 11:30:00' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date '2026-09-15' -Time '11:30:00.0000000000015975')) 'sched: float artefact rounds down'
Assert-Equal '2026/08/25 13:20:27' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date '2026-08-25' -Time '13:20:26.99999999999675925')) 'sched: seconds kept'
Assert-Equal '2026/09/04 10:30:00' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date '2026/9/4' -Time '10:30')) 'sched: slash date + HH:mm'
$serialDate = (New-Object DateTime 2026, 9, 17).ToOADate()
$serialTime = (10.0 * 3600 + 15.0 * 60) / 86400.0
Assert-Equal '2026/09/17 10:15:00' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date $serialDate -Time $serialTime)) 'sched: Excel serial date + time fraction'
Assert-Equal '2026/09/17 10:15:00' (Format-GiftMqStamp (ConvertTo-GiftMqScheduledTime -Date (New-Object DateTime 2026, 9, 17) -Time $serialTime)) 'sched: [datetime] date'
$noTime = ConvertTo-GiftMqScheduledTime -Date '2026-08-20' -Time '-' -Detail
Assert-True  (-not $noTime.HasTime)                    'sched: dash time -> HasTime false'
Assert-Equal '2026/08/20 00:00:00' (Format-GiftMqStamp $noTime.DateTime) 'sched: dash time -> midnight'
Assert-True  ($null -eq (ConvertTo-GiftMqScheduledTime -Date '' -Time '10:00'))  'sched: blank date -> null'
Assert-True  ($null -eq (ConvertTo-GiftMqScheduledTime -Date 'log row 307' -Time '')) 'sched: non-date text -> null'

# ---- Get-GiftMqMappingColumns ----
$headers = @($L.Job, 'EXCEL', 'user', 'RCVFILE', 'SNDFILE', 'type', 'zip', 'stamp', 'lib', 'kind', 'ext',
    $L.Owner, ($L.GiftDate + "`n(x)"), 'GIFT TIME', 'GFIX' + [char]0x5B9F + [char]0x884C + [char]0x65E5, 'GFIX TIME')
$cols = Get-GiftMqMappingColumns -Headers $headers
Assert-Equal 1  $cols.Job       'columns: job = A'
Assert-Equal 2  $cols.Excel     'columns: EXCEL = B'
Assert-Equal 12 $cols.Owner     'columns: owner = L'
Assert-Equal 13 $cols.GiftDate  'columns: GIFT run date = M (newline in header ignored)'
Assert-Equal 14 $cols.GiftTime  'columns: GIFT TIME = N'
Assert-Equal -1 (Get-GiftMqMappingColumns -Headers @('a', 'b')).GiftDate 'columns: missing -> -1'

# ---- ConvertTo-GiftMqSchedules ----
$ni = [string][char]0x306B   # hiragana 'ni'
$ge = [string][char]0x3052   # hiragana 'ge'
$rows = @(
    @{ Job = 'JJDSJI02'; Excel = 'JJDSWI02'; Owner = $ni; GiftDate = '2026-09-17'; GiftTime = '10:14:59.9999999999984025'; Row = 7 },
    @{ Job = 'MJODJBA4'; Excel = 'MJODWBA4'; Owner = $ge; GiftDate = '2026-09-17'; GiftTime = '10:59:59.9999999999984025'; Row = 106 },
    @{ Job = 'FJODJWBV'; Excel = 'FJODWWBV'; Owner = $ge; GiftDate = '2026-09-15'; GiftTime = '11:30:00.0000000000015975'; Row = 61 },
    @{ Job = 'KJODJWBV'; Excel = 'KJODWWBV'; Owner = $ge; GiftDate = '2026-09-15'; GiftTime = '11:44:59.9999999999984025'; Row = 91 },
    @{ Job = 'JJDLJA09'; Excel = 'JJDLWA09'; Owner = $ni; GiftDate = '2026-08-20'; GiftTime = '-'; Row = 2 },
    @{ Job = 'CJODJCP1'; Excel = 'CJODWCP1'; Owner = '';  GiftDate = '';           GiftTime = ''; Row = 44 },
    @{ Job = 'KJODJDBC'; Excel = 'KJODWDBC'; Owner = $ni; GiftDate = 'log row 202 file stored'; GiftTime = ''; Row = 87 }
)
$sch = @(ConvertTo-GiftMqSchedules -Rows $rows)
Assert-Equal 5 $sch.Count                              'schedules: rows without a run date dropped'
Assert-Equal 4 @(ConvertTo-GiftMqSchedules -Rows $rows -FromDate (Get-Date '2026-09-01')).Count 'schedules: FromDate window'
Assert-Equal 3 @(ConvertTo-GiftMqSchedules -Rows $rows -Owner $ge).Count 'schedules: owner filter'
Assert-Equal 1 @(ConvertTo-GiftMqSchedules -Rows $rows -Jobs @('mjodjba4')).Count 'schedules: job filter is case-insensitive'
$ja09 = $sch | Where-Object { $_.Job -eq 'JJDLJA09' }
Assert-True (-not $ja09.HasTime)                       'schedules: dash time -> HasTime false'

# ---- Resolve-GiftMqJobMatches ----
$res = Resolve-GiftMqJobMatches -Schedules $sch -Records $list.Records -ToleranceMinutes 10
$byJob = @{}
foreach ($m in $res.Matches) { $byJob[$m.Job] = $m }
Assert-Equal 24 $byJob['FJODJWBV'].Record.No           'match: 09/15 11:30 -> record 24'
Assert-Equal 'ok' $byJob['FJODJWBV'].Status            'match: single candidate is ok'
Assert-Equal 23 $byJob['FJODJWBV'].DeltaSec            'match: delta seconds reported'
Assert-Equal 25 $byJob['KJODJWBV'].Record.No           'match: 09/15 11:45 -> record 25'
Assert-Equal 'none' $byJob['MJODJBA4'].Status          'match: 09/17 11:00 has no record in this slice'
Assert-Equal 'none' $byJob['JJDSJI02'].Status          'match: 09/17 10:15 has no record in this slice'
Assert-Equal 'none' $byJob['JJDLJA09'].Status          'match: date-only schedule with no records that day -> none'
Assert-Equal 6 $res.Unmatched.Count                    'match: 8 records minus the 2 taken -> 6 unassigned'
$unNos = @($res.Unmatched | ForEach-Object { $_.No } | Sort-Object)
Assert-Equal '17,18,23,38,39,49' ($unNos -join ',')    'match: exact unmatched record numbers'

# order rule: two schedules 15 min apart, records 2 min late each -> first
# takes the earlier record, second cannot fall back onto it.
$two = @(
    [PSCustomObject]@{ Job = 'A'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-15 11:30:00'); HasTime = $true; Row = 1 },
    [PSCustomObject]@{ Job = 'B'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-15 11:45:00'); HasTime = $true; Row = 2 }
)
$res2 = Resolve-GiftMqJobMatches -Schedules $two -Records $list.Records -ToleranceMinutes 10
Assert-Equal 24 ($res2.Matches | Where-Object { $_.Job -eq 'A' }).Record.No 'order: A -> 24'
Assert-Equal 25 ($res2.Matches | Where-Object { $_.Job -eq 'B' }).Record.No 'order: B -> 25'

# wide tolerance: 11:30 schedule now sees 11:15:38 (14 min early), 11:30:23 and
# 11:45:29 (15 min late) -> three candidates, nearest wins, flagged ambiguous.
$one = @([PSCustomObject]@{ Job = 'A'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-15 11:30:00'); HasTime = $true; Row = 1 })
$res3 = Resolve-GiftMqJobMatches -Schedules $one -Records $list.Records -ToleranceMinutes 20
Assert-Equal 'ambiguous' $res3.Matches[0].Status       'ambiguous: several candidates flagged'
Assert-Equal 24 $res3.Matches[0].Record.No             'ambiguous: nearest still chosen'
Assert-Equal 3 $res3.Matches[0].Candidates.Count       'ambiguous: all candidates listed'
Assert-Equal 24 $res3.Matches[0].Candidates[0].No      'ambiguous: candidates nearest first'

# same-time duplicate schedules (a job with two mapping rows) take two records
$dup = @(
    [PSCustomObject]@{ Job = 'X'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-14 10:15:00'); HasTime = $true; Row = 1 },
    [PSCustomObject]@{ Job = 'X'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-14 10:15:00'); HasTime = $true; Row = 2 }
)
$res4 = Resolve-GiftMqJobMatches -Schedules $dup -Records $list.Records -ToleranceMinutes 20
Assert-Equal '17,18' (@($res4.Matches | ForEach-Object { $_.Record.No }) -join ',') 'duplicate job rows: consecutive records, no double assignment'

# date-only schedule lists that day's records as candidates but takes none
$dateOnly = @([PSCustomObject]@{ Job = 'D'; Excel = ''; Owner = ''; Scheduled = (Get-Date '2026-09-14'); HasTime = $false; Row = 1 })
$res5 = Resolve-GiftMqJobMatches -Schedules $dateOnly -Records $list.Records
Assert-Equal 'notime' $res5.Matches[0].Status          'notime: status'
Assert-Equal 2 $res5.Matches[0].Candidates.Count       'notime: that day''s records offered'
Assert-True ($null -eq $res5.Matches[0].Record)        'notime: nothing chosen automatically'

# correl filter
$res6 = Resolve-GiftMqJobMatches -Schedules $one -Records $list.Records -ToleranceMinutes 10 -CorrelId 'JJPCRS13'
Assert-Equal 'none' $res6.Matches[0].Status            'correl filter: other correl records ignored'

# ---- job <-> excel name ----
Assert-Equal 'CJODJCP1' (Get-GiftMqJobFromExcelName 'CJODWCP1') 'name: W -> J'
Assert-Equal 'CJODWCP1' (Get-GiftMqExcelNameFromJob 'cjodjcp1') 'name: J -> W (upper-cased)'
Assert-Equal 'JJRVJZ20' (Get-GiftMqJobFromExcelName 'JJRVJZ20') 'name: already J-form unchanged'
Assert-Equal '' (Get-GiftMqJobFromExcelName '')       'name: blank -> blank'

# ---- Teams text ----
$teams = @(
    'Keiichi Ishikawa (GE)',
    ($L.Job + ':MJODWCP1' + [char]0x304C + 'OK'),
    ($L.Job + ':QJODWCP1' + [char]0x3092 + ' (' + $L.SendPlan + ':338' + $L.Ken + ')'),
    ($L.Job + $L.FwColon + 'LJODWCP1' + ' (' + $L.SendPlan + $L.FwColon + '1,015' + $L.Ken + ')'),
    ($L.Job + ':CJODWCP1'),
    ('(' + $L.SendPlan + ':12' + $L.Ken + ')'),
    ($L.Job + ':QJODWCP1' + ' (' + $L.SendPlan + ':340' + $L.Ken + ')')
) -join "`n"
$entries = @(ConvertFrom-GiftMqTeamsText -Text $teams)
Assert-Equal 4 $entries.Count                          'teams: four count lines'
Assert-Equal 'QJODJCP1' $entries[0].Job                'teams: W name mapped to J job'
Assert-Equal 338 $entries[0].Count                     'teams: count'
Assert-Equal 1015 $entries[1].Count                    'teams: full-width colon + thousands comma'
Assert-Equal 'CJODJCP1' $entries[2].Job                'teams: count on the next line attaches to last job'
$cmap = Get-GiftMqTeamsCountMap -Text $teams
Assert-Equal 340 $cmap['QJODJCP1']                     'teams map: last mention wins'
Assert-Equal 3 $cmap.Count                             'teams map: distinct jobs'
Assert-Equal 0 @(ConvertFrom-GiftMqTeamsText -Text '').Count 'teams: empty text'

# ---- formatting ----
Assert-Equal ('338' + $L.Ken) (Format-GiftMqCount 338)          'count: n + ken'
Assert-Equal ('12,518' + $L.Ken) (Format-GiftMqCount '12518')   'count: thousands comma from 10,000'
Assert-Equal ('4836' + $L.Ken) (Format-GiftMqCount 4836)        'count: no comma under 10,000 (matches hand-typed rows)'
Assert-Equal '' (Format-GiftMqCount '')                          'count: blank stays blank'
Assert-Equal '' (Format-GiftMqCount 'abc')                       'count: junk -> blank'
Assert-Equal '' (Format-GiftMqStamp $null)                       'stamp: null -> blank'

# ---- output plan ----
$sheetRows = @(
    @{ Row = 158; Side = 'GIFT'; Job = 'LJODJCP1'; Start = '2026/09/18 11:00:30'; End = '2026/09/18 11:00:33'; Count = '381' + $L.Ken },
    @{ Row = 159; Side = 'GFIX'; Job = 'LJODJCP1'; Start = ''; End = ''; Count = '' },
    @{ Row = 164; Side = 'GIFT'; Job = 'FJODJWBV'; Start = '2026/09/15 11:30:23'; End = ''; Count = '' },
    @{ Row = 165; Side = 'GFIX'; Job = 'FJODJWBV'; Start = ''; End = ''; Count = '' },
    @{ Row = 40;  Side = 'GIFT'; Job = 'JJDSJM51'; Start = 'x'; End = 'y'; Count = 'z' },
    @{ Row = 42;  Side = 'GIFT'; Job = 'JJDSJM51'; Start = ''; End = ''; Count = '' }
)
$mFJ = [PSCustomObject]@{ Job = 'FJODJWBV'; Scheduled = (Get-Date '2026-09-15 11:30'); Record = $r49; Status = 'ok' }
$mKJ = [PSCustomObject]@{ Job = 'KJODJWBV'; Scheduled = (Get-Date '2026-09-15 11:45'); Record = $r49; Status = 'ok' }
$mNo = [PSCustomObject]@{ Job = 'MJODJBA4'; Scheduled = (Get-Date '2026-09-17 11:00'); Record = $null; Status = 'none' }
$m51a = [PSCustomObject]@{ Job = 'JJDSJM51'; Scheduled = (Get-Date '2026-09-03 10:30'); Record = $r49; Status = 'ok' }
$m51b = [PSCustomObject]@{ Job = 'JJDSJM51'; Scheduled = (Get-Date '2026-09-03 10:30'); Record = $r49; Status = 'ok' }
$plan = @(Get-GiftMqOutputPlan -SheetRows $sheetRows -Matches @($mFJ, $mKJ, $mNo, $m51a, $m51b))
Assert-Equal 4 $plan.Count                             'plan: unmatched job produces no entry'
$pFJ = $plan | Where-Object { $_.Job -eq 'FJODJWBV' }
Assert-Equal 'update' $pFJ.Action                      'plan: existing GIFT row -> update'
Assert-Equal 164 $pFJ.Row                              'plan: existing row number'
Assert-True  (-not $pFJ.NeedStart)                     'plan: filled start not rewritten'
Assert-True  $pFJ.NeedEnd                              'plan: blank end needs writing'
Assert-True  $pFJ.NeedCount                            'plan: blank count needs writing'
$pKJ = $plan | Where-Object { $_.Job -eq 'KJODJWBV' }
Assert-Equal 'append' $pKJ.Action                      'plan: unknown job -> append'
Assert-Equal 0 $pKJ.Row                                'plan: append has no row yet'
Assert-True  ($pKJ.NeedStart -and $pKJ.NeedEnd -and $pKJ.NeedCount) 'plan: append needs everything'
$p51 = @($plan | Where-Object { $_.Job -eq 'JJDSJM51' })
Assert-Equal '40,42' (@($p51 | ForEach-Object { $_.Row }) -join ',') 'plan: n-th match pairs with n-th existing row'
Assert-True  (-not $p51[0].NeedStart)                  'plan: first duplicate row already filled'
Assert-True  $p51[1].NeedStart                         'plan: second duplicate row blank'
$forced = @(Get-GiftMqOutputPlan -SheetRows $sheetRows -Matches @($mFJ) -Force)
Assert-True  $forced[0].NeedStart                      'plan: -Force rewrites filled cells'

# ---- tab count ----
Assert-Equal 24 (Get-GiftMqDetailTabCount -No 24)                          'tabs: No 24 with no leading controls'
Assert-Equal 25 (Get-GiftMqDetailTabCount -No 24 -TabsBeforeFirstDetail 1) 'tabs: one leading control'
Assert-Equal 0  (Get-GiftMqDetailTabCount -No 0)                           'tabs: invalid No -> 0'

exit (Complete-Tests)
