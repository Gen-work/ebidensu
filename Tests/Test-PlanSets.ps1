#Requires -Version 5.1
# Test-PlanSets.ps1 -- the gfix-recv plan's day sets (10/20): which WBS rows
# are today's / late / ahead / someone else's / done, what the operator's
# typed answer picks, and which WBS cells the deliver phase fills with the
# actual dates. The WBS rows below are made up but follow the 10/20 sheet:
# jobs done early (actual start before the planned start), DBC jobs late
# since 10/15-10/19, a two-file job with one row done, a multi-day row.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Profile.ps1')
. (Join-Path $repoRoot 'kernel/Key.ps1')
. (Join-Path $repoRoot 'kernel/Worklist.ps1')

Reset-Tests 'PlanSets'

$prof = (Read-EbiProfile -Dir (Join-Path $repoRoot 'profiles/gfix-recv')).value
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
$log = New-Object PSObject
$log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
$log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) [void]$this.Lines.Add($m) }
$log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
$ctx = @{ WorkDir = [System.IO.Path]::GetTempPath(); RunId = 'test'; Profile = $prof; Log = $log; DryRun = $false; Session = @{}; Item = $null; KeyColumns = @('Excel_NAME') }

$loaded = @{}
foreach ($u in @('table.day_sets', 'human.pick', 'excel.mark_rows', 'verify.filter_records', 'verify.time_candidates', 'verify.parse_text')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    if (-not $r['ok']) { throw ('cannot load ' + $u + ': ' + $r['message']) }
    $loaded[$u] = $r['Entry']
}
function Invoke-TestStep {
    param([string]$Use, [hashtable]$With)
    $e = $loaded[$Use]
    $res = Resolve-EbiStepInputs -Manifest $e['Manifest'] -With $With -Session $ctx['Session']
    if (-not $res['ok']) { throw ($Use + ' inputs: ' + $res['message']) }
    return (& $e['Invoke'] $res['In'] $ctx)
}
function Get-Set { param($Result, [string]$Id) foreach ($s in @($Result['sets'])) { if ([string]$s['id'] -eq $Id) { return $s } }; return $null }
function Get-Names { param($Set) if ($null -eq $Set) { return '' }; return (@($Set['names']) -join ',') }

# ---------------------------------------------------------------- worklist helpers
Assert-Equal '2026-10-09' (ConvertTo-EbiWorklistDay -Text '2026/10/9') 'day: an older slash date reads'
Assert-Equal '2026-10-14' (ConvertTo-EbiWorklistDay -Text '2026-10-14 10:30:00') 'day: a time after the date is ignored'
Assert-Equal '' (ConvertTo-EbiWorklistDay -Text '-') 'day: "-" is no date'
$ds = Get-EbiWorklistDaySpec -Profile $prof
Assert-True ($null -ne $ds -and $ds['column'] -eq 'WORKDAY' -and $ds['fallback'] -eq 'GFIX_DATE') 'day: the gfix-recv worklist names WORKDAY, falling back to GFIX_DATE'
Assert-Equal 'hold' (Get-EbiWorklistRoleColumn -Profile $prof -Role 'hold') 'hold: the gfix-recv worklist has a hold column'
$rows = @( @{ Excel_NAME = 'A'; WORKDAY = '2026-10-20'; GFIX_DATE = '2026-10-16'; hold = '' }, @{ Excel_NAME = 'B'; WORKDAY = ''; GFIX_DATE = '2026/10/20'; hold = '' }, @{ Excel_NAME = 'C'; WORKDAY = '2026-10-16'; GFIX_DATE = ''; hold = 'HOST' } )
$sel = Select-EbiDayRows -Rows $rows -DaySpec $ds -Day '2026-10-20' -KeyColumns @('Excel_NAME')
Assert-Equal 'A,B' (@($sel['rows'] | ForEach-Object { $_['Excel_NAME'] }) -join ',') 'day: today''s rows kept, B by its fallback'
Assert-Equal 'C (10/16)' (@($sel['leftOut']) -join ',') 'day: the left-out row says its day'
$uh = Select-EbiUnheldRows -Rows $rows -Column 'hold' -KeyColumns @('Excel_NAME')
Assert-True (@($uh['rows']).Count -eq 2 -and @($uh['held'])[0] -eq 'C (HOST)') 'hold: the held row is left out with its reason'

# ---------------------------------------------------------------- day sets
$me = 'ME'; $pat = '^\s*ME\s*($|' + [string][char]0x2190 + ')|' + [string][char]0x2192 + '\s*ME\s*$'
function New-WbsRow { param([string]$Job, [string]$Owner, [string]$Start, [string]$End = '', [string]$Done = '') return @{ Excel_NAME = ($Job.Substring(0, 4) + 'W' + $Job.Substring(5)); JOB = $Job; owner = $Owner; start = $Start; end = $End; done = $Done } }
$wbs = @(
    (New-WbsRow 'CJODJSPQ' 'ME' '2026-10-20' '2026-10-20' '2026-10-16'),     # done early on 10/16
    (New-WbsRow 'KJODJDBC' 'ME' '2026-10-20' '2026-10-20' '2026-10-16'),     # "done" (typed by mistake)
    (New-WbsRow 'MJODJDBC' 'ME' '2026-10-19' '2026-10-19'),                  # late
    (New-WbsRow 'SJODJDBC' 'ME' '2026-10-15' '2026-10-15'),                  # late, on hold in the worklist
    (New-WbsRow 'RJODJDBC' 'ME' '2026-10-16' '2026-10-16'),                  # late, track ok here on 10/16
    (New-WbsRow 'QJODJSPN' 'ME' '2026-10-19' '2026-10-19'),                  # every verdict ok here, sheet not filled
    (New-WbsRow 'FJODJWBV' 'ME' '2026-10-21' '2026-10-21'),                  # ahead 10/21
    (New-WbsRow 'JJDSJI02' 'ME' '2026-10-22' '2026-10-22'),                  # ahead 10/22
    (New-WbsRow 'CJODJAF1' 'ME' '2026-10-28' '2026-10-28'),                  # too far ahead
    (New-WbsRow 'FJODJSPQ' 'OTHER' '2026-10-20' '2026-10-20'),               # someone else's today
    (New-WbsRow 'JJDSJPA1' 'ME' '2026-10-02' '2026-10-25'),                  # multi-day, today inside
    (New-WbsRow 'TJODJDUP' ('ME' + [string][char]0x2190 + 'OTHER') '2026-10-20' '2026-10-20' '2026-10-20'),   # two files: one row done
    (New-WbsRow 'TJODJDUP' ('ME' + [string][char]0x2190 + 'OTHER') '2026-10-20' '2026-10-20' ''),
    (New-WbsRow 'ZJODJOLD' 'ME' '2026-08-01' '2026-08-01'),                  # older than backDays
    (New-WbsRow 'YJODJPCK' 'ME' '2026-10-23' '2026-10-23'),                  # picked earlier today (worklist)
    (New-WbsRow 'NJODJNOD' 'ME' '' '')                                       # no planned date
)
$wl = @{ path = ''; columns = @('Excel_NAME', 'GFIX_DATE', 'WORKDAY', 'hold', 'track', 'logs', 'evidence', 'deliver'); rows = @(
    @{ Excel_NAME = 'SJODWDBC'; GFIX_DATE = '2026-10-15'; WORKDAY = '2026-10-15'; hold = 'HOST'; track = ''; logs = ''; evidence = ''; deliver = '' },
    @{ Excel_NAME = 'RJODWDBC'; GFIX_DATE = '2026/10/16'; WORKDAY = ''; hold = ''; track = 'ok'; logs = 'ok'; evidence = ''; deliver = '' },
    @{ Excel_NAME = 'QJODWSPN'; GFIX_DATE = '2026-10-19'; WORKDAY = '2026-10-19'; hold = ''; gift = ''; track = 'ok'; logs = 'ok'; evidence = 'ok'; deliver = 'ok' },
    @{ Excel_NAME = 'YJODWPCK'; GFIX_DATE = '2026-10-20'; WORKDAY = '2026-10-20'; hold = ''; track = ''; logs = ''; evidence = ''; deliver = '' },
    @{ Excel_NAME = 'MJDSWM40'; GFIX_DATE = '2026-10-20'; WORKDAY = '2026-10-20'; hold = ''; track = 'ok'; logs = ''; evidence = ''; deliver = '' }
) }
$ctx['Session']['wl'] = @{ kind = 'worklist'; value = $wl }
$mapping = @( @{ Excel_NAME = 'XJDSWM99'; GFIX_TIME = '13:00:00' }, @{ Excel_NAME = 'FJODWWBV'; GFIX_TIME = '10:30:00' } )
$r = Invoke-TestStep 'table.day_sets' @{ worklist = 'wl'; records = $wbs; today = '2026-10-20'; ownerPattern = $pat; extra = $mapping; carryField = 'track' }
Assert-True ($r['ok']) 'sets: ok'
Assert-Equal 'JJDSWPA1,TJODWDUP,YJODWPCK,MJDSWM40' (Get-Names (Get-Set $r 'today')) 'sets: today = the multi-day row, the two-file job with one row open, the rows picked earlier today (MJDSWM40 is in no WBS row)'
Assert-Equal 'picked earlier today; track ok' (Get-Set $r 'today')['notes']['MJDSWM40'] 'sets: ... with what is done'
Assert-Equal 'MJODWDBC,RJODWDBC' (Get-Names (Get-Set $r 'late')) 'sets: late = planned before today, not done (the held one is not here)'
Assert-Equal 'FJODWWBV' (Get-Names (Get-Set $r 'ahead:2026-10-21')) 'sets: ahead 10/21'
Assert-Equal 'JJDSWI02' (Get-Names (Get-Set $r 'ahead:2026-10-22')) 'sets: ahead 10/22 (2 working days)'
Assert-Equal 'FJODWSPQ' (Get-Names (Get-Set $r 'others')) 'sets: someone else''s today'
Assert-Equal 'XJDSWM99' (Get-Names (Get-Set $r 'mapping')) 'sets: in mapping.xlsx only'
$hold = Get-Set $r 'hold'
Assert-True ((Get-Names $hold) -eq 'SJODWDBC' -and [int]$hold['n'] -eq 0 -and $hold['notes']['SJODWDBC'] -eq 'HOST') 'sets: the held row is shown, not pickable, with its reason'
Assert-Equal 1 (Get-Set $r 'today')['n'] 'sets: today is choice 1'
Assert-True ((Get-Set $r 'late')['notes']['RJODWDBC'] -like 'since 10/16; track ok, logs ok*') ('sets: a late row says since when and what is done here: ' + (Get-Set $r 'late')['notes']['RJODWDBC'])
Assert-True ((Get-Set $r 'ahead:2026-10-21')['notes']['FJODWWBV'] -eq 'mapping 10:30:00') 'sets: a mapping.xlsx time is noted on a placed row'
Assert-True ((Get-Set $r 'today')['notes']['JJDSWPA1'] -eq 'until 10/25') 'sets: a multi-day row says until when'
$info = (@($r['info']) -join ' | ')
Assert-True ($info -like '*done by the WBS actual start (AJ), never in the worklist: CJODWSPQ 10/16 (planned 10/20), KJODWDBC 10/16 (planned 10/20)*') ('info: done early rows listed with the planned day: ' + $info)
Assert-True ($info -like '*done here (every phase ok): QJODWSPN 10/19*') 'info: done here, WBS not filled yet'
Assert-True ($info -like '*1 of your row(s) have no planned date*') 'info: rows without a date are counted'
Assert-True (-not ($info -like '*ZJODWOLD*') -and -not ((@($r['sets'] | ForEach-Object { @($_['names']) }) -join ',') -like '*ZJODWOLD*')) 'sets: a row older than backDays is nowhere'
$recs = @{}; foreach ($x in @($r['records'])) { $recs[[string]$x['Excel_NAME']] = $x }
Assert-Equal '2026-10-16' $recs['RJODWDBC']['GFIX_DATE'] 'records: a late row whose track is ok keeps its run date (logs / evidence continue it)'
Assert-Equal '2026-10-20' $recs['MJODWDBC']['GFIX_DATE'] 'records: any other picked row runs today'
Assert-Equal '2026-10-20' $recs['FJODWWBV']['WORKDAY'] 'records: every pickable record carries WORKDAY = today'
Assert-True (-not $recs.Contains('SJODWDBC')) 'records: the held row is not a record'
$pre = @($r['presets'])
Assert-Equal 3 $pre.Count 'presets: today, today + late, today + the next working day'
Assert-Equal 'JJDSWPA1,TJODWDUP,YJODWPCK,MJDSWM40,MJODWDBC,RJODWDBC' (@($pre[1]['names']) -join ',') 'presets: 2 = today + late'
Assert-Equal 'JJDSWPA1,TJODWDUP,YJODWPCK,MJDSWM40,FJODWWBV' (@($pre[2]['names']) -join ',') 'presets: 3 = today + ahead 10/21'
# AJ is typed by hand: not proof on its own (10/20 office run)
$wbsAj = @(
    (New-WbsRow 'MJODJSPQ' 'ME' '2026-10-21' '2026-10-21' '2026-10-20'),     # AJ typed this morning, job not run yet
    (New-WbsRow 'LJODJDBC' 'ME' '2026-10-20' '2026-10-20' '2026-10-16'),     # AJ typed by mistake, row open in the worklist
    (New-WbsRow 'KJODJSPP' 'ME' '2026-10-20' '2026-10-20' '2026-10-16'),     # done early: deliver ok in the worklist
    (New-WbsRow 'QJODJSPP' 'OTHER' '2026-10-20' '2026-10-20' '2026-10-20')   # someone else's, AJ today
)
$ctx['Session']['wl2'] = @{ kind = 'worklist'; value = @{ path = ''; columns = @('Excel_NAME', 'GFIX_DATE', 'WORKDAY', 'hold', 'gift', 'track', 'logs', 'evidence', 'deliver'); rows = @(
    @{ Excel_NAME = 'LJODWDBC'; GFIX_DATE = '2026-10-15'; WORKDAY = ''; hold = ''; gift = ''; track = ''; logs = ''; evidence = ''; deliver = '' },
    @{ Excel_NAME = 'KJODWSPP'; GFIX_DATE = '2026-10-20'; WORKDAY = '2026-10-16'; hold = ''; gift = ''; track = ''; logs = ''; evidence = ''; deliver = 'ok' }
) } }
$ra = Invoke-TestStep 'table.day_sets' @{ worklist = 'wl2'; records = $wbsAj; today = '2026-10-20'; ownerPattern = $pat; carryField = 'track' }
Assert-Equal 'LJODWDBC' (Get-Names (Get-Set $ra 'today')) 'AJ: a row the worklist has not finished is offered although the WBS has an AJ'
Assert-Equal 'WBS AJ 10/16' (Get-Set $ra 'today')['notes']['LJODWDBC'] 'AJ: ... with the AJ as a note'
Assert-Equal 'MJODWSPQ' (Get-Names (Get-Set $ra 'ahead:2026-10-21')) 'AJ: an AJ dated today on a job the worklist never had is offered (typed in the morning)'
Assert-Equal 'QJODWSPP' (Get-Names (Get-Set $ra 'others')) 'AJ: the same for someone else''s row'
$ia = (@($ra['info']) -join ' | ')
Assert-True ($ia -like '*done here (every phase ok): KJODWSPP*' -and $ia -like '*WBS AJ filled, not done here -- still offered: MJODWSPQ 10/20, LJODWDBC 10/16*') ('AJ: info lines: ' + $ia)
$rs = @($ra['sets']); $rp = @($ra['presets'])
$ctx['DryRun'] = $true
$p = Invoke-TestStep 'human.pick' @{ sets = $rs; presets = $rp; title = 'PLAN'; answer = @('MJODWSPQ', 'qjodwspp') }
Assert-True ($p['ok'] -and (@($p['picked']) -join ',') -eq 'MJODWSPQ,QJODWSPP') 'pick: -Only names (run.only) are the answer -- any case, from the ahead and others sets'
$ctx['DryRun'] = $false
$script:answers = New-Object System.Collections.Queue
foreach ($a in @('')) { $script:answers.Enqueue($a) }
$ctx['Reader'] = { if ($script:answers.Count -gt 0) { return [string]$script:answers.Dequeue() }; return 'q' }
$p = Invoke-TestStep 'human.pick' @{ sets = $rs; presets = $rp; title = 'PLAN'; answer = @('MJODWSPQ') }
Assert-True ($p['ok'] -and (@($p['picked']) -join ',') -eq 'MJODWSPQ' -and -not $p['auto']) 'pick: with a console the given names are read back and Enter confirms them'
$script:answers.Clear(); foreach ($a in @('1', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $rs; presets = $rp; title = 'PLAN'; answer = @('NOSUCH') }
Assert-True ($p['ok'] -and (@($p['picked']) -join ',') -eq 'LJODWDBC') 'pick: a given name no set offers falls back to the panel'
$ctx['Reader'] = $null

$r2 = Invoke-TestStep 'table.day_sets' @{ worklist = 'wl'; records = $wbs; today = '2026-10-20'; ownerPattern = ''; extra = $mapping }
Assert-True ($null -eq (Get-Set $r2 'others') -and (Get-Names (Get-Set $r2 'today')) -like '*FJODWSPQ*') 'sets: no owner pattern -> every row is yours, no others set'
Assert-True ([bool](Get-Set $r2 'mapping')['default']) 'sets: no owner pattern -> mapping.xlsx-only jobs are in the default too'
$bad = Invoke-TestStep 'table.day_sets' @{ worklist = 'wl'; records = $wbs; today = '10/20' }
Assert-Equal 'input_invalid' $bad['failure'] 'sets: today must be yyyy-MM-dd'

# ---------------------------------------------------------------- pick: what an answer means
$sets = @($r['sets']); $presets = @($r['presets'])
$x = HumanPick-Resolve -Text '' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 'JJDSWPA1,TJODWDUP,YJODWPCK,MJDSWM40' (@($x['keys']) -join ',') 'pick: Enter = choice 1'
$x = HumanPick-Resolve -Text '2' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 6 @($x['keys']).Count 'pick: 2 = today + late'
$x = HumanPick-Resolve -Text 'DBC' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 'MJODWDBC,RJODWDBC' (@($x['keys']) -join ',') 'pick: part of a name = every offered name holding it (the held one is not offered)'
$x = HumanPick-Resolve -Text ([string][char]0xFF24 + [char]0xFF22 + [char]0xFF23) -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 2 @($x['keys']).Count 'pick: full-width letters read the same'
$x = HumanPick-Resolve -Text '2 -RJODWDBC' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 'JJDSWPA1,TJODWDUP,YJODWPCK,MJDSWM40,MJODWDBC' (@($x['keys']) -join ',') 'pick: -NAME drops one'
$x = HumanPick-Resolve -Text '1, fjodwwbv' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 5 @($x['keys']).Count 'pick: a choice plus a name, any case, comma or space'
$x = HumanPick-Resolve -Text 'm' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 'more' $x['kind'] 'pick: m = more sets'
$x = HumanPick-Resolve -Text '3' -Sets $sets -Presets $presets -Level 'more'
Assert-Equal 'FJODWWBV' (@($x['keys']) -join ',') 'pick: on the second screen a number is a set (3 = ahead 10/21)'
$x = HumanPick-Resolve -Text '' -Sets $sets -Presets $presets -Level 'more'
Assert-True (-not $x['ok']) 'pick: Enter on the second screen is not an answer'
$x = HumanPick-Resolve -Text 'h DBC' -Sets $sets -Presets $presets -Level 'first'
Assert-True ($x['kind'] -eq 'hold' -and (@($x['keys']) -join ',') -eq 'MJODWDBC,RJODWDBC') 'pick: h DBC = put them on hold'
$x = HumanPick-Resolve -Text 'NOSUCH' -Sets $sets -Presets $presets -Level 'first'
Assert-True (-not $x['ok'] -and $x['message'] -like '*NOSUCH (not in any set)*') 'pick: an unknown name is said, nothing is picked'
$x = HumanPick-Resolve -Text '7' -Sets $sets -Presets $presets -Level 'first'
Assert-True (-not $x['ok']) 'pick: a number past the last choice that no name holds is refused'
$x = HumanPick-Resolve -Text '0' -Sets $sets -Presets $presets -Level 'first'
Assert-Equal 'none' $x['kind'] 'pick: 0 = nothing today'
$lines = @(HumanPick-Lines -Sets $sets -Presets $presets -Level 'first')
Assert-True ($lines[0] -like '1  today 10/20 (recommended)  <- Enter' -and $lines.Count -eq 6) ('pick: the first screen shows 3 choices, 1 marked Enter: ' + $lines[0])
$lines = @(HumanPick-Lines -Sets $sets -Presets $presets -Level 'more')
Assert-True (($lines -join '|') -like '*-  on hold  (1)  not pickable*') 'pick: the second screen shows the held row as not pickable'

# ---------------------------------------------------------------- pick: the conversation
$script:answers = New-Object System.Collections.Queue
$ctx['Reader'] = { if ($script:answers.Count -gt 0) { return [string]$script:answers.Dequeue() }; return 'q' }
foreach ($a in @('2', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-True ($p['ok'] -and @($p['picked']).Count -eq 6 -and -not $p['auto']) 'pick: "2" then Enter confirms today + late'
$script:answers.Clear(); foreach ($a in @('h MJOD', '', '', '', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN'; holdReason = 'HOST' }
Assert-True ($p['ok'] -and (@($p['hold']) -join ',') -eq 'MJODWDBC' -and $p['holdReason'] -eq 'HOST' -and @($p['picked']).Count -eq 4) 'pick: h MJOD, Enter (reason), Enter (confirm), then Enter = 1: one held, today picked'
$script:answers.Clear(); foreach ($a in @('2', 'DBC', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-Equal 'MJODWDBC,RJODWDBC' (@($p['picked']) -join ',') 'pick: an answer typed at the read-back replaces the first one'
$script:answers.Clear(); foreach ($a in @('MJODWDBC', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-True ($p['picked'] -is [System.Collections.IList] -and @($p['picked']).Count -eq 1 -and (ConvertTo-EbiInputValue -Type 'list' -Value $p['picked'])['ok']) ('pick: ONE name picked is still a list (10/20 gift: file.download names got a string): ' + (Get-EbiValueTypeName $p['picked']))
$script:answers.Clear(); foreach ($a in @('0', '')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-True ($p['ok'] -and $p['picked'] -is [System.Collections.IList] -and @($p['picked']).Count -eq 0 -and -not $p['any']) 'pick: 0 = nothing today is an empty list, not null'
$script:answers.Clear(); foreach ($a in @('m', 'b', 'q')) { $script:answers.Enqueue($a) }
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-Equal 'operator_quit' $p['failure'] 'pick: m, b, q -> stopped'
$ctx['Reader'] = $null
$ctx['DryRun'] = $true
$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN' }
Assert-True ($p['ok'] -and $p['auto'] -and @($p['picked']).Count -eq 4) 'pick: nobody to ask -> choice 1'
$ctx['DryRun'] = $false

$p = Invoke-TestStep 'human.pick' @{ sets = $sets; presets = $presets; title = 'PLAN'; autoPick = $true }
Assert-True ($p['ok'] -and $p['auto'] -and $p['pickedCount'] -eq 4 -and $p['any']) 'pick: autoPick takes choice 1 without a panel; pickedCount / any say how many'

# ---------------------------------------------------------------- GIFT files on the Jenkins list (gfixRecv.gift)
$tab = [string][char]9
$jkText = (@(
    ('F202609110030.csv' + $tab + '2026/09/11 10:30:29' + $tab + '12 KB' + $tab + 'view'),
    ('F202609110031.csv' + $tab + '2026/09/11 10:44:10' + $tab + '3 KB' + $tab + 'view'),
    ('F202609110025.csv' + $tab + '2026/09/11 9:58:00' + $tab + '1 KB' + $tab + 'view'),
    ('F202609110026.csv' + $tab + '2026/09/11 9:59:30' + $tab + '1 KB' + $tab + 'view'),
    ('F202609100001.csv' + $tab + '2026/09/10 10:31:00' + $tab + '1 KB' + $tab + 'view'),
    ('notes.txt' + $tab + '2026/09/11 10:31:00' + $tab + '1 KB' + $tab + 'view')
) -join "`r`n")
$jr = Invoke-TestStep 'verify.parse_text' @{ text = $jkText; grammar = $prof['grammar']['jenkinsReport'] }
Assert-Equal 6 $jr['recordCount'] 'gift: the sample Jenkins lines parse with the real grammar'
$pat = [string]$prof['pages']['jenkinsReport']['fileNamePattern']
$c = VerifyTimeCandidates-Find -Records $jr['records'] -Date '2026/9/11' -Clock '10:30' -NamePattern $pat -Label 'GIFT run'
Assert-True ($c['ok'] -and $c['certain'] -and (@($c['picked']) -join ',') -eq 'F202609110030.csv') ('gift: one file in 10:28..10:43 and none expected otherwise -> certain: ' + $c['reason'])
Assert-Equal 4 $c['dayCount'] 'gift: that day''s files matching the name pattern (other days and notes.txt left out)'
$c = VerifyTimeCandidates-Find -Records $jr['records'] -Date '2026-09-11' -Clock '0.4375' -NamePattern $pat
Assert-True ($c['certain']) 'gift: an Excel day fraction (0.4375 = 10:30) is a time'
$c = VerifyTimeCandidates-Find -Records $jr['records'] -Date '2026/9/11' -Clock '10:30' -NamePattern $pat -Expect 2 -Label 'GIFT run'
Assert-True (-not $c['certain'] -and @($c['picked']).Count -eq 0) 'gift: 2 expected, 1 in the window -> not certain, nothing taken'
Assert-Equal 'window,hour:09,hour:10' (@($c['sets'] | ForEach-Object { $_['id'] }) -join ',') 'gift: sets = the window, then the rest of that day by hour, earliest first'
Assert-Equal 'F202609110025.csv,F202609110026.csv' (@($c['sets'][1]['names']) -join ',') 'gift: the 09 hour holds both 9:5x files'
Assert-Equal '09:58:00' $c['sets'][1]['notes']['F202609110025.csv'] 'gift: each name notes its time'
Assert-Equal 'F202609110031.csv' (@($c['sets'][2]['names']) -join ',') 'gift: the 10 hour keeps only what is outside the window'
Assert-True ($c['presets'][0]['label'] -like 'GIFT run 09/11 10:30:00: 10:28-10:43 (expect 2)*') ('gift: preset 1 is the window: ' + $c['presets'][0]['label'])
$c = VerifyTimeCandidates-Find -Records $jr['records'] -Date '2026/9/11' -Clock '-' -NamePattern $pat
Assert-True (-not $c['certain'] -and @($c['window']).Count -eq 0 -and (@($c['info']) -join ' ') -like '*"-" is not a time*') 'gift: GIFT TIME "-" -> no window, pick by hour'
$c = VerifyTimeCandidates-Find -Records $jr['records'] -Date '' -Clock '10:30' -NamePattern $pat
Assert-Equal 'no_date' $c['failure'] 'gift: no GIFT date -> no_date'
$daySets = @((VerifyTimeCandidates-Find -Records $jr['records'] -Date '2026/9/11' -Clock '' -NamePattern $pat)['sets'])
$x = HumanPick-Resolve -Text '0026' -Sets $daySets -Presets @() -Level 'first'
Assert-Equal 'F202609110026.csv' (@($x['keys']) -join ',') 'gift: a file is picked by part of its name'

$f = Invoke-TestStep 'verify.filter_records' @{ records = @($r['records']); where = @(@{ field = 'Excel_NAME'; op = 'in'; value = @() }); emptyOk = $true }
Assert-True ($f['ok'] -and $f['matched'] -eq 0) 'filter_records emptyOk: picking nothing is an answer, not not_found'

# ---------------------------------------------------------------- deliver: WBS actual dates
$recv = 'R'; $task = 'T'
$grid = @{
    A  = @('KJODJSPP', 'KJODJSPP', 'RJODJDB6', 'RJODJDB6', 'MJODJWBV')
    B  = @('S', $recv, $recv, $recv, $recv)
    G  = @($task, $task, 'before', $task, $task)
    AJ = @($null, $null, $null, [double]46311, $null)        # 46311 = 2026-10-16
    AK = @($null, $null, $null, $null, $null)
}
$items = @( @{ JOB = 'KJODJSPP'; GFIX_DATE = '2026/10/16' }, @{ JOB = 'RJODJDB6'; GFIX_DATE = '2026-10-15' }, @{ JOB = 'NOROW'; GFIX_DATE = '2026-10-20' } )
$plan = ExcelMarkRows-Plan -Grid $grid -FirstRow 7 -KeyColumn 'A' -Items $items -KeyField 'JOB' -Where @{ B = $recv; G = $task } -SetFields @{ AJ = 'GFIX_DATE' } -SetValues @{ AK = '2026-10-20' } -DateColumns @('AJ', 'AK')
$w = @($plan['writes'] | ForEach-Object { $_['letter'] + $_['row'] + '=' + $_['value'] }) -join ','
Assert-Equal 'AJ8=2026-10-16,AK8=2026-10-20,AK10=2026-10-20' $w 'mark_rows: the receive / task row only (not the send row, not another task); a blank AJ gets the run date, AK today'
Assert-Equal 'RJODJDB6 row 10: AJ already 2026-10-16' (@($plan['kept']) -join ',') 'mark_rows: a date typed by hand is kept (and said)'
Assert-Equal 'NOROW' (@($plan['notFound']) -join ',') 'mark_rows: a job with no row is reported'
$plan = ExcelMarkRows-Plan -Grid $grid -FirstRow 7 -KeyColumn 'A' -Items $items -KeyField 'JOB' -Where @{ B = $recv; G = $task } -SetFields @{ AJ = 'GFIX_DATE' } -SetValues @{} -DateColumns @('AJ') -OnlyBlank $false
Assert-True ((@($plan['writes'] | ForEach-Object { $_['letter'] + $_['row'] }) -join ',') -eq 'AJ8,AJ10') 'mark_rows: onlyBlank false overwrites'
# 10/20 second day: both task rows (after + summary) of the operator's own rows, one asked date, hours as a number
$own = 'OWN'; $oth = 'OTHER'
$g2 = @{
    A  = @('LJODJSPQ', 'LJODJSPQ', 'LJODJSPQ', 'QJODJSPP', 'QJODJSPP')
    B  = @($recv, $recv, $recv, $recv, $recv)
    G  = @('before', 'after', 'summary', 'after', 'summary')
    P  = @($own, ($own + '<-' + $oth), $own, $oth, ($oth + '<-' + $own))
    AJ = @($null, [double]46315, $null, $null, $null)      # 46315 = 2026-10-20, typed by hand in the morning
    AK = @($null, $null, $null, $null, $null)
    AQ = @($null, $null, $null, $null, $null)
}
$it2 = @( @{ JOB = 'LJODJSPQ' }, @{ JOB = 'QJODJSPP' } )
$ownPat = '^\s*' + $own + '\s*($|<-)'
$plan = ExcelMarkRows-Plan -Grid $g2 -FirstRow 7 -KeyColumn 'A' -Items $it2 -KeyField 'JOB' -Where @{ B = $recv; G = @('after', 'summary') } -WhereMatch @{ P = $ownPat } -SetFields @{} -SetValues @{ AJ = '10/21'; AK = '10/21'; AQ = '0.1' } -DateColumns @('AJ', 'AK') -NumberColumns @('AQ') -Today ([datetime]'2026-10-20')
$w = @($plan['writes'] | ForEach-Object { $_['letter'] + $_['row'] + '=' + $_['value'] + $(if ($_['number']) { '#' } else { '' }) }) -join ','
Assert-Equal 'AK8=2026-10-21,AQ8=0.1#,AJ9=2026-10-21,AK9=2026-10-21,AQ9=0.1#' $w ('mark_rows: a list where = both task rows; whereMatch keeps own / own<-other rows only; M/d is this year; AQ is a number; the hand-typed AJ stays: ' + $w)
Assert-Equal 'QJODJSPP' (@($plan['notFound']) -join ',') 'mark_rows: a job whose rows are all someone else''s is not found (not written)'
Assert-Equal '2026-10-21' (ExcelMarkRows-Day -Text '10/21' -Today ([datetime]'2026-10-20')) 'mark_rows: 10/21 -> this year'
Assert-Equal '' (ExcelMarkRows-Day -Text 'tomorrow') 'mark_rows: not a date -> empty'

$ctx['DryRun'] = $true
$ctx['Session']['x'] = @{ kind = 'workbook'; value = $null }
$m = Invoke-TestStep 'excel.mark_rows' @{ workbook = 'x'; sheet = 'WBS'; keyColumn = 'A'; items = $items; keyField = 'JOB'; setValues = @{ AK = '2026-10-20' } }
$ctx['DryRun'] = $false
Assert-True ($m['ok'] -and $m['writtenCount'] -eq 0) 'mark_rows: a dry run writes nothing'

exit (Complete-Tests)
