#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'GoAnywhereLogFetchLib.ps1')

Reset-Tests 'GoAnywhereLogFetch'

$t = "`t"
# Page layout WITHOUT a User column and with a trailing 'resolved' column
# (the shape seen on the completed-jobs page in 2026-10).
$text = @(
    'GoAnywhere'
    "JobNoHeader${t}Project${t}Folder${t}Status${t}Start${t}End${t}Sec${t}Method${t}Resolved"
    "1000000000101${t}LXXX_VER1_IF0009_FTP${t}/lxxxusr/Send${t}${t}2026-10-05 13:55:12${t}2026-10-05 13:55:12${t}0.16${t}Trigger${t}"
    "1000000000102${t}LXXX_DEV1_IF0001_001_SSH${t}/aaausr/Receive${t}${t}2026-10-05 13:59:45${t}2026-10-05 13:59:50${t}4.88${t}Trigger${t}"
    "1000000000103${t}LXXX_DEV1_IF0001_001_SSH${t}/aaausr/Receive${t}${t}2026-10-05 14:15:50${t}2026-10-05 14:15:52${t}2.39${t}Trigger${t}"
    "1000000000104${t}LXXX_VER1_IF0010_001_SSH${t}/bbb2usr/Receive${t}${t}2026-10-05 14:26:04${t}2026-10-05 14:26:12${t}8.38${t}Trigger${t}"
    "1000000000105${t}LYYY_VER1_IF0001_001_SSH${t}/bbb2usr/Receive${t}${t}2026-10-07 10:00:00${t}2026-10-07 10:00:05${t}5.00${t}Trigger${t}"
    "1000000000106${t}LXXX_VER1_IF0001_FTP${t}/lxxxusr/Send${t}${t}2026-10-05 13:59:45${t}2026-10-05 13:59:45${t}0.13${t}Trigger${t}"
    'displaying 1 - 6 of 6'
) -join "`r`n"

$rows = @(ConvertFrom-GoAnywhereJobList $text)
Assert-Equal 6 $rows.Count 'parses the 6 numeric-JobNo rows, skips header/footer'
Assert-Equal '/aaausr/Receive' $rows[1].Folder 'folder found by leading slash'
Assert-Equal '2026-10-05 13:59:45' $rows[1].StartTime 'start = first datetime field (no User column)'
Assert-Equal '2026-10-05 13:59:50' $rows[1].EndTime 'end = second datetime field'
# The older layout WITH a User column parses the same
$withUser = @(ConvertFrom-GoAnywhereJobList "1000000000201${t}P_IF0002_001_SSH${t}/aaausr/Receive${t}${t}admin${t}2026-07-01 11:02:36${t}2026-07-01 11:02:41${t}5.24${t}Trigger")
Assert-Equal '2026-07-01 11:02:36' $withUser[0].StartTime 'User column present: start still found'

Assert-Equal '0001' (ConvertTo-GoAnywhereIfNo '1') 'IF number padded to 4'
Assert-Equal '0012' (ConvertTo-GoAnywhereIfNo 'IF0012') 'IF prefix stripped'
Assert-True (Test-GoAnywhereIfMatch 'LXXX_DEV1_IF0001_001_SSH' '0001') 'IF0001 matches DEV1 project'
Assert-True (Test-GoAnywhereIfMatch 'LXXX_VER1_IF0001_001_SSH' '0001') 'IF0001 matches VER1 project'
Assert-True (-not (Test-GoAnywhereIfMatch 'LXXX_VER1_IF0010_001_SSH' '0001')) 'IF0001 does not match IF0010'
Assert-True (-not (Test-GoAnywhereIfMatch 'LXXX_VER1_IF1001_001_SSH' '0001')) 'IF0001 does not match IF1001'

# Several IF numbers per target (received under one, resent under another)
Assert-Equal '0001|0002' ((ConvertTo-GoAnywhereIfNos '1, 0002') -join '|') 'IF list: comma separated, padded'
Assert-Equal '0001|0002' ((ConvertTo-GoAnywhereIfNos @('0001', '2', '0001')) -join '|') 'IF list: array, duplicates dropped'
$multi = @(Get-GoAnywhereCandidateRows -Rows $rows -IfNo '0001,0010' -GoFolder '/bbb2usr/')
Assert-Equal '1000000000104|1000000000105' ((@($multi | ForEach-Object { $_.JobNo }) | Sort-Object) -join '|') 'IF list: candidates of every listed IF'

$systems = @{ AAA = @{ goFolder = '/aaausr/'; dest = 'AAA' }; BBB = @{ goFolder = '/bbb2usr/'; dest = 'BBB2' } }
$targets = @(
    [pscustomobject]@{ Correl = 'JAAAL25M'; IfNo = '0001'; Recv = 'AAA' }
    [pscustomobject]@{ Correl = 'JAAAL26M'; IfNo = '0001'; Recv = 'AAA' }
    [pscustomobject]@{ Correl = 'JBBBC51M'; IfNo = '0001'; Recv = 'BBB' }
    [pscustomobject]@{ Correl = 'JBBBK99M'; IfNo = '0099'; Recv = 'BBB' }
    [pscustomobject]@{ Correl = 'JCCCX01M'; IfNo = '0001'; Recv = 'CCC' }
)
$plan = Get-GoAnywhereFetchPlan -Targets $targets -Rows $rows -Systems $systems
Assert-Equal 2 $plan.Items[0].Candidates.Count 'shared IF: both receive runs are candidates (send row excluded)'
Assert-Equal '1000000000103' $plan.Items[0].Candidates[0].JobNo 'candidates newest first'
Assert-Equal 1 $plan.Items[2].Candidates.Count 'other folder: only its own receive job'
Assert-Equal '1000000000105' $plan.Items[2].Candidates[0].JobNo 'same IF in another folder is a different job'
Assert-Equal 3 @($plan.JobNumbers).Count 'distinct jobs deduped across correls sharing an IF'
Assert-Equal 2 @($plan.Missing).Count 'IF0099 and unknown recv are missing'
Assert-True (-not $plan.Items[4].KnownRecv) 'unknown recv flagged'

$merged = @(Merge-GoAnywhereRows -Existing $rows -New @($rows[0], $withUser[0]))
Assert-Equal 7 $merged.Count 'merge dedupes by job number'

$logA = "2026-10-05 13:59:47 INFO Command: '/appl/AAA/shell/run.sh /appl/AAA/AAAVer1/gfix/recv/JAAAL25M J'"
$logB = "2026-10-05 14:15:51 INFO Command: '/appl/AAA/shell/run.sh /appl/AAA/AAAVer1/gfix/recv/JAAAL26S J'"
$logs = @(
    [pscustomobject]@{ JobNo = '1000000000102'; StartTime = '2026-10-05 13:59:45'; Text = $logA }
    [pscustomobject]@{ JobNo = '1000000000103'; StartTime = '2026-10-05 14:15:50'; Text = $logB }
)
$s25 = Select-GoAnywhereLogForCorrel -Correl 'JAAAL25M' -Logs $logs
Assert-Equal 'ok' $s25.Status 'content names the correl'
Assert-Equal '1000000000102' $s25.Chosen.JobNo 'shared IF: the log naming this correl wins, not the newest'
$s26 = Select-GoAnywhereLogForCorrel -Correl 'JAAAL26M' -Logs $logs
Assert-Equal '1000000000103' $s26.Chosen.JobNo 'last char free: JAAAL26M owns the JAAAL26S log'
Assert-True (-not (Test-GoAnywhereLogNamesCorrel "/gfix/recv/JAAAL251M J" 'JAAAL25M')) 'a longer id is not a match'
Assert-True (Test-GoAnywhereLogNamesCorrel "/gfix/recv/JAAAL25M.261005.13594512 J" 'JAAAL25M') 'transfer stamp suffix allowed'
$rerun = @($logs[0], [pscustomobject]@{ JobNo = '1000000000110'; StartTime = '2026-10-06 09:00:00'; Text = $logA })
$sR = Select-GoAnywhereLogForCorrel -Correl 'JAAAL25M' -Logs $rerun
Assert-Equal '1000000000110' $sR.Chosen.JobNo 'reruns: newest wins'
Assert-Equal 1 @($sR.Others).Count 'reruns: older one reported'
Assert-Equal 'nocontent' (Select-GoAnywhereLogForCorrel -Correl 'JZZZZ01M' -Logs $logs).Status 'no log names it -> nocontent'
Assert-Equal 'nolog' (Select-GoAnywhereLogForCorrel -Correl 'JAAAL25M' -Logs @()).Status 'nothing downloaded -> nolog'

Assert-Equal '1000000000102.log' (Format-GoAnywhereLogName -Format '' -Correl 'JAAAL25M' -JobNo '1000000000102' -StartTime '2026-10-05 13:59:45' -OrigName '1000000000102.log') 'default: the name as downloaded'
Assert-Equal 'JAAAL25M_1000000000102.log' (Format-GoAnywhereLogName -Format '{correl}_{jobNo}.log' -Correl 'JAAAL25M' -JobNo '1000000000102' -StartTime '' -OrigName 'x.log') 'correl + jobNo tokens'
Assert-Equal 'JAAAL25M_20261005_135945.log' (Format-GoAnywhereLogName -Format '{correl}_{start}.log' -Correl 'JAAAL25M' -JobNo '1' -StartTime '2026-10-05 13:59:45' -OrigName 'x.log') 'start token'
Assert-Equal '1000000000102.log' (Format-GoAnywhereLogName -Format '{orig}.log' -Correl 'J' -JobNo '1' -StartTime '' -OrigName '1000000000102.log') 'orig token'

# Config read (UTF-8 JSON, Japanese in the path survives)
$tmp = Join-Path ([IO.Path]::GetTempPath()) ('galf_' + [guid]::NewGuid().ToString('N') + '.json')
$jp = [string][char]0x53D7 + [char]0x4FE1
$json = '{ "destRoot": "\\\\srv\\' + $jp + '\\LOG", "systems": { "AAA": { "goFolder": "/aaausr/", "dest": "AAA" } }, "jobs": [ { "correl": " JAAAL25M ", "if": "1", "recv": "AAA" }, { "correl": "JAAAL26M", "if": "0001,0002", "recv": "AAA" } ] }'
[IO.File]::WriteAllText($tmp, $json, [Text.UTF8Encoding]::new($false))
try {
    $cfg = Read-GoAnywhereLogFetchConfig $tmp
    Assert-Equal ('\\srv\' + $jp + '\LOG') $cfg.DestRoot 'config: Japanese destRoot read as UTF-8'
    Assert-Equal 'JAAAL25M' $cfg.Targets[0].Correl 'config: correl trimmed'
    Assert-Equal '0001' $cfg.Targets[0].IfNo 'config: IF padded'
    Assert-Equal '0001,0002' $cfg.Targets[1].IfNo 'config: several IF numbers kept'
    Assert-Equal '/aaausr/' $cfg.Systems['AAA'].goFolder 'config: systems'
    Assert-Equal '' $cfg.NameFormat 'config: nameFormat optional'
} finally { Remove-Item -LiteralPath $tmp -Force -ErrorAction SilentlyContinue }

exit (Complete-Tests)
