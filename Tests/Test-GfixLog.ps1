#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'GfixLog.ps1')

Reset-Tests 'GfixLog'

# -- SS code + fragment --
Assert-Equal 'F' (Get-GfixSsCode 'JIDSF48S')   'SS code = 5th char (Substring(4,1))'
Assert-Equal 'J' (Get-GfixSsCode 'JIDSL48S')   'J-biz L receive correl uses SS code J'
Assert-Equal 'J' (Get-GfixSsCode 'JIGPLB1S')   'J-biz L receive correl with B1 suffix uses SS code J'
Assert-Equal 'Z' (Get-GfixSsCode 'JIDSL48S' 'Z') 'mapping SS override wins'
Assert-Equal ''  (Get-GfixSsCode 'JID')        'short id -> empty SS'
Assert-Equal '/appl/IDS/IDSVer1/gfix/recv/JIDSF48S F' (Get-GfixExpectedCommandFragment 'IDS' 'JIDSF48S') 'expected command fragment'
Assert-Equal '/appl/IDS/IDSVer1/gfix/recv/JIDSL48S J' (Get-GfixExpectedCommandFragment 'IDS' 'JIDSL48S') 'J-biz L command fragment'
$tsPattern = Get-GfixExpectedCommandPattern 'IDS' 'JIDSU86S'
Assert-True ([regex]::IsMatch('/appl/IDS/IDSVer1/gfix/recv/JIDSU86S.260729.10515511 U', $tsPattern)) 'command pattern accepts timestamped correl'
Assert-True ([regex]::IsMatch('/appl/IDS/IDSVer1/gfix/recv/JIDSU86S U', $tsPattern)) 'command pattern still accepts plain correl'
Assert-True (-not [regex]::IsMatch('/appl/IDS/IDSVer1/gfix/recv/JIDSU86SX U', $tsPattern)) 'command pattern rejects a sibling correl prefix'

# -- timestamp parse --
$ts = Get-GfixLogTimestamp "2026-05-29 10:59:29 INFO Command: 'x'"
Assert-True ($ts -is [datetime])                 'timestamp parses to datetime'
Assert-True ($null -eq (Get-GfixLogTimestamp 'no time here')) 'no timestamp -> null'

# -- Test-GfixCommandLine --
$good = "2026-05-29 10:59:29 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F'"
$frag = Get-GfixExpectedCommandFragment 'IDS' 'JIDSF48S'
Assert-True (Test-GfixCommandLine $good $frag)            'matching Command line'
Assert-True (-not (Test-GfixCommandLine 'random' $frag))  'non-Command line rejected'

# -- Find-GfixLogForCorrel against fixtures --
# -- run identity: what tells two candidate logs apart (pure, no I/O) --
$cmdTs    = "2026-07-29 10:51:55 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSU86S.260729.10515511 U'"
$cmdPlain = "2026-05-29 10:59:29 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F'"
Assert-Equal 'JIDSU86S.260729.10515511' (Get-GfixCommandCorrelToken $cmdTs 'IDS') 'command token keeps the batch stamp'
Assert-Equal 'JIDSF48S' (Get-GfixCommandCorrelToken $cmdPlain 'IDS') 'plain command token has no stamp'
Assert-Equal '' (Get-GfixCommandCorrelToken $cmdTs 'IGP') 'a token for another TO_CODE is not read'
Assert-Equal '' (Get-GfixCommandCorrelToken '' 'IDS') 'empty line has no token'

Assert-Equal '260729.10515511' (Get-GfixCorrelStamp 'JIDSU86S.260729.10515511') 'batch stamp extracted'
Assert-Equal '' (Get-GfixCorrelStamp 'JIDSU86S') 'plain id has no stamp'
Assert-Equal '' (Get-GfixCorrelStamp '') 'empty id has no stamp'

$ts1 = [datetime]'2026-06-01 09:00:00'
Assert-Equal (Get-GfixLogRunKey -CorrelToken 'JIDSK11S.260601.09000012' -Timestamp $ts1 -CommandLine 'a') `
             (Get-GfixLogRunKey -CorrelToken 'JIDSK11S.260601.09000012' -Timestamp $ts1 -CommandLine 'b') `
             'same token + timestamp = same run, whatever the file was called'
Assert-True ((Get-GfixLogRunKey -CorrelToken 'JIDSK11S.260601.09000012' -Timestamp $ts1 -CommandLine 'a') -ne `
             (Get-GfixLogRunKey -CorrelToken 'JIDSK11S.260602.11223344' -Timestamp $ts1 -CommandLine 'a')) `
             'different batch stamps are different runs'

# -- Select-GfixLogCandidate (pure: fabricated candidates, no files) --
$mk = {
    param([string]$file, [string]$token, $stamp)
    [pscustomobject]@{ File = $file; CorrelToken = $token; Timestamp = $stamp; CommandLine = ('cmd ' + $token) }
}
$dupPair = @(
    (& $mk 'C:\log\JIDSK11S_a.log'                 'JIDSK11S.260601.09000012' $ts1),
    (& $mk 'C:\log\JIDSK11S.260601.09000012_a.log' 'JIDSK11S.260601.09000012' $ts1)
)
$selDup = Select-GfixLogCandidate -Candidates $dupPair -CorrelIdS 'JIDSK11S'
Assert-Equal '' $selDup.Warning 'two spellings of one run produce no warning'
Assert-Equal 1 $selDup.Duplicates 'the collapsed duplicate is counted'
Assert-Equal 'C:\log\JIDSK11S.260601.09000012_a.log' $selDup.Chosen.File 'the stamped file name is kept'

$ts2 = [datetime]'2026-06-02 11:22:33'
$twoRuns = $dupPair + @((& $mk 'C:\log\JIDSK11S.260602.11223344_a.log' 'JIDSK11S.260602.11223344' $ts2))
$selTwo = Select-GfixLogCandidate -Candidates $twoRuns -CorrelIdS 'JIDSK11S'
Assert-True ($selTwo.Warning -ne '') 'two real runs warn'
Assert-Equal 2 (@($selTwo.Runs)).Count 'two runs survive'
Assert-Equal 'C:\log\JIDSK11S.260602.11223344_a.log' $selTwo.Chosen.File 'newest run chosen'

$selExact = Select-GfixLogCandidate -Candidates $twoRuns -CorrelIdS 'JIDSK11S.260601.09000012'
Assert-Equal '' $selExact.Warning 'an exact batch stamp settles the choice'
Assert-Equal 'C:\log\JIDSK11S.260601.09000012_a.log' $selExact.Chosen.File 'the named run wins over the newest'

$selMiss = Select-GfixLogCandidate -Candidates $twoRuns -CorrelIdS 'JIDSK11S.260609.99999999'
Assert-Equal 2 (@($selMiss.Runs)).Count 'a stamp matching no candidate falls back to every run'

Assert-True ($null -eq (Select-GfixLogCandidate -Candidates @() -CorrelIdS 'JIDSK11S').Chosen) 'no candidates -> nothing chosen'

$tmp = Join-Path ([System.IO.Path]::GetTempPath()) ('gfixlog_' + [guid]::NewGuid().ToString('N'))
New-Item -ItemType Directory -Path $tmp -Force | Out-Null
try {
    $older = @(
        "2026-05-29 10:59:29 INFO Start",
        "2026-05-29 10:59:29 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F'",
        "2026-05-29 10:59:30 INFO Done"
    )
    $newer = @(
        "2026-05-30 08:00:00 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSF48S F'"
    )
    $other = @(
        "2026-05-29 09:00:00 INFO Command: '/appl/IGP/shell/IGPLB001run.sh /appl/IGP/IGPVer1/gfix/recv/JIGPF05S F'"
    )
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSF48S_20260529_a.log') -Value $older -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSF48S_20260530_b.log') -Value $newer -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'JIGPF05S_x.log')          -Value $other -Encoding UTF8

    $timestamped = @(
        "2026-07-29 10:51:55 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSU86S.260729.10515511 U'"
    )
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSU86S.260729.10515511_a.log') -Value $timestamped -Encoding UTF8

    # single-correl match -> picks the newest, warns about multiple
    $res = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSF48S'
    Assert-Equal '' $res.Error 'JIDSF48S: no error'
    Assert-True ($null -ne $res.Chosen) 'JIDSF48S: chosen set'
    Assert-True ((Split-Path -Leaf $res.Chosen.File) -eq 'JIDSF48S_20260530_b.log') 'JIDSF48S: newest wins'
    Assert-True ($res.Warning -ne '') 'JIDSF48S: warns when multiple command-matching logs exist'
    Assert-True ($res.Chosen.Lines.Count -ge 1) 'JIDSF48S: whole-file lines returned'

    $resTs = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSU86S'
    Assert-Equal '' $resTs.Error 'plain correl finds timestamped receive command'
    Assert-True ((Split-Path -Leaf $resTs.Chosen.File) -eq 'JIDSU86S.260729.10515511_a.log') 'timestamped log file chosen'

    # zero match -> error, no chosen
    $res0 = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSF99S'
    Assert-True ($res0.Error -ne '') 'JIDSF99S: error set on zero match'
    Assert-True ($null -eq $res0.Chosen) 'JIDSF99S: no chosen on zero match'

    # -- duplicate spellings of ONE run must not look like an ambiguity ------
    # The same download saved under both the legacy plain name and the
    # timestamped name: identical command line, identical log timestamp.
    $runA = @(
        "2026-06-01 09:00:00 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSK11S.260601.09000012 K'"
    )
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSK11S_a.log')                   -Value $runA -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSK11S.260601.09000012_a.log')   -Value $runA -Encoding UTF8

    $dup = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSK11S'
    Assert-Equal '' $dup.Error 'JIDSK11S: no error'
    Assert-Equal '' $dup.Warning 'plain + timestamped spellings of ONE run do not warn'
    Assert-Equal 1 $dup.Duplicates 'the duplicate spelling is counted, not warned about'
    Assert-Equal 1 (@($dup.Runs)).Count 'both files collapse to a single run'
    Assert-Equal 'JIDSK11S.260601.09000012_a.log' (Split-Path -Leaf $dup.Chosen.File) `
        'the timestamped file name is the kept representative'
    Assert-Equal 2 (@($dup.Candidates)).Count 'both raw candidates are still reported'

    # A genuine second run of the same job IS an ambiguity worth warning about.
    $runB = @(
        "2026-06-02 11:22:33 INFO Command: '/appl/IDS/shell/IDSLB053run.sh /appl/IDS/IDSVer1/gfix/recv/JIDSK11S.260602.11223344 K'"
    )
    Set-Content -LiteralPath (Join-Path $tmp 'JIDSK11S.260602.11223344_a.log') -Value $runB -Encoding UTF8

    $two = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSK11S'
    Assert-True ($two.Warning -ne '') 'two genuinely different runs still warn'
    Assert-Equal 2 (@($two.Runs)).Count 'the two distinct runs survive deduplication'
    Assert-Equal 1 $two.Duplicates 'the duplicate spelling is still collapsed'
    Assert-Equal 'JIDSK11S.260602.11223344_a.log' (Split-Path -Leaf $two.Chosen.File) 'newest run wins'

    # ...unless the mapping row names the exact transfer, which settles it.
    $exact = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSK11S.260601.09000012'
    Assert-Equal '' $exact.Warning 'a batch-stamped Correl_ID_S picks its own run without warning'
    Assert-Equal 'JIDSK11S.260601.09000012_a.log' (Split-Path -Leaf $exact.Chosen.File) `
        'the run named by the mapping id is chosen, not the newest'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
