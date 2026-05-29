#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'GfixLog.ps1')

Reset-Tests 'GfixLog'

# -- SS code + fragment --
Assert-Equal 'F' (Get-GfixSsCode 'JIDSF48S')   'SS code = 5th char (Substring(4,1))'
Assert-Equal ''  (Get-GfixSsCode 'JID')        'short id -> empty SS'
Assert-Equal '/appl/IDS/IDSVer1/gfix/recv/JIDSF48S F' (Get-GfixExpectedCommandFragment 'IDS' 'JIDSF48S') 'expected command fragment'

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

    # single-correl match -> picks the newest, warns about multiple
    $res = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSF48S'
    Assert-Equal '' $res.Error 'JIDSF48S: no error'
    Assert-True ($null -ne $res.Chosen) 'JIDSF48S: chosen set'
    Assert-True ((Split-Path -Leaf $res.Chosen.File) -eq 'JIDSF48S_20260530_b.log') 'JIDSF48S: newest wins'
    Assert-True ($res.Warning -ne '') 'JIDSF48S: warns (prefix filter matched only its own files -> 1, so check broad)'
    Assert-True ($res.Chosen.Lines.Count -ge 1) 'JIDSF48S: whole-file lines returned'

    # zero match -> error, no chosen
    $res0 = Find-GfixLogForCorrel -LogDir $tmp -ToCode 'IDS' -CorrelIdS 'JIDSF99S'
    Assert-True ($res0.Error -ne '') 'JIDSF99S: error set on zero match'
    Assert-True ($null -eq $res0.Chosen) 'JIDSF99S: no chosen on zero match'
} finally {
    Remove-Item -LiteralPath $tmp -Recurse -Force -ErrorAction SilentlyContinue
}

exit (Complete-Tests)
