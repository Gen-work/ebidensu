$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'kernel/Trace.ps1')

Reset-Tests 'Trace'

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-trace-' + [guid]::NewGuid().ToString('N'))
try {
    Write-TraceEvent -WorkDir $temp -Phase 'each' -Key 'A-001' -Tags @{ owner = 'alice' } -Action 'start' -Status 'start'
    Write-TraceEvent -WorkDir $temp -Phase 'each' -Key 'A-001' -Tags @{ owner = 'alice'; attempt = 1 } -Action 'capture' -Status 'ok'
    Write-TraceEvent -WorkDir $temp -Phase 'teardown' -Key 'A-001' -Tags @{} -Action 'finish' -Status 'ok' -Message 'done'

    $events = @(Read-TraceEvents -WorkDir $temp)
    Assert-Equal 3 $events.Count 'three appended events are read back'
    Assert-Equal 'A-001' $events[0].key 'generic key is preserved'
    Assert-Equal 'alice' $events[1].tags.owner 'arbitrary tags are preserved'
    Assert-Equal 'done' $events[2].message 'message is preserved'
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'correl_id_s')) 'legacy correl_id_s is absent'
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'job_name')) 'legacy job_name is absent'

    $tail = @(Read-TraceEvents -WorkDir $temp -Tail 2)
    Assert-Equal 2 $tail.Count 'Tail limits returned events'
    Assert-Equal 'capture' $tail[0].action 'Tail returns the last events in order'

    $bytes = [System.IO.File]::ReadAllBytes((Get-TraceFile $temp))
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    Assert-True (-not $hasBom) 'trace JSONL has no UTF-8 BOM'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

$fail = Complete-Tests
exit $fail
