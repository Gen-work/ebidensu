$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'kernel/Trace.ps1')

Reset-Tests 'Trace'

$temp = Join-Path ([System.IO.Path]::GetTempPath()) ('ebidensu-trace-' + [guid]::NewGuid().ToString('N'))
try {
    Write-TraceEvent -WorkDir $temp -RunId 'r1' -Phase 'each' -Key 'A-001' -Tags @{ owner = 'alice' } -Action 'start' -Status 'start'
    Write-TraceEvent -WorkDir $temp -RunId 'r1' -Phase 'each' -Key 'A-001' -Tags @{ owner = 'alice'; attempt = 1 } -Action 'capture' -Status 'ok'
    Write-TraceEvent -WorkDir $temp -RunId 'r1' -Phase 'teardown' -Key 'A-001' -Tags @{} -Action 'finish' -Status 'ok' -Message 'done'

    $events = @(Read-TraceEvents -WorkDir $temp -RunId 'r1')
    Assert-Equal 3 $events.Count 'three appended events are read back'
    Assert-Equal 'A-001' $events[0].key 'generic key is preserved'
    Assert-Equal 'alice' $events[1].tags.owner 'arbitrary tags are preserved'
    Assert-Equal 'done' $events[2].message 'message is preserved'
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'correl_id_s')) 'legacy correl_id_s is absent'
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'job_name')) 'legacy job_name is absent'

    $tail = @(Read-TraceEvents -WorkDir $temp -RunId 'r1' -Tail 2)
    Assert-Equal 2 $tail.Count 'Tail limits returned events'
    Assert-Equal 'capture' $tail[0].action 'Tail returns the last events in order'

    # --- the trace lives in the run's own directory (spec/VOCABULARY.md 3.3) ---
    $expected = Join-Path (Join-Path (Join-Path $temp 'run') 'r1') 'trace.jsonl'
    Assert-Equal $expected (Get-TraceFile $temp 'r1') 'trace file is run/<runId>/trace.jsonl'
    Assert-True (Test-Path -LiteralPath $expected) 'the run directory is created on first write'
    Assert-Equal 'r1' $events[0].runId 'the event also carries its run id'

    # --- the timestamp field is named 'ts', matching the ledger record shape ---
    Assert-True ($events[0].PSObject.Properties.Name -contains 'ts') 'the timestamp field is named ts'
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'timestamp')) 'the old timestamp spelling is gone'
    $parsed = [datetime]::MinValue
    Assert-True ([datetime]::TryParse(([string]($events[0].ts)), [ref]$parsed)) 'ts round-trips as a real timestamp'

    # --- a blank run id buckets instead of losing the event ---
    Write-TraceEvent -WorkDir $temp -Phase 'each' -Key 'D-004' -Action 'start' -Status 'start'
    $unknown = @(Read-TraceEvents -WorkDir $temp -RunId 'unknown')
    Assert-Equal 1 $unknown.Count 'a blank run id still writes an event'
    Assert-Equal 'unknown' $unknown[0].runId 'the directory and the runId field agree'

    # --- a second run does not land in the first run's file ---
    Write-TraceEvent -WorkDir $temp -RunId 'r2' -Phase 'each' -Key 'B-002' -Action 'start' -Status 'start'
    $r1 = @(Read-TraceEvents -WorkDir $temp -RunId 'r1')
    $r2 = @(Read-TraceEvents -WorkDir $temp -RunId 'r2')
    Assert-Equal 3 $r1.Count 'run r1 is unchanged by a second run'
    Assert-Equal 1 $r2.Count 'run r2 has its own trace file'
    Assert-Equal 'B-002' $r2[0].key 'run r2 kept its own event'
    $none = @(Read-TraceEvents -WorkDir $temp -RunId 'never-ran')
    Assert-Equal 0 $none.Count 'an unknown run reads back empty'

    # --- Data carries structured payloads (STEP-CONTRACT 3.1 warnings channel) ---
    Assert-True (-not ($events[0].PSObject.Properties.Name -contains 'data')) 'data is absent when not supplied'
    $warn = @(@{ code = 'unrecognized_line'; message = '3 lines did not match'; data = @{ lines = @(4, 9, 12) } })
    Write-TraceEvent -WorkDir $temp -RunId 'r3' -Phase 'each' -Key 'C-003' -Action 'parse' -Status 'ok' `
        -Data @{ warnings = $warn; elapsedMs = 42 }
    $r3 = @(Read-TraceEvents -WorkDir $temp -RunId 'r3')
    Assert-Equal 1 $r3.Count 'the event carrying Data was written'
    Assert-Equal 42 $r3[0].data.elapsedMs 'Data survives the round trip'
    Assert-Equal 'unrecognized_line' $r3[0].data.warnings[0].code 'a nested warnings array survives'
    Assert-Equal 9 $r3[0].data.warnings[0].data.lines[1] 'nested warning data survives'

    # --- a half-written final line must not eat a Tail slot ---
    # A reader can see an incomplete final line while a writer is appending.
    [System.IO.File]::AppendAllText((Get-TraceFile $temp 'r1'), '{"ts":', (New-Object System.Text.UTF8Encoding($false)))
    $tailPartial = @(Read-TraceEvents -WorkDir $temp -RunId 'r1' -Tail 1)
    Assert-Equal 1 $tailPartial.Count 'Tail ignores an incomplete final line'
    Assert-Equal 'finish' $tailPartial[0].action 'Tail returns the last complete event'
    $stillThree = @(Read-TraceEvents -WorkDir $temp -RunId 'r1')
    Assert-Equal 3 $stillThree.Count 'the fragment is skipped, not counted'

    $bytes = [System.IO.File]::ReadAllBytes((Get-TraceFile $temp 'r1'))
    $hasBom = $bytes.Length -ge 3 -and $bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF
    Assert-True (-not $hasBom) 'trace JSONL has no UTF-8 BOM'
} finally {
    Remove-Item -LiteralPath $temp -Recurse -Force -ErrorAction SilentlyContinue
}

$fail = Complete-Tests
exit $fail
