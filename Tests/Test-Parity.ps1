# Test-Parity.ps1 -- P2-06, the half of the old-vs-new comparison that runs
# without an office PC: the SAME page texts through the legacy judgment
# (modules/verify/SnapVerify.ps1: ConvertFrom-MqPageText + Test-MqRecord,
# ConvertFrom-HmPageText + Test-HmAbend) and through the new engine
# (kernel/ProfileCheck.ps1 Invoke-EbiFixtureCase over profiles/host-open);
# the CSV the new engine writes read back by the legacy MappingStore /
# MqSnap pending logic; and a resume that does not repeat a capture.
# Every place the two engines differ on purpose is asserted AS a
# difference, so it stays visible. PNG size / crop parity needs the
# office PC and stays on the BACKLOG card. ASCII source; no param().

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'modules/verify/SnapVerify.ps1')
. (Join-Path $repoRoot 'MappingStore.ps1')
. (Join-Path $repoRoot 'kernel/Profile.ps1')
. (Join-Path $repoRoot 'kernel/ProfileCheck.ps1')
. (Join-Path $repoRoot 'kernel/Table.ps1')
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Runner.ps1')

Reset-Tests 'Parity'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-parity-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$utf8 = New-Object System.Text.UTF8Encoding($false)
$hostOpen = Join-Path (Join-Path $repoRoot 'profiles') 'host-open'
$prof = (Read-EbiProfile -Dir $hostOpen -WorkDir $tmpRoot)['value']
function Fx { param([string]$Page, [string]$Name) return [System.IO.File]::ReadAllText((Join-Path (Join-Path (Join-Path $hostOpen 'fixtures') $Page) $Name), $utf8) }
function NewVerdict { param([string]$Page, [string]$Text, [string]$Key, $Window = $null) return (Invoke-EbiFixtureCase -Text $Text -Page $prof['pages'][$Page] -Grammar $prof['grammar'][$Page] -Rules $prof['rules'][$Page] -Key $Key -TimeWindow $Window) }

# --- 1. MQ transfer status: legacy Test-MqRecord vs the new pipeline --------------------
Write-Host '  -- transferStatus: same text, both engines'
foreach ($case in @(@{ file = 'ok.txt'; want = 'ok' }, @{ file = 'ng-rtncd.txt'; want = 'ng' }, @{ file = 'single-digit-hour.txt'; want = 'ok' })) {
    $text = Fx 'transferStatus' $case['file']
    $old = Test-MqRecord -Parsed (ConvertFrom-MqPageText $text) -CorrelId 'JIDSK05S' -Expected $null
    $new = NewVerdict 'transferStatus' $text 'JIDSK05S'
    Assert-Equal $case['want'] $old['Verdict'] ('legacy MQ verdict for ' + $case['file'])
    Assert-Equal $case['want'] $new['verdict'] ('new MQ verdict for ' + $case['file'])
    Assert-True ($old['Verdict'] -eq $new['verdict']) ('PARITY MQ ' + $case['file'] + ': ' + $old['Verdict'] + ' == ' + $new['verdict'])
}
$oldParsed = ConvertFrom-MqPageText (Fx 'transferStatus' 'single-digit-hour.txt')
$newCase = NewVerdict 'transferStatus' (Fx 'transferStatus' 'single-digit-hour.txt') 'JIDSK05S'
Assert-True ($oldParsed['Rows'].Count -eq 2 -and $newCase['records'] -eq 3) 'DIFFERENCE (fixed): the legacy MQ parser drops the 7:54 row (\d{2} hour); the new grammar keeps all three'
# No Data: legacy ng; new = the empty fingerprint -> a gate (PROFILE-SCHEMA 3.1)
$old = Test-MqRecord -Parsed (ConvertFrom-MqPageText 'No Data!') -CorrelId 'JIDSK05S' -Expected $null -IsNoData $true
$new = NewVerdict 'transferStatus' (Fx 'transferStatus' 'no-data.txt') 'JIDSK05S'
Assert-True ($old['Verdict'] -eq 'ng' -and $new['page'] -eq 'empty' -and $new['verdict'] -eq '') 'DIFFERENCE (by spec): No Data is ng in the legacy code, an empty page (gate) in the new engine'
# a key the page does not have: legacy ng; new = record_not_found (a transient failure -> ask)
$old = Test-MqRecord -Parsed (ConvertFrom-MqPageText (Fx 'transferStatus' 'ok.txt')) -CorrelId 'ZZZZZZZZ' -Expected $null
$new = NewVerdict 'transferStatus' (Fx 'transferStatus' 'ok.txt') 'ZZZZZZZZ'
Assert-True ($old['Verdict'] -eq 'ng' -and $new['failure'] -eq 'not_found') 'DIFFERENCE (by spec): no row is ng in the legacy code, record_not_found (ask) in the new engine'
# time window: legacy takes the NEWEST row then checks the window; new prefers the newest INSIDE the window
$expected = [datetime]'2026-06-12T10:30:00'
$old = Test-MqRecord -Parsed (ConvertFrom-MqPageText (Fx 'transferStatus' 'ok.txt')) -CorrelId 'JIDSK05S' -Expected $expected -ToleranceMin 20
$new = NewVerdict 'transferStatus' (Fx 'transferStatus' 'ok.txt') 'JIDSK05S' @{ from = '2026/06/12 10:10:00'; to = '2026/06/12 10:50:00' }
Assert-True ($old['Verdict'] -eq 'ng' -and $old['MatchedRow'].No -eq 3 -and $new['verdict'] -eq 'ok' -and $new['matchedRow'] -eq 2) 'DIFFERENCE (recorded): with a window, legacy judges the newest row overall (11:01, outside -> ng); the new engine judges the newest row inside the window (10:32 -> ok)'
$old = Test-MqRecord -Parsed (ConvertFrom-MqPageText (Fx 'transferStatus' 'ok.txt')) -CorrelId 'JIDSK05S' -Expected ([datetime]'2026-06-12T11:00:00') -ToleranceMin 20
$new = NewVerdict 'transferStatus' (Fx 'transferStatus' 'ok.txt') 'JIDSK05S' @{ from = '2026/06/12 10:40:00'; to = '2026/06/12 11:20:00' }
Assert-True ($old['Verdict'] -eq 'ok' -and $new['verdict'] -eq 'ok' -and $new['matchedRow'] -eq 3) 'PARITY MQ: window around the newest run -> ok on both'

# --- 2. HM result: legacy Test-HmAbend vs the new pipeline --------------------------------
Write-Host '  -- hmResult: same text, both engines'
$hmText = Fx 'hmResult' 'ok-retried.txt'
$oldRows = @(ConvertFrom-HmPageText $hmText)
Assert-Equal 3 $oldRows.Count 'legacy HM parser: 3 rows'
$old = Test-HmAbend -Rows $oldRows -CorrelId 'JIDSK01S' -Expected $null
$new = NewVerdict 'hmResult' $hmText 'JIDSK01S'
Assert-True ($old['Verdict'] -eq 'ask' -and $new['verdict'] -eq 'ok' -and $new['matchedRow'] -eq 1) 'DIFFERENCE (recorded): without a window, legacy asks when ANY abend row exists; the new engine judges the newest run (11:05 normal -> ok)'
$old = Test-HmAbend -Rows $oldRows -CorrelId 'JIDSK01S' -Expected ([datetime]'2026-06-12T10:30:00') -ToleranceMin 20
$new = NewVerdict 'hmResult' $hmText 'JIDSK01S' @{ from = '2026/06/12 10:10:00'; to = '2026/06/12 10:50:00' }
Assert-True ($old['Verdict'] -eq 'ng' -and $new['verdict'] -eq 'ng' -and $new['matchedRow'] -eq 2) 'PARITY HM: window around the abend run -> ng on both'
$old = Test-HmAbend -Rows $oldRows -CorrelId 'JIDSK01S' -Expected ([datetime]'2026-06-12T11:00:00') -ToleranceMin 20
$new = NewVerdict 'hmResult' $hmText 'JIDSK01S' @{ from = '2026/06/12 10:40:00'; to = '2026/06/12 11:20:00' }
Assert-True ($old['Verdict'] -eq 'ok' -and $new['verdict'] -eq 'ok' -and $new['matchedRow'] -eq 1) 'PARITY HM: window around the newest normal run -> ok on both'
Assert-Equal ([string]$oldRows[1].CorrelId) ([string](ConvertFrom-EbiGrammar -Text $hmText -Grammar $prof['grammar']['hmResult'])['records'][1]['key']) 'PARITY HM: the abend row (extra empty cell) keys the same in both parsers'

# --- 3. the CSV the new engine writes, read by the legacy tools ---------------------------
Write-Host '  -- mixed run: new worklist, legacy readers'
$modules = Join-Path $repoRoot 'modules'
$reg = New-EbiRegistry -ModulesRoot $modules
foreach ($u in @('table.load', 'flow.checkpoint')) { $r = . Import-EbiStep -Registry $reg -Use $u; Assert-True $r['ok'] ('loads ' + $u) }
$csv = Join-Path $tmpRoot 'mapping_p.csv'
[System.IO.File]::WriteAllText($csv, "`"Correl_ID_S`",`"JOB_NAME`",`"GIFT_MQ_snap`"`r`n`"JIDSK05S`",`"JOB_A`",`"0`"`r`n`"JIDSK06S`",`"JOB_A`",`"0`"`r`n`"JIDSK07S`",`"JOB_B`",`"0`"`r`n", (New-Object System.Text.UTF8Encoding($true)))
$log = New-Object PSObject
$log | Add-Member -MemberType ScriptMethod -Name Info -Value { param($m) }
$log | Add-Member -MemberType ScriptMethod -Name Warn -Value { param($m) }
$ctx = @{ WorkDir = $tmpRoot; RunId = 'p'; Profile = $prof; Log = $log; DryRun = $false; Session = @{}; Item = $null; KeyColumns = @('Correl_ID_S', 'JOB_NAME'); Run = @{} }
$loaded = & (Get-EbiStep -Registry $reg -Use 'table.load')['Invoke'] @{ path = $csv; keyColumns = @(); mustExist = $true } $ctx
$wl = $loaded['resource']
$ctx['Item'] = $wl['rows'][0]; [void](& (Get-EbiStep -Registry $reg -Use 'flow.checkpoint')['Invoke'] @{ worklist = $wl; field = 'GIFT_MQ_snap'; value = 'ok'; bit = ''; key = '' } $ctx)
$ctx['Item'] = $wl['rows'][1]; [void](& (Get-EbiStep -Registry $reg -Use 'flow.checkpoint')['Invoke'] @{ worklist = $wl; field = 'GIFT_MQ_snap'; value = 'ng'; bit = ''; key = '' } $ctx)
$legacyRows = @(Import-Mapping -Path $csv)
Assert-True ($legacyRows.Count -eq 3 -and $legacyRows[0].GIFT_MQ_snap -eq '1' -and $legacyRows[1].GIFT_MQ_snap -eq '2' -and $legacyRows[2].GIFT_MQ_snap -eq '0') 'CSV: legacy Import-Mapping reads the codes the new engine stored (1 / 2 / 0)'
function Test-MqSnapDoneLegacy([string]$Value) { return ($Value -eq '1') }   # MqSnap.ps1:333 verbatim (that file has param(), so it cannot be dot-sourced)
Assert-True ((Test-MqSnapDoneLegacy '1') -and -not (Test-MqSnapDoneLegacy '2') -and -not (Test-MqSnapDoneLegacy '0')) 'CSV: MqSnap.ps1''s own done test (value -eq 1) sees ok done, ng and pending not done'
$pending = @(Get-PendingRows -Rows $legacyRows -Field 'GIFT_MQ_snap')
Assert-True ($pending.Count -eq 1 -and $pending[0].Correl_ID_S -eq 'JIDSK07S') 'CSV: Mark.ps1''s Get-PendingRows selects exactly the row the new engine left pending (its known rule: any non-0 counts as done, so the ng row is not re-offered there)'
Assert-True (Test-BitDone '3' 1) 'CSV: legacy bit test unchanged'
$bytes = [System.IO.File]::ReadAllBytes($csv)
Assert-True ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB) 'CSV: the rewritten file keeps the BOM Excel needs'

# --- 4. interrupted run -> resume does not repeat the capture ------------------------------
Write-Host '  -- resume without repeating a capture'
$fx = Join-Path $tmpRoot 'modules'
New-Item -ItemType Directory -Path (Join-Path $fx 'fake') -Force | Out-Null
$counter = Join-Path $tmpRoot 'shots.txt'
[System.IO.File]::WriteAllText((Join-Path $fx 'fake/fake.shot.ps1'), @'
$Manifest = @{ id = 'fake.shot'; group = 'fake'; summary = 'fixture capture: appends to a counter file'; tier = 'core'; effects = 'write'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ counter = @{ type = 'path'; required = $true } }; outputs = @{ path = @{ type = 'path' } }; failures = @( @{ id = 'save_failed'; transient = $true } ); example = @{ use = 'fake.shot'; with = @{ counter = 'c.txt' } } }
function Invoke-Step { param($In, $Ctx) if ($Ctx['DryRun']) { $Ctx.Log.Info('would shoot'); return @{ ok = $true; path = $In['counter'] } }; [System.IO.File]::AppendAllText($In['counter'], 'x'); return @{ ok = $true; path = $In['counter'] } }
'@, $utf8)
[System.IO.File]::WriteAllText((Join-Path $fx 'fake/fake.crash.ps1'), @'
$Manifest = @{ id = 'fake.crash'; group = 'fake'; summary = 'fixture: fails while a marker file exists (the Ctrl+C)'; tier = 'core'; effects = 'pure'; needs = @(); provides = @(); releases = @(); idempotent = $true
  inputs = @{ marker = @{ type = 'path'; required = $true } }; outputs = @{}; failures = @( @{ id = 'interrupted'; transient = $false } ); example = @{ use = 'fake.crash'; with = @{ marker = 'm.txt' } } }
function Invoke-Step { param($In, $Ctx) if (Test-Path -LiteralPath $In['marker']) { return @{ ok = $false; failure = 'interrupted'; message = 'operator pressed Ctrl+C' } }; return @{ ok = $true } }
'@, $utf8)
[System.IO.File]::WriteAllText((Join-Path $fx 'fake/fake.load.ps1'), @'
$Manifest = @{ id = 'fake.load'; group = 'fake'; summary = 'fixture worklist'; tier = 'core'; effects = 'read'; needs = @(); provides = @('worklist'); releases = @(); idempotent = $true
  inputs = @{}; outputs = @{ rowCount = @{ type = 'int' } }; failures = @( @{ id = 'never'; transient = $false } ); example = @{ use = 'fake.load'; with = @{ as = 'wl' } } }
function Invoke-Step { param($In, $Ctx) $rows = @(@{ K = 'A' }, @{ K = 'B' }, @{ K = 'C' }); return @{ ok = $true; resource = @{ path = ''; columns = @('K', 'done'); rows = $rows }; rowCount = 3 } }
'@, $utf8)
$marker = Join-Path $tmpRoot 'ctrlc.txt'
$wfPath = Join-Path $tmpRoot 'resume.json'
[System.IO.File]::WriteAllText($wfPath, ('{ "schema": 1, "id": "parity.resume", "source": { "table": "wl", "select": { "field": "done", "pendingWhen": "empty" }, "keyColumns": ["K"] }, "onError": { "policy": "fail" }, ' +
  '"setup": [ { "id": "load", "use": "fake.load", "with": { "as": "wl" } } ], ' +
  '"each": [ { "id": "shot", "use": "fake.shot", "with": { "counter": "' + ($counter -replace '\\', '/') + '" } }, { "id": "crash", "use": "fake.crash", "with": { "marker": "' + ($marker -replace '\\', '/') + '" } } ] }'), $utf8)
[System.IO.File]::WriteAllText($marker, 'stop', $utf8)
$wd = Join-Path $tmpRoot 'work'
$first = Invoke-EbiWorkflow -Path $wfPath -WorkDir $wd -ModulesRoot $fx -RunId 'r-int'
Assert-True (-not $first['ok'] -and ([System.IO.File]::ReadAllText($counter)).Length -eq 1) 'interrupted run: one capture happened, then the run stopped'
Remove-Item -LiteralPath $marker -Force
$second = Invoke-EbiWorkflow -Path $wfPath -WorkDir $wd -ModulesRoot $fx -RunId 'r-int' -Resume
Assert-True ($second['ok'] -and ([System.IO.File]::ReadAllText($counter)).Length -eq 3) 'resumed run: the first item''s capture is replayed from the ledger, not taken again; the other two are captured once each'
Assert-True (@($second['steps'] | Where-Object { $_['status'] -eq 'replayed' }).Count -ge 1) 'resumed run: the ledger replay is visible in the records'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
