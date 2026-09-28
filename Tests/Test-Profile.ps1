# Test-Profile.ps1 -- P2-01 / P2-04 / P2-09: the host-open profile loads,
# passes the schema check, and every fixture under profiles/host-open/
# fixtures judges as expected.json says; a deliberately broken profile
# reports each error class; diff and skeleton. ASCII source; no param().

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Profile.ps1')
. (Join-Path $repoRoot 'kernel/ProfileCheck.ps1')

Reset-Tests 'Profile'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-profile-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$hostOpen = Join-Path (Join-Path $repoRoot 'profiles') 'host-open'

# --- 1. host-open loads and is schema-clean ---------------------------------------
Write-Host '  -- profiles/host-open'
$p = Read-EbiProfile -Dir $hostOpen -WorkDir $tmpRoot
Assert-True $p['ok'] ('host-open loads: ' + $p['message'])
Assert-True (@($p['missing']).Count -eq 1 -and $p['missing'][0] -eq 'calibration') 'host-open: only calibration.json is absent (no fallback tier yet)'
$prof = $p['value']
Assert-True ($prof['pages'].Count -eq 5 -and $prof['pages'].Contains('transferStatus') -and $prof['pages']['transferStatus']['role'] -eq 'list' -and $prof['pages']['hmResult']['role'] -eq 'record') 'host-open: the five pages of PROFILE-SCHEMA 3.0, keyed by page name'
$s = Test-EbiProfileSchema -Profile $prof -Missing $p['missing']
Assert-Equal 0 @($s['errors']).Count ('host-open: schema check has no errors' + $(if (@($s['errors']).Count) { ': ' + (@($s['errors']) -join ' | ') } else { '' }))
Assert-True (@($s['warnings']).Count -gt 0) 'host-open: the known holes (empty expired fingerprints, pages without grammar) are warnings'
Assert-True ((@($s['warnings']) -join ' ') -like '*fingerprint.expired is empty*') 'host-open: an empty expired fingerprint is named'
$fx = Invoke-EbiFixtureCheck -Dir $hostOpen -Profile $prof
foreach ($r in @($fx['results'])) { Assert-True $r['ok'] ('fixture ' + $r['page'] + '/' + $r['name'] + ': ' + $r['message']) }
Assert-True ($fx['ok'] -and $fx['passed'] -ge 12 -and $fx['failed'] -eq 0) ('host-open fixtures: ' + $fx['passed'] + ' passed, ' + $fx['failed'] + ' failed, ' + $fx['skipped'] + ' skipped')

# the mixed-run rule: the profile stores the legacy codes
$col = $null
foreach ($c in $prof['worklist']['columns']) { if ($c['name'] -eq 'GIFT_MQ_snap') { $col = $c } }
Assert-True ($null -ne $col -and $col['values']['ok'] -eq '1' -and $col['values']['ng'] -eq '2' -and $col['values']['pending'] -eq '0') 'host-open: GIFT_MQ_snap keeps the legacy 1 / 2 / 0 codes (P0-R11)'

# --- 2. a broken profile: every error class reported -----------------------------------
Write-Host '  -- broken profile'
function Copy-Deep { param($V) if ($V -is [System.Collections.IDictionary]) { $h = @{}; foreach ($k in $V.Keys) { $h[$k] = Copy-Deep $V[$k] }; return $h }; if ($V -is [System.Collections.IList] -and -not ($V -is [string])) { $l = New-Object System.Collections.ArrayList; foreach ($x in $V) { [void]$l.Add((Copy-Deep $x)) }; return ,$l.ToArray() }; return $V }
$bad = Copy-Deep $prof
$bad['rules']['transferStatus']['rules'][0]['else'] = 'ok'                       # else: ok
$bad['worklist']['key']['columns'] = @('Correl_ID_S')                             # JOB_NAME has role key but is no longer in key.columns
$bad['worklist']['columns'][6]['values']['maybe'] = '9'                           # unknown verdict in values
$bad['vocabulary']['roles'].Remove('document')                                    # four roles
$bad['grammar']['nosuchpage'] = @{ parser = 'regex'; pattern = 'x' }              # grammar keyed by a non-page
$s = Test-EbiProfileSchema -Profile $bad -Missing @()
$errs = @($s['errors']) -join "`n"
Assert-True ($errs -like '*else "ok"*never ok*') 'broken: else ok is refused'
Assert-True ($errs -like '*role: key but is not in key.columns*') 'broken: key column drift is refused'
Assert-True ($errs -like '*values has "maybe"*') 'broken: an unknown verdict code is refused'
Assert-True ($errs -like '*roles lacks "document"*') 'broken: a missing role is refused'
Assert-True ($errs -like '*"nosuchpage" is not a page*') 'broken: grammar keyed by a non-page is refused'
Assert-True (@($s['errors']).Count -ge 5) ('broken: ' + @($s['errors']).Count + ' errors in all')

# a fixture whose expectation is wrong is a failed check, not silence
$wrongDir = Join-Path $tmpRoot 'wrong'
New-Item -ItemType Directory -Path (Join-Path $wrongDir 'fixtures/transferStatus') -Force | Out-Null
Copy-Item -LiteralPath (Join-Path $hostOpen 'fixtures/transferStatus/ok.txt') -Destination (Join-Path $wrongDir 'fixtures/transferStatus/ok.txt')
[System.IO.File]::WriteAllText((Join-Path $wrongDir 'fixtures/transferStatus/expected.json'), '{ "ok.txt": { "key": "JIDSK05S", "verdict": "ng" }, "gone.txt": { "verdict": "ok" } }')
$fx = Invoke-EbiFixtureCheck -Dir $wrongDir -Profile $prof
Assert-True (-not $fx['ok'] -and $fx['failed'] -eq 2) 'fixtures: a wrong expectation and a missing file both fail'
$okRes = @(@($fx['results']) | Where-Object { $_['name'] -eq 'ok.txt' })
Assert-True ($okRes.Count -eq 1 -and $okRes[0]['message'] -like 'verdict: expected ng, got ok*') 'fixtures: the mismatch says expected vs got'

# --- 3. diff and skeleton ---------------------------------------------------------------
Write-Host '  -- diff / new'
$other = Copy-Deep $prof
$other['pages']['transferStatus']['label'] = 'changed'
$other['pages'].Remove('reportPreview')
$other['vocabulary']['sides']['third'] = 'x'
$d = Compare-EbiProfile -A $prof -B $other -NameA 'host-open' -NameB 'other'
Assert-True (@($d['changed']) -contains 'pages.transferStatus.label' -and (@($d['missing']) | Where-Object { $_ -like 'pages.reportPreview.*' }).Count -gt 0 -and (@($d['extra']) -contains 'vocabulary.sides.third')) 'diff: changed / only-in-a / only-in-b leaves'
$last = [string]@($d['lines'])[-1]
Assert-True ($last -match '^\d+ same, \d+ changed, \d+ only in host-open, 1 only in other$' -and @($d['changed']).Count -ge 1) ('diff: the summary line (' + $last + ')')
$skel = Join-Path $tmpRoot 'newprof'
$n = New-EbiProfileSkeleton -Dir $skel -Name 'newprof'
Assert-True ($n['ok'] -and (Test-Path -LiteralPath (Join-Path $skel 'pages.json')) -and (Test-Path -LiteralPath (Join-Path $skel 'worklist.json'))) 'new: skeleton files written'
$p2 = Read-EbiProfile -Dir $skel -WorkDir $tmpRoot
Assert-True ($p2['ok'] -and $p2['value']['pages'].Contains('examplePage') -and $p2['value']['pages'].Contains('_doc')) 'new: the skeleton loads and carries _doc keys'
$n2 = New-EbiProfileSkeleton -Dir $skel -Name 'newprof'
Assert-True (-not $n2['ok']) 'new: refuses to overwrite'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
