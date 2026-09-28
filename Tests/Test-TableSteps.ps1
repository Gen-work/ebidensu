# Test-TableSteps.ps1 -- kernel/Key.ps1 matching (P1-27), kernel/Table.ps1
# CSV (P1-24) and the table / flow / progress steps (P1-24..P1-29), run for
# real against a temp CSV. ASCII source; no param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Key.ps1')
. (Join-Path $repoRoot 'kernel/Table.ps1')
. (Join-Path $repoRoot 'kernel/Worklist.ps1')
. (Join-Path $repoRoot 'kernel/Trace.ps1')

Reset-Tests 'TableSteps'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-table-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$fwA = [string][char]0xFF21
$ja = '' + [char]0x8EE2 + [char]0x9001

function New-TestLog {
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) [void]$this.Lines.Add('info:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) [void]$this.Lines.Add('warn:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
    return $log
}
$testProfile = @{
    worklist = @{
        file = 'wl.csv'
        key = @{ columns = @('Correl_ID_S', 'JOB_NAME'); confirmedRules = @(@{ kind = 'suffix'; pattern = '\.\d{6}\.\d{8}$' }, @{ kind = 'fullwidth' }, @{ kind = 'case-insensitive' }) }
        columns = @(
            @{ name = 'Correl_ID_S'; role = 'key' }
            @{ name = 'JOB_NAME'; role = 'key' }
            @{ name = 'before_transferStatus'; role = 'verdict'; default = ''; values = @{ ok = '1'; ng = '2'; unknown = ''; pending = '0' } }
            @{ name = 'after_transferStatus'; role = 'verdict'; default = '' }
            @{ name = 'composed'; role = 'bitmask'; default = '0'; bits = @{ before = 1; after = 2; compare = 4 } }
            @{ name = 'note'; role = 'text'; default = '' }
        )
    }
}
function New-TestCtx { param([bool]$DryRun = $false, $Item = $null, $Profile = $null) if ($null -eq $Profile) { $Profile = $testProfile }; return @{ WorkDir = $tmpRoot; RunId = 'r1'; Profile = $Profile; Log = (New-TestLog); DryRun = $DryRun; Session = @{}; Item = $Item; KeyColumns = @('Correl_ID_S', 'JOB_NAME') } }
$reg = New-EbiRegistry -ModulesRoot (Join-Path $repoRoot 'modules')
foreach ($u in @('table.load', 'table.save', 'table.ensure_columns', 'table.select', 'table.key', 'table.set', 'flow.checkpoint', 'progress.event', 'progress.status')) {
    $r = . Import-EbiStep -Registry $reg -Use $u
    Assert-True $r['ok'] ('loads: ' + $u + ' ' + $r['message'])
}
function Step { param([string]$Use, [hashtable]$With, $Ctx) return (& (Get-EbiStep -Registry $reg -Use $Use)['Invoke'] $With $Ctx) }

# --- 1. Key.ps1 matching --------------------------------------------------------------
Write-Host '  -- Key.ps1 matching'
$rules = Get-EbiKeyDefaultRules
Assert-Equal 'exact' (Get-EbiKeyMatchTier -Value 'ABC123' -Key 'ABC123' -Rules $rules) 'tier exact'
Assert-Equal 'stripped' (Get-EbiKeyMatchTier -Value 'ABC123.260824.10515511' -Key 'ABC123' -Rules $rules) 'tier stripped: file has the stamp'
Assert-Equal 'stripped' (Get-EbiKeyMatchTier -Value 'ABC123' -Key 'ABC123.260824.10515511' -Rules $rules) 'tier stripped: key has the stamp'
Assert-Equal 'fullwidth' (Get-EbiKeyMatchTier -Value ($fwA + 'BC123') -Key 'ABC123' -Rules $rules) 'tier fullwidth'
Assert-Equal 'case' (Get-EbiKeyMatchTier -Value 'abc123' -Key 'ABC123' -Rules $rules) 'tier case'
Assert-Equal 'case' (Get-EbiKeyMatchTier -Value ($fwA + 'bc123.260824.10515511') -Key 'ABC123' -Rules $rules) 'tiers stack: stamp + fullwidth + case'
Assert-Equal '' (Get-EbiKeyMatchTier -Value 'ABC124' -Key 'ABC123' -Rules $rules) 'no match'
Assert-Equal '' (Get-EbiKeyMatchTier -Value 'abc123' -Key 'ABC123' -Rules @(@{ kind = 'suffix'; pattern = '\.\d{6}\.\d{8}$' })) 'a rule not declared never matches (no case rule -> case differs)'
Assert-Equal '' (Get-EbiKeyMatchTier -Value 'ABC123.260824.10515511' -Key 'ABC123' -Rules @(@{ kind = 'fullwidth' })) 'no suffix rule -> the stamped name is not the key'
Assert-Equal 'stripped' (Get-EbiKeyMatchTier -Value 'X_ABC123' -Key 'ABC123' -Rules @(@{ kind = 'prefix'; pattern = '^X_' })) 'a learned prefix rule'
Assert-Equal '' (Get-EbiKeyMatchTier -Value '' -Key 'ABC123') 'empty value never matches'
Assert-Equal 3 @(Get-EbiKeyRules -Profile $testProfile).Count 'rules come from the profile'
Assert-Equal 3 @(Get-EbiKeyRules -Profile @{}).Count 'no profile rules -> the defaults'

$rows = @(
    @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB_A' }
    @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB_B' }
    @{ Correl_ID_S = 'abc123.260824.10515511'; JOB_NAME = 'job_a' }
    @{ Correl_ID_S = 'DEF456'; JOB_NAME = 'JOB_A' }
)
$m = Find-EbiKeyMatches -Records $rows -Key 'ABC123 / JOB_A' -KeyColumns @('Correl_ID_S', 'JOB_NAME') -Rules $rules
Assert-True ($m['tier'] -eq 'exact' -and @($m['indexes']).Count -eq 1 -and $m['indexes'][0] -eq 0) 'composite key: exact row found, the case/stamp variant not preferred'
$m = Find-EbiKeyMatches -Records $rows -Key 'ABC123 / JOB_C' -KeyColumns @('Correl_ID_S', 'JOB_NAME') -Rules $rules
Assert-Equal '' $m['tier'] 'composite key: second column differs -> no match'
$m = Find-EbiKeyMatches -Records @($rows[1], $rows[2]) -Key 'ABC123 / JOB_A' -KeyColumns @('Correl_ID_S', 'JOB_NAME') -Rules $rules
Assert-True ($m['tier'] -eq 'case' -and $m['indexes'][0] -eq 1) 'composite key: only the variant left -> matched at the case tier'
$m = Find-EbiKeyMatches -Records @('ABC123.260824.10515511.dat', 'ABC123.260824.10515533.dat', 'DEF456.dat') -Key 'ABC123.dat' -Rules @(@{ kind = 'suffix'; pattern = '\.\d{6}\.\d{8}(?=\.dat$)' })
Assert-True ($m['tier'] -eq 'stripped' -and @($m['indexes']).Count -eq 2) 'strings: two stamped reruns both match, ambiguity is the caller''s'
$c = Find-EbiKeySafeCollisions -Rows @(@{ K = 'A_B'; J = 'C' }, @{ K = 'A'; J = 'B_C' }, @{ K = 'ABC/1'; J = 'x' }, @{ K = 'ABC_1'; J = 'x' }, @{ K = 'Z'; J = 'z' }) -KeyColumns @('K', 'J')
Assert-True (-not $c['ok'] -and @($c['collisions']).Count -eq 2 -and $c['collisions'][0]['keySafe'] -eq 'A_B_C' -and @($c['collisions'][0]['keys']).Count -eq 2) 'keySafe collisions: A_B+C vs A+B_C and ABC/1 vs ABC_1'
$c = Find-EbiKeySafeCollisions -Rows @(@{ K = 'A' }, @{ K = 'B' }) -KeyColumns @('K')
Assert-True $c['ok'] 'no collisions'
$cl = New-EbiCandidateList -Items @(@{ candidate = 'x'; evidence = @{ a = 1 } }, 'y') -SuggestIndex 1 -Reason 'newest' -Doubts 'd'
Assert-True (@($cl['candidates']).Count -eq 2 -and $cl['candidates'][0]['id'] -eq 'c1' -and $cl['candidates'][1]['candidate'] -eq 'y' -and $cl['suggestion']['id'] -eq 'c2' -and $cl['doubts'] -eq 'd') 'candidate list: P0-R4 shape'
$cl = New-EbiCandidateList -Items @('x') -SuggestIndex -1
Assert-True ($null -eq $cl['suggestion']) 'candidate list: no suggestion when none given'

# --- 2. Table.ps1 CSV -----------------------------------------------------------------
Write-Host '  -- Table.ps1 CSV'
Assert-Equal '"a","b ""q""",""' (ConvertTo-EbiCsvLine -Values @('a', 'b "q"', $null)) 'csv line: quoted, quotes doubled, null empty'
$csv = Join-Path $tmpRoot 'wl.csv'
$w = Write-EbiCsvAtomic -Path $csv -Columns @('Correl_ID_S', 'JOB_NAME', 'before_transferStatus', 'note') -Rows @(
    @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB_A'; before_transferStatus = ''; note = $ja }
    @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB_B'; before_transferStatus = '1'; note = 'a, "b"' }
    @{ Correl_ID_S = 'DEF456'; JOB_NAME = 'JOB_A'; before_transferStatus = '2'; note = '' }
)
Assert-True $w['ok'] 'csv written'
$bytes = [System.IO.File]::ReadAllBytes($csv)
Assert-True ($bytes[0] -eq 0xEF -and $bytes[1] -eq 0xBB -and $bytes[2] -eq 0xBF) 'csv has a UTF-8 BOM (Excel)'
Assert-True ((Get-ChildItem -LiteralPath $tmpRoot -Force -Filter '*.tmp.*' | Measure-Object).Count -eq 0) 'csv: no temp file left behind'
$r = Read-EbiCsv -Path $csv
Assert-True ($r['ok'] -and @($r['rows']).Count -eq 3 -and @($r['columns']).Count -eq 4 -and $r['columns'][3] -eq 'note') 'csv read: rows and header order'
Assert-True ($r['rows'][0] -is [hashtable] -and $r['rows'][0]['note'] -eq $ja -and $r['rows'][1]['note'] -eq 'a, "b"') 'csv read: hashtables, Japanese and quoted commas intact'
$r = Read-EbiCsv -Path (Join-Path $tmpRoot 'none.csv')
Assert-True (-not $r['ok'] -and -not $r['exists']) 'csv read: missing'
[System.IO.File]::WriteAllText((Join-Path $tmpRoot 'empty.csv'), '')
$r = Read-EbiCsv -Path (Join-Path $tmpRoot 'empty.csv')
Assert-True (-not $r['ok'] -and $r['exists']) 'csv read: empty file is an error'

# --- 3. table.load ----------------------------------------------------------------------
Write-Host '  -- table.load'
$ctx = New-TestCtx
$ret = Step 'table.load' @{ path = 'wl.csv'; keyColumns = @(); mustExist = $true } $ctx
Assert-True ($ret['ok'] -and $ret['rowCount'] -eq 3 -and $null -ne $ret['resource']) 'load: ok, resource is the worklist'
Assert-True (@($ret['columns']) -contains 'composed' -and @($ret['columns']) -contains 'after_transferStatus' -and $ret['resource']['rows'][0]['composed'] -eq '0') 'load: profile columns added with defaults'
Assert-True (@($ret['warnings'] | Where-Object { $_['code'] -eq 'columns_added' }).Count -eq 1) 'load: added columns reported'
$wl = $ret['resource']
$ret = Step 'table.load' @{ path = 'none.csv'; keyColumns = @(); mustExist = $true } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'load: missing file fails'
$ret = Step 'table.load' @{ path = 'none.csv'; keyColumns = @(); mustExist = $false } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['rowCount'] -eq 0 -and @($ret['columns']).Count -eq 6) 'load: mustExist=false -> empty table with the profile columns'
$ret = Step 'table.load' @{ path = 'none.csv'; keyColumns = @(); mustExist = $true } (New-TestCtx -DryRun $true)
Assert-True ($ret['ok'] -and $ret['rowCount'] -eq 0) 'load: dry run of a missing file is ok with a warning'
$ret = Step 'table.load' @{ path = 'wl.csv'; keyColumns = @('Nope'); mustExist = $true } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'key_column_missing') 'load: key column must exist'
[void](Write-EbiCsvAtomic -Path (Join-Path $tmpRoot 'dup.csv') -Columns @('Correl_ID_S', 'JOB_NAME') -Rows @(@{ Correl_ID_S = 'A_B'; JOB_NAME = 'C' }, @{ Correl_ID_S = 'A'; JOB_NAME = 'B_C' }))
$ret = Step 'table.load' @{ path = 'dup.csv'; keyColumns = @(); mustExist = $true } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'key_collision' -and $ret['message'] -like '*A_B_C*') 'load: keySafe collision refused and named'
[System.IO.File]::WriteAllText((Join-Path $tmpRoot 'bad.csv'), "`n`n")
$ret = Step 'table.load' @{ path = 'bad.csv'; keyColumns = @(); mustExist = $true } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'csv_invalid') 'load: unreadable csv'

# --- 4. table.select / table.key --------------------------------------------------------
Write-Host '  -- table.select / table.key'
$ret = Step 'table.select' @{ worklist = $wl; field = 'before_transferStatus'; pendingWhen = '!= ok'; only = @(); limit = 0 } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['selected'] -eq 2 -and $ret['total'] -eq 3 -and (@($ret['keyList']) -contains 'DEF456 / JOB_A')) 'select: != ok keeps the blank and the ng (stored 2) rows; the stored 1 is ok'
$ret = Step 'table.select' @{ worklist = $wl; field = 'before_transferStatus'; pendingWhen = '== ng'; only = @(); limit = 0 } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['selected'] -eq 1 -and $ret['rows'][0]['Correl_ID_S'] -eq 'DEF456') 'select: == ng through verdict.values'
$ret = Step 'table.select' @{ worklist = $wl; field = 'composed'; pendingWhen = 'bit !before'; only = @(); limit = 1 } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['selected'] -eq 1 -and $ret['total'] -eq 3) 'select: bit by name, limit'
$ret = Step 'table.select' @{ worklist = $wl; field = 'x'; pendingWhen = '!= done'; only = @(); limit = 0 } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'input_invalid') 'select: a sixth form is refused'
$ret['rows'] = $null
$sel = Step 'table.select' @{ worklist = $wl; field = ''; pendingWhen = 'always'; only = @('ABC123 / JOB_B'); limit = 0 } (New-TestCtx)
Assert-True ($sel['ok'] -and $sel['selected'] -eq 1) 'select: only'
$sel['rows'][0]['note'] = 'changed on the copy'
Assert-True ($wl['rows'][1]['note'] -ne 'changed on the copy') 'select: rows are copies, the worklist is untouched'

$ret = Step 'table.key' @{ key = 'ABC123'; among = @('DEF456.dat', 'ABC123.260824.10515511.dat'); rules = @(@{ kind = 'suffix'; pattern = '\.\d{6}\.\d{8}\.dat$' }, @{ kind = 'suffix'; pattern = '\.dat$' }) } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 1 -and $ret['matchedBy'] -eq 'stripped') 'key: one match through rules'
$ret = Step 'table.key' @{ key = 'ABC123 / JOB_A'; among = $wl['rows']; rules = @() } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['index'] -eq 0 -and ($ret['match'] -is [hashtable])) 'key: records compared by the key columns'
$ret = Step 'table.key' @{ key = 'ABC123'; among = @('abc123', 'Abc123'); rules = @() } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'ambiguous' -and @($ret['candidates']['candidates']).Count -eq 2 -and $ret['found'] -eq 2) 'key: two same-tier hits -> ambiguous with candidates'
$ret = Step 'table.key' @{ key = 'ZZZ'; among = @('a', 'b'); rules = @() } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'key_not_found') 'key: none'

# --- 5. table.set / flow.checkpoint ------------------------------------------------------
Write-Host '  -- table.set / flow.checkpoint'
$ret = Step 'table.set' @{ worklist = $wl; key = 'ABC123 / JOB_A'; field = 'note'; value = 'hand'; bit = ''; clear = $false } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['updated'] -eq 1 -and $wl['rows'][0]['note'] -eq 'hand') 'set: by key'
Assert-True ((Read-EbiCsv -Path $csv)['rows'][0]['note'] -eq 'hand') 'set: flushed to disk'
$ret = Step 'table.set' @{ worklist = $wl; key = 'ABC123 / JOB_A'; field = 'before_transferStatus'; value = 'ng'; bit = ''; clear = $false } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['stored'] -eq '2' -and $wl['rows'][0]['before_transferStatus'] -eq '2') 'set: logical ng stored as the profile code 2'
$ret = Step 'table.set' @{ worklist = $wl; key = ''; field = 'composed'; value = ''; bit = 'after'; clear = $false } (New-TestCtx -Item $wl['rows'][2])
Assert-True ($ret['ok'] -and $ret['stored'] -eq '2' -and $wl['rows'][2]['composed'] -eq '2') 'set: named bit on the current item'
$ret = Step 'table.set' @{ worklist = $wl; key = ''; field = 'composed'; value = ''; bit = 'before'; clear = $false } (New-TestCtx -Item $wl['rows'][2])
Assert-True ($ret['ok'] -and $ret['stored'] -eq '3') 'set: second bit ORs in'
$ret = Step 'table.set' @{ worklist = $wl; key = ''; field = 'composed'; value = ''; bit = 'after'; clear = $true } (New-TestCtx -Item $wl['rows'][2])
Assert-True ($ret['ok'] -and $ret['stored'] -eq '1') 'set: clear a bit'
$ret = Step 'table.set' @{ worklist = $wl; key = ''; field = 'composed'; value = ''; bit = 'nosuch'; clear = $false } (New-TestCtx -Item $wl['rows'][2])
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'bit_unknown') 'set: an undeclared bit name is refused (no 1/2/4 guessing)'
$ret = Step 'table.set' @{ worklist = $wl; key = 'ZZZ / Q'; field = 'note'; value = 'x'; bit = ''; clear = $false } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'row_not_found') 'set: unknown key'
$ret = Step 'table.set' @{ worklist = $wl; key = ''; field = 'note'; value = 'x'; bit = ''; clear = $false } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'row_not_found') 'set: no key and no current item'
$ret = Step 'table.set' @{ worklist = $wl; key = 'ABC123 / JOB_A'; field = 'nosuch'; value = 'x'; bit = ''; clear = $false } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'column_missing') 'set: unknown column'
$ret = Step 'table.set' @{ worklist = $wl; key = 'abc123 / job_a'; field = 'note'; value = 'via case rule'; bit = ''; clear = $false } (New-TestCtx)
Assert-True ($ret['ok'] -and $wl['rows'][0]['note'] -eq 'via case rule') 'set: the row is found through the case-insensitive rule (no -eq of its own)'
$ctx = New-TestCtx -DryRun $true
$ret = Step 'table.set' @{ worklist = $wl; key = 'ABC123 / JOB_A'; field = 'note'; value = 'dry'; bit = ''; clear = $false } $ctx
Assert-True ($ret['ok'] -and $ret['updated'] -eq 0 -and $wl['rows'][0]['note'] -eq 'via case rule' -and $ctx['Log'].Lines.Count -eq 1) 'set: dry run changes nothing'

$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'before_transferStatus'; value = 'ok'; bit = ''; key = '' } (New-TestCtx -Item $wl['rows'][0])
Assert-True ($ret['ok'] -and $ret['stored'] -eq '1' -and $ret['key'] -eq 'ABC123 / JOB_A' -and $wl['rows'][0]['before_transferStatus'] -eq '1') 'checkpoint: current item, ok -> stored 1'
Assert-True ((Read-EbiCsv -Path $csv)['rows'][0]['before_transferStatus'] -eq '1') 'checkpoint: flushed'
$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'after_transferStatus'; value = 'unknown'; bit = ''; key = 'DEF456 / JOB_A' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['stored'] -eq 'unknown' -and $wl['rows'][2]['after_transferStatus'] -eq 'unknown') 'checkpoint: no values map -> the logical value is stored; by key'
$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'composed'; value = ''; bit = 'compare'; key = '' } (New-TestCtx -Item $wl['rows'][2])
Assert-True ($ret['ok'] -and $ret['stored'] -eq '5') 'checkpoint: named bit ORed onto the existing 1'
$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'before_transferStatus'; value = '1'; bit = ''; key = '' } (New-TestCtx -Item $wl['rows'][0])
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'value_invalid') 'checkpoint: a stored code in the workflow is refused'
$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'before_transferStatus'; value = 'ok'; bit = ''; key = '' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'row_not_found') 'checkpoint: outside each with no key'
$ctx = New-TestCtx -DryRun $true -Item $wl['rows'][1]
$ret = Step 'flow.checkpoint' @{ worklist = $wl; field = 'before_transferStatus'; value = 'ng'; bit = ''; key = '' } $ctx
Assert-True ($ret['ok'] -and $ret['stored'] -eq '2' -and $wl['rows'][1]['before_transferStatus'] -eq '1') 'checkpoint: dry run reports the code but writes nothing'

# --- 6. table.ensure_columns / table.save -----------------------------------------------
Write-Host '  -- table.ensure_columns / table.save'
$ret = Step 'table.ensure_columns' @{ worklist = $wl; columns = @{ extra = 'x' } } (New-TestCtx)
Assert-True ($ret['ok'] -and @($ret['added']) -contains 'extra' -and $wl['rows'][0]['extra'] -eq 'x' -and (@($wl['columns']) -contains 'extra')) 'ensure_columns: extra added with default'
Assert-True ((Read-EbiCsv -Path $csv)['columns'] -contains 'extra') 'ensure_columns: flushed'
$ret = Step 'table.ensure_columns' @{ worklist = $wl; columns = @{ extra = 'x' } } (New-TestCtx)
Assert-True ($ret['ok'] -and @($ret['added']).Count -eq 0) 'ensure_columns: idempotent'
$ret = Step 'table.save' @{ worklist = $wl; path = 'copy.csv' } (New-TestCtx)
Assert-True ($ret['ok'] -and (Test-Path -LiteralPath (Join-Path $tmpRoot 'copy.csv')) -and $ret['rowCount'] -eq 3) 'save: to another path'
$ret = Step 'table.save' @{ worklist = @{ path = ''; columns = @('a'); rows = @() }; path = '' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'no_path') 'save: no path anywhere'
$ret = Step 'table.save' @{ worklist = $wl; path = 'dry.csv' } (New-TestCtx -DryRun $true)
Assert-True ($ret['ok'] -and -not (Test-Path -LiteralPath (Join-Path $tmpRoot 'dry.csv'))) 'save: dry run writes nothing'

# --- 7. progress.status / progress.event --------------------------------------------------
Write-Host '  -- progress.status / progress.event'
$ret = Step 'progress.status' @{ worklist = $wl; field = ''; groupBy = 'JOB_NAME' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['summary'].Contains('before_transferStatus') -and $ret['summary'].Contains('composed')) 'status: every verdict / bitmask column'
$s = $ret['summary']['before_transferStatus']
Assert-True ($s['total'] -eq 3 -and $s['done'] -eq 2 -and $s['pending'] -eq 1 -and $s['ng'] -eq 1) 'status: counts through verdict.values (stored 1, 1, 2 -> 2 done, 1 pending which is ng)'
$s2 = $ret['summary']['composed']
Assert-True ($s2['total'] -eq 3 -and $s2['done'] -eq 0 -and $null -ne $s2['bits'] -and $s2['bits']['before'] -eq 1) 'status: bitmask column reports per-bit done counts'
Assert-True (@($ret['lines']).Count -ge 4 -and $ret['lines'][0] -like 'field*total*') 'status: ascii table lines'
$ret = Step 'progress.status' @{ worklist = $wl; field = 'nosuch'; groupBy = '' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'column_missing') 'status: unknown field'
$ret = Step 'progress.event' @{ action = 'checked'; status = 'ok'; message = 'm'; key = ''; data = @{ n = 1 } } (New-TestCtx -Item $wl['rows'][0])
Assert-True ($ret['ok'] -and $ret['key'] -eq 'ABC123 / JOB_A') 'event: key from the current item'
$evs = @(Read-TraceEvents -WorkDir $tmpRoot -RunId 'r1')
Assert-True ($evs.Count -eq 1 -and $evs[0]['action'] -eq 'checked' -and $evs[0]['key'] -eq 'ABC123 / JOB_A' -and $evs[0]['data']['n'] -eq 1) 'event: written to the run trace'
$ret = Step 'progress.event' @{ action = 'x'; status = 'info'; message = ''; key = ''; data = @{} } (New-TestCtx -DryRun $true)
Assert-True ($ret['ok'] -and @(Read-TraceEvents -WorkDir $tmpRoot -RunId 'r1').Count -eq 1) 'event: dry run writes nothing'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
