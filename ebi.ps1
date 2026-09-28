#Requires -Version 5.1
# ============================================================
#  ebi.ps1 - the ebi-dance command-line entry point (P1-07..P1-10).
#
#    ebi.ps1 help [<step>]                     all steps by group / one manifest
#    ebi.ps1 lint    <workflow.json>           static checks (WORKFLOW-SCHEMA 9)
#    ebi.ps1 explain <workflow.json>           the execution plan, nothing runs
#    ebi.ps1 dryrun  <workflow.json>           every step in DryRun
#    ebi.ps1 run     <workflow.json>           the real thing
#    ebi.ps1 doctor                            PS version, Excel COM, Edge, encoding
#    ebi.ps1 catalog                           regenerate docs/ebi-dance/CATALOG.md + catalog.json
#    ebi.ps1 profile check|new|diff ...        PROFILE-SCHEMA 10 (P2-09)
#
#  run / dryrun options:
#    -WorkDir <dir>        default: the current directory
#    -Profile <name|dir>   default: the workflow's "profile" field under profiles/
#    -Resume [-RunId <id>] continue an unfinished run (no id: the newest one
#                          of this workflow); an unfinished run WITHOUT -Resume
#                          asks first instead of silently starting anew (P0-R16)
#    -Only <key,key>       only these rows (by key display), on top of pendingWhen
#    -Operator <name>      run.operator (default $env:USERNAME)
#    -Limit <n>            row cap
#    -Var k=v [-Var ...]   override workflow vars
#    -TimeWindow "<from>..<to>"  run.timeWindow for rules' within (P2-07), e.g.
#                          "2026/06/12 09:00..2026/06/12 12:00" or "09:00..12:00" (today)
#
#  Exit codes: 0 ok, 1 a step / lint / doctor failed, 2 usage, 3 the
#  operator cancelled. Has param(): call via -File or &, never dot-source.
# ============================================================
param(
    [Parameter(Position = 0)] [string]$Command = 'help',
    [Parameter(Position = 1)] [string]$Target  = '',
    [Parameter(Position = 2)] [string]$Target2 = '',
    [string]$WorkDir  = '',
    [string]$RunId    = '',
    [string]$Profile  = '',
    [string]$Only     = '',
    [string]$Operator = '',
    [int]$Limit       = 0,
    [string[]]$Var    = @(),
    [string]$TimeWindow = '',
    [switch]$Resume,
    [switch]$DryRun
)

$ErrorActionPreference = 'Stop'
# Capture switches BEFORE dot-sourcing anything (CLAUDE.md switch pattern).
$dryRunFlag = [bool]$DryRun.IsPresent
$resumeFlag = [bool]$Resume.IsPresent
try {
    [Console]::OutputEncoding = New-Object System.Text.UTF8Encoding($false)
} catch { }

. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Runner.ps1')
. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Help.ps1')
. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Lint.ps1')
. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Explain.ps1')
. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Profile.ps1')
. (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Parse.ps1')     # ConvertTo-EbiDateTime for -TimeWindow

function Show-EbiUsage {
    Write-Host ''
    Write-Host 'ebi - declarative evidence workflows' -ForegroundColor Cyan
    Write-Host '  ebi.ps1 help [<step>]              steps by group, or one manifest in full'
    Write-Host '  ebi.ps1 lint    <workflow.json>    static checks; exit 1 on errors'
    Write-Host '  ebi.ps1 explain <workflow.json>    the execution plan (nothing runs)'
    Write-Host '  ebi.ps1 dryrun  <workflow.json>    every step in DryRun'
    Write-Host '  ebi.ps1 run     <workflow.json>    [-Resume [-RunId <id>]] [-Only k1,k2] [-Operator n] [-Limit n] [-Var k=v] [-TimeWindow "from..to"]'
    Write-Host '  ebi.ps1 doctor                     PS version, Excel COM, Edge, encoding policy'
    Write-Host '  ebi.ps1 catalog                    regenerate docs/ebi-dance/CATALOG.md + catalog.json'
    Write-Host '  ebi.ps1 profile check <name|dir>   schema check + every fixture through grammar -> rules'
    Write-Host '  ebi.ps1 profile new   <name>       write profiles/<name>/ with commented skeleton files'
    Write-Host '  ebi.ps1 profile diff  <a> <b>      what still differs between two profiles'
    Write-Host '  common: -WorkDir <dir> (default .)  -Profile <name|dir> (default: the workflow''s "profile")'
    Write-Host ''
}

function Resolve-EbiCliWorkflowPath {
    param([string]$Given)
    if ([string]::IsNullOrWhiteSpace($Given)) { return '' }
    if ([System.IO.Path]::IsPathRooted($Given)) { return $Given }
    $candidate = Join-Path (Get-Location).ProviderPath $Given
    if (-not (Test-Path -LiteralPath $candidate)) { $candidate = Join-Path $PSScriptRoot $Given }
    return $candidate
}

function Read-EbiCliWorkflow {
    # -> @{ ok; value; path; message }
    param([string]$Given)
    $path = Resolve-EbiCliWorkflowPath -Given $Given
    if ($path -eq '') { return @{ ok = $false; value = $null; path = ''; message = 'a workflow file is required' } }
    $r = Read-EbiJson -Path $path
    if (-not $r['ok']) { return @{ ok = $false; value = $null; path = $path; message = $r['message'] } }
    if (-not ($r['value'] -is [hashtable])) { return @{ ok = $false; value = $null; path = $path; message = 'workflow JSON must be an object' } }
    return @{ ok = $true; value = $r['value']; path = $path; message = '' }
}

function Read-EbiCliProfile {
    # The -Profile option, else the workflow's "profile" field. -> @{ ok; value; message; name }
    # value is $null when the workflow says none / nothing is configured.
    param([hashtable]$Workflow, [string]$Option, [string]$WorkDirValue)
    $spec = if (-not [string]::IsNullOrWhiteSpace($Option)) { $Option } elseif ($null -ne $Workflow -and $Workflow.Contains('profile') -and $null -ne $Workflow['profile']) { [string]$Workflow['profile'] } else { '' }
    $dir = Resolve-EbiProfileDir -NameOrPath $spec
    if ($dir -eq '') { return @{ ok = $true; value = $null; message = ''; name = '' } }
    $r = Read-EbiProfile -Dir $dir -WorkDir $WorkDirValue
    if (-not $r['ok']) { return @{ ok = $false; value = $null; message = $r['message']; name = $spec } }
    return @{ ok = $true; value = $r['value']; message = ''; name = $r['name'] }
}

function ConvertFrom-EbiCliTimeWindow {
    # "from..to" -> @{ from; to } (ISO text) or $null; a bad text is usage.
    param([string]$Text)
    if ([string]::IsNullOrWhiteSpace($Text)) { return @{ ok = $true; value = $null; message = '' } }
    $parts = @($Text -split '\.\.')
    if ($parts.Count -ne 2) { return @{ ok = $false; value = $null; message = ('-TimeWindow must be "<from>..<to>", got "' + $Text + '"') } }
    $from = ConvertTo-EbiDateTime -Text $parts[0]; $to = ConvertTo-EbiDateTime -Text $parts[1]
    if (-not $from['ok'] -or -not $to['ok']) { return @{ ok = $false; value = $null; message = ('-TimeWindow: cannot read "' + $(if (-not $from['ok']) { $parts[0] } else { $parts[1] }) + '" as a time (yyyy/MM/dd H:mm[:ss] or H:mm)') } }
    if ($to['value'] -lt $from['value']) { return @{ ok = $false; value = $null; message = '-TimeWindow: "to" is before "from"' } }
    return @{ ok = $true; value = @{ from = $from['value'].ToString('yyyy-MM-ddTHH:mm:ss'); to = $to['value'].ToString('yyyy-MM-ddTHH:mm:ss') }; message = '' }
}

function ConvertFrom-EbiCliVars {
    # -Var k=v [-Var k2=v2] -> hashtable
    param([string[]]$Pairs)
    $h = @{}
    foreach ($p in @($Pairs)) {
        if ([string]::IsNullOrWhiteSpace($p)) { continue }
        $eq = $p.IndexOf('=')
        if ($eq -le 0) { Write-Host ('[WARN] -Var "' + $p + '" is not k=v; ignored') -ForegroundColor Yellow; continue }
        $h[$p.Substring(0, $eq)] = $p.Substring($eq + 1)
    }
    return $h
}

function Invoke-EbiDoctor {
    # Environment self-check. Prints one line per check; returns the number of failures.
    $fails = 0
    function Report { param([string]$Status, [string]$What, [string]$Detail) $c = switch ($Status) { 'OK' { 'Green' } 'WARN' { 'Yellow' } default { 'Red' } }; Write-Host ('  [{0,-4}] {1,-18} {2}' -f $Status, $What, $Detail) -ForegroundColor $c }
    Write-Host ''
    Write-Host 'ebi doctor' -ForegroundColor Cyan
    $v = $PSVersionTable.PSVersion
    if ($v.Major -ge 5) { Report 'OK' 'PowerShell' ([string]$v + $(if ($v.Major -ge 6) { '  (pwsh; the office PC runs Windows PowerShell 5.1)' } else { '' })) } else { Report 'FAIL' 'PowerShell' ('{0} -- 5.1 or newer is required' -f $v); $fails++ }
    $onWindows = ($env:OS -eq 'Windows_NT')
    if (-not $onWindows) { Report 'WARN' 'Platform' 'not Windows: Excel COM and Edge checks skipped (CI / dev box)' }
    else {
        try {
            $xl = New-Object -ComObject Excel.Application
            $ver = [string]$xl.Version
            $xl.Quit()
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($xl)
            Report 'OK' 'Excel COM' ('Excel.Application version ' + $ver)
        } catch { Report 'FAIL' 'Excel COM' ('cannot create Excel.Application: ' + $_.Exception.Message); $fails++ }
        $edge = @(Get-Process -Name msedge -ErrorAction SilentlyContinue)
        if ($edge.Count -gt 0) { Report 'OK' 'Edge' ('{0} process(es) running' -f $edge.Count) }
        else {
            $exe = @('C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe', 'C:\Program Files\Microsoft\Edge\Application\msedge.exe') | Where-Object { Test-Path -LiteralPath $_ }
            if (@($exe).Count -gt 0) { Report 'WARN' 'Edge' 'installed but not running (browser.ensure needs an open window)' } else { Report 'FAIL' 'Edge' 'msedge.exe not found'; $fails++ }
        }
    }
    $enc = Join-Path $PSScriptRoot 'Check-Encoding.ps1'
    if (Test-Path -LiteralPath $enc) {
        $out = & $enc -Root $PSScriptRoot 2>&1 | Out-String
        $rc = $LASTEXITCODE
        if ($rc -eq 0) { Report 'OK' 'Encoding policy' 'Check-Encoding.ps1 passed' } else { Report 'FAIL' 'Encoding policy' ('Check-Encoding.ps1 exit ' + $rc + ' -- run it for details'); $fails++ }
    } else { Report 'WARN' 'Encoding policy' 'Check-Encoding.ps1 not found' }
    $entries = @(Get-EbiCatalogEntries)
    $broken = @($entries | Where-Object { -not $_['ok'] })
    if ($broken.Count -eq 0) { Report 'OK' 'Step catalog' ('{0} step(s) load' -f $entries.Count) } else { Report 'FAIL' 'Step catalog' ('{0} of {1} step file(s) do not load: {2}' -f $broken.Count, $entries.Count, (($broken | ForEach-Object { $_['use'] }) -join ', ')); $fails++ }
    $spec = Get-EbiMustReleaseKinds
    if ($spec.Count -gt 0) { Report 'OK' 'Spec' ('STEP-CONTRACT.md readable; {0} resource kind(s)' -f $spec.Count) } else { Report 'WARN' 'Spec' 'docs/ebi-dance/spec/STEP-CONTRACT.md not readable; lint release checks off' }
    $wd = if ([string]::IsNullOrWhiteSpace($WorkDir)) { (Get-Location).ProviderPath } else { $WorkDir }
    $unfinished = @(Find-EbiUnfinishedRuns -WorkDir $wd)
    if ($unfinished.Count -gt 0) { Report 'WARN' 'Unfinished runs' ('{0} in {1}: {2}' -f $unfinished.Count, $wd, (($unfinished | ForEach-Object { [string]$_['runId'] }) -join ', ')) } else { Report 'OK' 'Unfinished runs' ('none in ' + $wd) }
    Write-Host ''
    return $fails
}

$cmd = $Command.ToLowerInvariant()
if ($cmd -eq 'help' -or $cmd -eq '-h' -or $cmd -eq '--help' -or $cmd -eq '') {
    if ([string]::IsNullOrWhiteSpace($Target)) {
        Show-EbiUsage
        foreach ($line in @(Format-EbiHelpList -Entries @(Get-EbiCatalogEntries))) { Write-Host $line }
        exit 0
    }
    $entry = Find-EbiHelpEntry -Entries @(Get-EbiCatalogEntries) -Use $Target
    if ($null -eq $entry) { Write-Host ('[ERROR] no step "' + $Target + '"; `ebi.ps1 help` lists them') -ForegroundColor Red; exit 1 }
    foreach ($line in @(Format-EbiHelpStep -Entry $entry)) { Write-Host $line }
    exit 0
}

if ($cmd -eq 'doctor') {
    $n = Invoke-EbiDoctor
    if ($n -gt 0) { exit 1 }
    exit 0
}

if ($cmd -eq 'catalog') {
    . (Join-Path (Join-Path $PSScriptRoot 'kernel') 'Docs.ps1')
    $w = Write-EbiCatalog
    if (-not $w['ok']) { Write-Host ('[ERROR] ' + $w['message']) -ForegroundColor Red; exit 1 }
    Write-Host ('wrote {0} and {1} ({2} step(s), {3} broken)' -f $w['markdownPath'], $w['jsonPath'], $w['steps'], $w['broken'])
    exit $(if ($w['broken'] -gt 0) { 1 } else { 0 })
}

if ($cmd -eq 'profile') {
    . (Join-Path (Join-Path $PSScriptRoot 'kernel') 'ProfileCheck.ps1')
    $sub = $Target.ToLowerInvariant()
    $wd = if ([string]::IsNullOrWhiteSpace($WorkDir)) { (Get-Location).ProviderPath } else { $WorkDir }
    if ($sub -eq 'check') {
        if ([string]::IsNullOrWhiteSpace($Target2)) { Write-Host '[ERROR] profile check <name|dir>' -ForegroundColor Red; exit 2 }
        $dir = Resolve-EbiProfileDir -NameOrPath $Target2
        if ($dir -eq '' -or -not (Test-Path -LiteralPath $dir -PathType Container)) { Write-Host ('[ERROR] profile directory not found: ' + $Target2) -ForegroundColor Red; exit 2 }
        $pr = Read-EbiProfile -Dir $dir -WorkDir $wd
        if (-not $pr['ok']) { Write-Host ('[ERROR] ' + $pr['message']) -ForegroundColor Red; exit 1 }
        $s = Test-EbiProfileSchema -Profile $pr['value'] -Missing $pr['missing']
        Write-Host ''
        Write-Host ('profile ' + $pr['name'] + ' (' + $dir + ')') -ForegroundColor Cyan
        foreach ($e in @($s['errors'])) { Write-Host ('  [ERROR] ' + $e) -ForegroundColor Red }
        foreach ($w in @($s['warnings'])) { Write-Host ('  [WARN ] ' + $w) -ForegroundColor Yellow }
        $fx = Invoke-EbiFixtureCheck -Dir $dir -Profile $pr['value']
        foreach ($r in @($fx['results'])) { Write-Host ('  [' + $(if ($r['ok']) { 'ok  ' } else { 'FAIL' }) + '] fixtures/' + $r['page'] + '/' + $r['name'] + '  ' + $r['message']) -ForegroundColor $(if ($r['ok']) { 'Gray' } else { 'Red' }) }
        $learned = 0
        if ($pr['overlay'] -is [System.Collections.IDictionary] -and $pr['overlay'].Contains('worklist') -and ($pr['overlay']['worklist'] -is [System.Collections.IDictionary]) -and $pr['overlay']['worklist'].Contains('key') -and ($pr['overlay']['worklist']['key'] -is [System.Collections.IDictionary]) -and $pr['overlay']['worklist']['key'].Contains('confirmedRules')) {
            $learned = @($pr['overlay']['worklist']['key']['confirmedRules']).Count
            if ($learned -gt 0) { Write-Host ('  [WARN ] ' + $learned + ' learned rule(s) in ebi.local.json not yet written back to ' + $pr['name'] + '/worklist.json (PROFILE-SCHEMA 6.6 b)') -ForegroundColor Yellow }
        }
        Write-Host ('  schema: {0} error(s), {1} warning(s); fixtures: {2} passed, {3} failed, {4} skipped' -f @($s['errors']).Count, @($s['warnings']).Count, $fx['passed'], $fx['failed'], $fx['skipped']) -ForegroundColor $(if ($s['ok'] -and $fx['ok']) { 'Green' } else { 'Red' })
        if ($s['ok'] -and $fx['ok']) { exit 0 }
        exit 1
    }
    if ($sub -eq 'new') {
        if ([string]::IsNullOrWhiteSpace($Target2) -or $Target2 -notmatch '^[A-Za-z0-9._-]+$') { Write-Host '[ERROR] profile new <name>  (letters, digits, . _ -)' -ForegroundColor Red; exit 2 }
        $n = New-EbiProfileSkeleton -Dir (Join-Path (Get-EbiDefaultProfilesRoot) $Target2) -Name $Target2
        if (-not $n['ok']) { Write-Host ('[ERROR] ' + $n['message']) -ForegroundColor Red; exit 1 }
        Write-Host ('wrote ' + $n['dir'])
        Write-Host ('fill it in, then: ebi.ps1 profile check ' + $Target2)
        exit 0
    }
    if ($sub -eq 'diff') {
        if ([string]::IsNullOrWhiteSpace($Target2) -or @($Var).Count -eq 0) { Write-Host '[ERROR] profile diff <a> -Var <b>   (the second profile goes in -Var)' -ForegroundColor Red; exit 2 }
        $da = Resolve-EbiProfileDir -NameOrPath $Target2; $db = Resolve-EbiProfileDir -NameOrPath ([string]$Var[0])
        if ($da -eq '' -or $db -eq '') { Write-Host ('[ERROR] profile directory not found: ' + $(if ($da -eq '') { $Target2 } else { [string]$Var[0] })) -ForegroundColor Red; exit 2 }
        $pa = Read-EbiProfile -Dir $da -WorkDir $wd; $pb = Read-EbiProfile -Dir $db -WorkDir $wd
        if (-not $pa['ok'] -or -not $pb['ok']) { Write-Host ('[ERROR] ' + $pa['message'] + $pb['message']) -ForegroundColor Red; exit 1 }
        $d = Compare-EbiProfile -A $pa['value'] -B $pb['value'] -NameA $pa['name'] -NameB $pb['name']
        foreach ($l in @($d['lines'])) { Write-Host ('  ' + $l) -ForegroundColor $(if ($l -like '- *') { 'Yellow' } elseif ($l -like '+ *') { 'Cyan' } elseif ($l -like '~ *') { 'Gray' } else { 'Green' }) }
        exit 0
    }
    Write-Host ('[ERROR] profile ' + $Target + ': check | new | diff') -ForegroundColor Red
    exit 2
}

if ($cmd -ne 'lint' -and $cmd -ne 'explain' -and $cmd -ne 'run' -and $cmd -ne 'dryrun') {
    Write-Host ('[ERROR] unknown command: ' + $Command) -ForegroundColor Red
    Show-EbiUsage
    exit 2
}

$wf = Read-EbiCliWorkflow -Given $Target
if (-not $wf['ok']) { Write-Host ('[ERROR] ' + $wf['message']) -ForegroundColor Red; if ($wf['path'] -eq '') { Show-EbiUsage }; exit 2 }

$wd = if ([string]::IsNullOrWhiteSpace($WorkDir)) { (Get-Location).ProviderPath } else { $WorkDir }
if (-not (Test-Path -LiteralPath $wd)) { New-Item -ItemType Directory -Path $wd -Force | Out-Null }
$wd = (Resolve-Path -LiteralPath $wd).ProviderPath

$prof = Read-EbiCliProfile -Workflow $wf['value'] -Option $Profile -WorkDirValue $wd
if (-not $prof['ok']) { Write-Host ('[ERROR] profile: ' + $prof['message']) -ForegroundColor Red; exit 2 }
$profileData = if ($null -ne $prof['value']) { $prof['value'] } else { @{} }

if ($cmd -eq 'lint') {
    $catalog = ConvertTo-EbiLintCatalog -Entries @(Get-EbiCatalogEntries)
    $res = Invoke-EbiLint -Workflow $wf['value'] -Catalog $catalog -Profile $(if ($null -ne $prof['value']) { $prof['value'] } else { $null }) -MustRelease (Get-EbiMustReleaseKinds)
    foreach ($line in @(Format-EbiLintReport -Result $res -Path $wf['path'])) {
        $color = if ($line -like '*[[]ERROR]*') { 'Red' } elseif ($line -like '*[[]WARN ]*') { 'Yellow' } elseif ($line -like '*-- OK*') { 'Green' } else { 'Gray' }
        Write-Host $line -ForegroundColor $color
    }
    if ($res['ok']) { exit 0 }
    exit 1
}

if ($cmd -eq 'explain') {
    $catalog = ConvertTo-EbiLintCatalog -Entries @(Get-EbiCatalogEntries)
    $plan = Format-EbiExplain -Workflow $wf['value'] -Catalog $catalog -Profile $(if ($null -ne $prof['value']) { $prof['value'] } else { $null })
    Write-Host ''
    foreach ($line in @($plan['lines'])) {
        $color = if ($line -like '*[[]DESTR]*' -or $line -like '*confirm:false*' -or $line -like '*UNKNOWN STEPS*') { 'Red' } elseif ($line -like '*[[]human]*') { 'Yellow' } else { 'Gray' }
        Write-Host ('  ' + $line) -ForegroundColor $color
    }
    Write-Host ''
    exit 0
}

# ---- run / dryrun ------------------------------------------------------------------
if ($cmd -eq 'dryrun') { $dryRunFlag = $true }
$wfId = [string]$wf['value']['id']
$unfinished = @(Find-EbiUnfinishedRuns -WorkDir $wd -WorkflowId $wfId)
if ($resumeFlag) {
    if ([string]::IsNullOrWhiteSpace($RunId)) {
        if ($unfinished.Count -eq 0) { Write-Host ('[ERROR] -Resume: no unfinished run of "' + $wfId + '" under ' + $wd) -ForegroundColor Red; exit 2 }
        $RunId = [string]$unfinished[0]['runId']
        Write-Host ('resuming the newest unfinished run: ' + $RunId) -ForegroundColor Cyan
    }
} elseif ($unfinished.Count -gt 0 -and -not $dryRunFlag) {
    # P0-R16: never silently start anew next to an unfinished run.
    $r = Show-EbiGate -Title 'UNFINISHED RUN' `
        -What @(('{0} unfinished run(s) of "{1}" under {2}:' -f $unfinished.Count, $wfId, $wd), (($unfinished | ForEach-Object { [string]$_['runId'] + ' (started ' + [string]$_['startedAt'] + ')' }) -join '; ')) `
        -Next @('r: resume the newest one', 'n: start a new run anyway', 'q: quit') `
        -Actions @(@{ key = 'r'; label = 'resume' }, @{ key = 'n'; label = 'new run' }, @{ key = 'q'; label = 'quit' }) -Auto 'n'
    if ($r['action'] -eq 'q') { exit 3 }
    if ($r['action'] -eq 'r') { $resumeFlag = $true; $RunId = [string]$unfinished[0]['runId'] }
}
$onlyList = $null
if (-not [string]::IsNullOrWhiteSpace($Only)) { $onlyList = @($Only -split ',' | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }) }

$tw = ConvertFrom-EbiCliTimeWindow -Text $TimeWindow
if (-not $tw['ok']) { Write-Host ('[ERROR] ' + $tw['message']) -ForegroundColor Red; exit 2 }
$summary = Invoke-EbiWorkflow -Path $wf['path'] -WorkDir $wd -RunId $RunId -DryRun:$dryRunFlag -Profile $profileData -Vars (ConvertFrom-EbiCliVars -Pairs $Var) -Operator $Operator -Only $onlyList -Limit $Limit -Resume:$resumeFlag -TimeWindow $tw['value']
if ([string]$summary['failure'] -eq 'cancelled') { exit 3 }
if ([bool]$summary['ok']) { exit 0 }
exit 1
