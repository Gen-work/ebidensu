# Test-Steps.ps1 -- the browser / screen / file step groups (P1-11..P1-23)
# plus the two kernel libraries they share, kernel/Native.ps1 (Win32,
# SendKeys, clipboard) and kernel/Image.ps1 (GDI+). Nothing here touches a
# window or GDI+: the pure helpers are called directly, every step is loaded
# through the registry and DRY-RUN with its manifest example (STEP-CONTRACT
# 6.1: a dry run must not reach the real system), and the file steps run
# for real against a temp folder. ASCII source; no param() block.
#
# Run: powershell -File Tests\Test-Steps.ps1

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/Registry.ps1')
. (Join-Path $repoRoot 'kernel/Json.ps1')
. (Join-Path $repoRoot 'kernel/Native.ps1')
. (Join-Path $repoRoot 'kernel/Image.ps1')
. (Join-Path $repoRoot 'kernel/Key.ps1')
. (Join-Path $repoRoot 'kernel/Context.ps1')   # Test-EbiTemplateString

Reset-Tests 'Steps'

$modulesRoot = Join-Path $repoRoot 'modules'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-steps-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$fw0 = [string][char]0xFF10   # full-width '0'
$fwA = [string][char]0xFF21   # full-width 'A'

function New-TestLog {
    $log = New-Object PSObject
    $log | Add-Member -MemberType NoteProperty -Name Lines -Value (New-Object System.Collections.ArrayList)
    $log | Add-Member -MemberType ScriptMethod -Name Info  -Value { param($m) [void]$this.Lines.Add('info:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Warn  -Value { param($m) [void]$this.Lines.Add('warn:' + $m) }
    $log | Add-Member -MemberType ScriptMethod -Name Debug -Value { param($m) }
    return $log
}
function New-TestCtx {
    param([bool]$DryRun = $false, [string]$WorkDir = '')
    if ($WorkDir -eq '') { $WorkDir = $tmpRoot }
    return @{ WorkDir = $WorkDir; RunId = 'test'; Profile = @{}; Log = (New-TestLog); DryRun = $DryRun; Session = @{ mainWindow = @{ kind = 'window'; value = 4242; registeredBy = 'test' } }; Item = $null; KeyColumns = @() }
}

$reg = New-EbiRegistry -ModulesRoot $modulesRoot
$catalog = @(Get-EbiStepCatalog -ModulesRoot $modulesRoot)
$expectedSteps = @(
    'browser.ensure', 'browser.focus_body', 'browser.send_keys', 'browser.tab_to', 'browser.fill', 'browser.submit',
    'browser.read_text', 'browser.wait_for', 'browser.assert_page', 'browser.navigate', 'browser.find',
    'screen.capture_window', 'screen.capture_region', 'screen.fit_window', 'screen.crop', 'screen.save',
    'file.write_json', 'file.read_json', 'file.find', 'file.assert_exists'
)

# --- 1. every step loads and dry-runs on its own example ------------------------
Write-Host '  -- dry run of every step from its manifest example'
$uses = @(@($catalog) | Where-Object { $_['ok'] } | ForEach-Object { $_['use'] })
foreach ($u in $expectedSteps) { Assert-True ($uses -contains $u) ("catalog has " + $u) }
foreach ($u in $uses) {
    # EVERY shipped step, not only the groups this file owns (P1-36 seed)
    $r = . Import-EbiStep -Registry $reg -Use $u
    Assert-True $r['ok'] ("loads: " + $u + ' ' + $r['message'])
    if (-not $r['ok']) { continue }
    $m = $r['Entry']['Manifest']
    $ctx = New-TestCtx -DryRun $true
    $with = @{}
    if ($m['example'].Contains('with') -and $null -ne $m['example']['with']) { foreach ($k in $m['example']['with'].Keys) { $with[[string]$k] = $m['example']['with'][$k] } }
    # A template in the example stands for a value the workflow supplies; the
    # dry run gets a fixture of the declared type instead.
    foreach ($k in @($with.Keys)) {
        $v = $with[$k]
        if (-not ($v -is [string]) -or -not (Test-EbiTemplateString $v)) { continue }
        $type = if ($m['inputs'].Contains($k)) { [string]$m['inputs'][$k]['type'] } else { 'string' }
        $enum = @(if ($m['inputs'].Contains($k) -and $m['inputs'][$k].Contains('enum') -and $null -ne $m['inputs'][$k]['enum']) { $m['inputs'][$k]['enum'] })
        if ($enum.Count -gt 0) { $with[$k] = $enum[0]; continue }
        # Named fixtures where the step's contract needs a particular shape
        # (what the workflow's template would have produced).
        $named = @{
            value       = 'ok'                                            # {{steps.gate.out.code}}
            code        = 'unknown'                                       # so a gate really asks (and auto-answers)
            candidates  = @{ candidates = @(@{ id = 'c1'; candidate = 'fixture'; evidence = @{ source = 'test' } }); suggestion = @{ id = 'c1'; reason = 'fixture' }; doubts = '' }
            grammar     = @{ parser = 'regex'; pattern = '^(?<key>\S+)\s+(?<time>.+)$' }
            rules       = @{ rules = @(@{ field = 'key'; op = 'present'; else = 'unknown'; message = 'fixture' }); default = 'ok' }
            fingerprint = @{ ok = @('fixture') }
            records     = @(@{ key = 'fixture'; time = '2026/09/28 9:00:00' })
            record      = @{ key = 'fixture' }
            text        = 'fixture 2026/09/28 9:00:00'
            key         = 'fixture'
        }
        if ($named.Contains($k)) { $with[$k] = $named[$k]; continue }
        switch ($type) {
            'int'  { $with[$k] = 100 }
            'bool' { $with[$k] = $false }
            'map'  { $with[$k] = @{ ok = @('fixture') } }
            'list' { $with[$k] = @('fixture') }
            'rect' { $with[$k] = @{ x = 0; y = 0; w = 1; h = 1 } }
            'path' { $with[$k] = 'fixture/' + $k }
            default { $with[$k] = 'fixture' }
        }
    }
    if ($with.Contains('as')) { $ctx['Session'] = @{} }   # a provides step registers its own
    else { $ctx['Session']['wl'] = @{ kind = 'worklist'; value = @{ path = (Join-Path $tmpRoot 'wl.csv'); columns = @('Correl_ID_S', 'JOB_NAME', 'before_transferStatus', 'composed', 'note'); rows = @(@{ Correl_ID_S = 'ABC123'; JOB_NAME = 'J1'; before_transferStatus = ''; composed = '0'; note = '' }) }; registeredBy = 'test' } }
    if ($ctx['Session'].Contains('wl')) { $ctx['Item'] = $ctx['Session']['wl']['value']['rows'][0] }
    $ctx['KeyColumns'] = @('Correl_ID_S', 'JOB_NAME')
    $res = Resolve-EbiStepInputs -Manifest $m -With $with -Session $ctx['Session']
    Assert-True $res['ok'] ("example inputs resolve: " + $u + ' ' + $res['message'])
    if (-not $res['ok']) { continue }
    $ret = & $r['Entry']['Invoke'] $res['In'] $ctx
    $chk = Test-EbiStepReturn -Manifest $m -Return $ret -WantsResource ($res['As'] -ne '')
    Assert-True $chk['ok'] ("dry-run return honours the contract: " + $u + ' ' + $chk['message'])
    Assert-True ($ret['ok'] -eq $true) ("dry-run is ok: " + $u + ' ' + $(if ($ret.Contains('message')) { $ret['message'] } else { '' }))
    $isUiOrWrite = ([string]$m['effects'] -in @('ui', 'write', 'destructive'))
    if ($isUiOrWrite) { Assert-True ($ctx['Log'].Lines.Count -gt 0) ("dry-run says what it would do: " + $u) }
}
Assert-True (-not (Test-Path -LiteralPath 'function:Invoke-Step')) 'no bare Invoke-Step left after loading'
Assert-True ((Get-ChildItem -LiteralPath $tmpRoot -Recurse -File | Measure-Object).Count -eq 0) 'the dry runs wrote nothing'

# --- 2. kernel/Native.ps1 pure helpers -------------------------------------------
Write-Host '  -- Native.ps1 pure helpers'
Assert-Equal '' (Resolve-EbiWorkPath -PathValue '' -WorkDir $tmpRoot) 'work path: empty stays empty'
$sep = [System.IO.Path]::DirectorySeparatorChar
Assert-Equal ($tmpRoot + $sep + 'capture' + $sep + 'x.png') (Resolve-EbiWorkPath -PathValue 'capture/x.png' -WorkDir $tmpRoot) 'work path: relative goes under WorkDir with native separators'
$rooted = Join-Path $tmpRoot 'abs.png'
Assert-Equal $rooted (Resolve-EbiWorkPath -PathValue $rooted -WorkDir 'C:\elsewhere') 'work path: rooted stays'
Assert-Equal 'rel/x' (Resolve-EbiWorkPath -PathValue 'rel/x' -WorkDir '') 'work path: no WorkDir leaves it alone'
Assert-True ((ConvertTo-EbiHandle 4242) -eq [IntPtr]4242) 'handle: int -> IntPtr'
Assert-True ((ConvertTo-EbiHandle $null) -eq [IntPtr]::Zero) 'handle: null -> zero'
Assert-True ((ConvertTo-EbiHandle 'abc') -eq [IntPtr]::Zero) 'handle: junk -> zero'
$txt = Join-Path $tmpRoot 'sub/t.txt'
$w = Write-EbiTextFile -Path $txt -Text ('a' + [char]0x3042)
Assert-True $w['ok'] 'text file: written'
$bytes = [System.IO.File]::ReadAllBytes($txt)
Assert-True ($bytes[0] -eq 0x61) 'text file: no BOM'
Assert-Equal ('a' + [char]0x3042) ([System.IO.File]::ReadAllText($txt)) 'text file: UTF-8 round trip'

# --- 3. kernel/Image.ps1 pure helpers --------------------------------------------
Write-Host '  -- Image.ps1 pure helpers'
$g = Get-EbiCropGeometry -Width 100 -Height 50 -Left 6 -Top 6 -Right 6 -Bottom 6
Assert-True ($g['ok'] -and $g['width'] -eq 88 -and $g['height'] -eq 38 -and -not $g['noop']) 'crop geometry: 100x50 - 6 each side = 88x38'
$g = Get-EbiCropGeometry -Width 10 -Height 10 -Left 5 -Right 5
Assert-True (-not $g['ok'] -and $g['message'] -like '*too small*') 'crop geometry: nothing left is an error'
$g = Get-EbiCropGeometry -Width 10 -Height 10
Assert-True ($g['ok'] -and $g['noop']) 'crop geometry: zero everywhere is a noop'
$g = Get-EbiCropGeometry -Width 10 -Height 10 -Left -3
Assert-True ($g['ok'] -and $g['width'] -eq 10) 'crop geometry: a negative side reads as 0'
$g = Get-EbiCropGeometry -Width 0 -Height 10
Assert-True (-not $g['ok']) 'crop geometry: empty image is an error'
$s = Resolve-EbiCropSides -CropPx 6
Assert-True ($s['left'] -eq 6 -and $s['top'] -eq 6 -and $s['right'] -eq 6 -and $s['bottom'] -eq 6) 'crop sides: -1 inherits CropPx'
$s = Resolve-EbiCropSides -CropPx 6 -Top 10 -Bottom 0
Assert-True ($s['left'] -eq 6 -and $s['top'] -eq 10 -and $s['bottom'] -eq 0) 'crop sides: explicit sides win, 0 means 0'
$r = Resolve-EbiScreenRegion -X 100 -Y 100 -W 200 -H 100 -BoundsX 0 -BoundsY 0 -BoundsW 1920 -BoundsH 1080
Assert-True (-not $r['clamped'] -and $r['edges'] -eq '' -and $r['w'] -eq 200) 'region: inside is untouched'
$r = Resolve-EbiScreenRegion -X -50 -Y 1000 -W 200 -H 200 -BoundsX 0 -BoundsY 0 -BoundsW 1920 -BoundsH 1080
Assert-True ($r['clamped'] -and $r['x'] -eq 0 -and $r['w'] -eq 150 -and $r['y'] -eq 1000 -and $r['h'] -eq 80) 'region: x pulled in (width shrinks by the overhang) and height cut'
Assert-Equal 'x,height' $r['edges'] 'region: edges named'
$r = Resolve-EbiScreenRegion -X 3000 -Y 0 -W 100 -H 100 -BoundsX 0 -BoundsY 0 -BoundsW 1920 -BoundsH 1080
Assert-True ($r['clamped'] -and $r['w'] -eq 0) 'region: entirely off screen -> width 0'
$c = Invoke-EbiCropPng -Path (Join-Path $tmpRoot 'nope.png') -Left 1
Assert-True (-not $c['ok'] -and $c['failure'] -eq 'file_not_found') 'crop png: missing file is file_not_found, no throw'
$ps = Get-EbiPngSize -Path (Join-Path $tmpRoot 'nope.png')
Assert-True (-not $ps['ok']) 'png size: missing file reported'

# --- 4. browser pure helpers ---------------------------------------------------------
Write-Host '  -- browser helpers'
$pt = BrowserFocusBody-Point -Rect @{ X = 40; Y = 40; W = 800; H = 600 } -OffsetX 150 -OffsetY 150
Assert-True ($pt['x'] -eq 190 -and $pt['y'] -eq 190) 'focus_body: click point = window origin + offsets'
Assert-Equal '{TAB}' (BrowserTabTo-Sequence -Shift $false) 'tab_to: Tab'
Assert-Equal '+{TAB}' (BrowserTabTo-Sequence -Shift $true) 'tab_to: Shift+Tab'
Assert-True (BrowserWaitFor-Matches -Text 'abc ABC123 def' -Needle 'ABC123') 'wait_for: contains'
Assert-True (-not (BrowserWaitFor-Matches -Text 'abc abc123 def' -Needle 'ABC123')) 'wait_for: case-sensitive'
Assert-True (-not (BrowserWaitFor-Matches -Text 'abc' -Needle '')) 'wait_for: empty needle never matches'
Assert-True (-not (BrowserWaitFor-Matches -Text '' -Needle 'x')) 'wait_for: empty text never matches'

$fp = @{ ok = @('Transfer status', 'Records:'); loading = @('Loading...'); empty = @('No data found'); expired = @('Session expired', 'Please log in') }
$k = BrowserAssertPage-Classify -Text "Transfer status`nRecords: 2`nABC123 ok" -Fingerprint $fp
Assert-True ($k['kind'] -eq 'ok' -and @($k['matched']).Count -eq 2) 'assert_page: ok needs ALL ok strings'
$k = BrowserAssertPage-Classify -Text "Transfer status`nLoading..." -Fingerprint $fp
Assert-Equal 'loading' $k['kind'] 'assert_page: loading'
$k = BrowserAssertPage-Classify -Text "Transfer status`nNo data found" -Fingerprint $fp
Assert-Equal 'empty' $k['kind'] 'assert_page: empty'
$k = BrowserAssertPage-Classify -Text "Please log in" -Fingerprint $fp
Assert-Equal 'expired' $k['kind'] 'assert_page: expired (any one string)'
$k = BrowserAssertPage-Classify -Text "Transfer status only" -Fingerprint $fp
Assert-Equal 'unknown' $k['kind'] 'assert_page: only some ok strings -> unknown'
$k = BrowserAssertPage-Classify -Text "Some other page entirely" -Fingerprint $fp
Assert-Equal 'unknown' $k['kind'] 'assert_page: unknown'
$k = BrowserAssertPage-Classify -Text "   " -Fingerprint $fp
Assert-Equal 'loading' $k['kind'] 'assert_page: blank text is loading'
$k = BrowserAssertPage-Classify -Text "Transfer status Records:" -Fingerprint @{ ok = 'Transfer status' }
Assert-Equal 'ok' $k['kind'] 'assert_page: a single string instead of a list is accepted'
$k = BrowserAssertPage-Classify -Text "Transfer status" -Fingerprint @{ loading = @('x') }
Assert-Equal 'unknown' $k['kind'] 'assert_page: no ok list -> never ok'
$ap = Get-EbiStep -Registry $reg -Use 'browser.assert_page'
$ret = & $ap['Invoke'] @{ text = 'Some other page'; fingerprint = $fp } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'page_unknown') 'assert_page step: unknown FAILS (never continues)'
$ret = & $ap['Invoke'] @{ text = 'Loading...'; fingerprint = $fp } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'page_loading') 'assert_page step: loading fails (transient)'
$ret = & $ap['Invoke'] @{ text = 'Transfer status Records: 0'; fingerprint = $fp } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['kind'] -eq 'ok') 'assert_page step: ok passes'
$nav = Get-EbiStep -Registry $reg -Use 'browser.navigate'
$ctx = New-TestCtx
$ret = & $nav['Invoke'] @{ window = 4242; url = ''; waitMs = 0; verifyChange = $false } $ctx
Assert-True ($ret['ok'] -and -not $ret['navigated'] -and @($ret['warnings']).Count -eq 1 -and $ret['warnings'][0]['code'] -eq 'no_url') 'navigate: empty URL -> ok, navigated=false, warning no_url (no keys sent)'

# --- 5. screen.save name rule + real move -----------------------------------------
Write-Host '  -- screen.save'
$n = ScreenSave-Name -Key 'ABC123' -Ext 'png'
Assert-True ($n['name'] -eq 'ABC123.png' -and -not $n['folded']) 'save name: <key>.png'
$n = ScreenSave-Name -Key 'ABC123' -Tag 'row' -Ext '.PNG'
Assert-Equal 'ABC123__row.PNG' $n['name'] 'save name: <key>__<tag>.ext, leading dot tolerated'
$n = ScreenSave-Name -Key ('J' + $fw0 + '1/X') -Tag 'a b' -Ext ''
Assert-True ($n['name'] -eq 'J01_X__a b.png' -and $n['folded']) 'save name: raw key folded to keySafe, reported'
$src = Join-Path $tmpRoot 'shot.png'
[System.IO.File]::WriteAllBytes($src, [byte[]](1, 2, 3))
$sv = Get-EbiStep -Registry $reg -Use 'screen.save'
$ret = & $sv['Invoke'] @{ source = 'shot.png'; dir = 'capture/before_list'; key = 'ABC123'; tag = 'row'; ext = 'png'; copy = $false } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['name'] -eq 'ABC123__row.png' -and (Test-Path -LiteralPath $ret['path']) -and -not (Test-Path -LiteralPath $src)) 'save: moved into <dir>/<key>__<tag>.png'
[System.IO.File]::WriteAllBytes($src, [byte[]](9, 9))
$ret = & $sv['Invoke'] @{ source = 'shot.png'; dir = 'capture/before_list'; key = 'ABC123'; tag = 'row'; ext = 'png'; copy = $true } (New-TestCtx)
Assert-True ($ret['ok'] -and (Test-Path -LiteralPath $src) -and ([System.IO.File]::ReadAllBytes($ret['path']).Length -eq 2)) 'save: copy keeps the source and replaces the old file'
$ret = & $sv['Invoke'] @{ source = 'gone.png'; dir = 'capture'; key = 'K'; tag = ''; ext = 'png'; copy = $false } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'save: missing source is file_not_found'
$cr = Get-EbiStep -Registry $reg -Use 'screen.crop'
$ret = & $cr['Invoke'] @{ path = 'gone.png'; out = ''; left = 1; top = 0; right = 0; bottom = 0 } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'crop: missing file is file_not_found before any GDI+ call'

# --- 6. file.write_json / file.read_json -----------------------------------------
Write-Host '  -- file.write_json / read_json'
$wj = Get-EbiStep -Registry $reg -Use 'file.write_json'
$rj = Get-EbiStep -Registry $reg -Use 'file.read_json'
$data = @{ key = 'ABC123'; n = 2; ok = $true; list = @(1, 2); name = ('' + [char]0x8EE2 + [char]0x9001) }
$ret = & $wj['Invoke'] @{ path = 'capture/ABC123.meta.json'; data = $data } (New-TestCtx)
Assert-True ($ret['ok'] -and (Test-Path -LiteralPath $ret['path'])) 'write_json: written under WorkDir'
$raw = [System.IO.File]::ReadAllText($ret['path'])
Assert-True ($raw.Contains([string][char]0x8EE2)) 'write_json: Japanese as characters'
$ret = & $rj['Invoke'] @{ path = 'capture/ABC123.meta.json' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['exists'] -and ($ret['data'] -is [hashtable]) -and $ret['data']['n'] -eq 2 -and @($ret['data']['list']).Count -eq 2) 'read_json: round trip as hashtable'
$ret = & $rj['Invoke'] @{ path = 'capture/none.meta.json' } (New-TestCtx)
Assert-True ($ret['ok'] -and -not $ret['exists'] -and $null -eq $ret['data'] -and @($ret['warnings']).Count -eq 1 -and $ret['warnings'][0]['code'] -eq 'not_found') 'read_json: missing file -> ok, data=null, warning'
[System.IO.File]::WriteAllText((Join-Path $tmpRoot 'bad.json'), '{ not json')
$ret = & $rj['Invoke'] @{ path = 'bad.json' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'json_invalid') 'read_json: corrupt file is json_invalid'
$ret = & $wj['Invoke'] @{ path = 'x.json'; data = @{ h = [IntPtr]5 } } (New-TestCtx)
Assert-True ((-not $ret['ok'] -and $ret['failure'] -eq 'not_serializable') -or $ret['ok']) 'write_json: an unserializable value is refused (or serialized as a number by this runtime)'
$ctx = New-TestCtx -DryRun $true
$ret = & $wj['Invoke'] @{ path = 'dry.json'; data = @{ a = 1 } } $ctx
Assert-True ($ret['ok'] -and -not (Test-Path -LiteralPath (Join-Path $tmpRoot 'dry.json'))) 'write_json: dry run writes nothing'

# --- 7. file.find ---------------------------------------------------------------------
Write-Host '  -- file.find'
Assert-Equal 'exact' (FileFind-Tier -Name 'ABC123.dat' -Key 'ABC123') 'tier: exact'
Assert-Equal 'stripped' (FileFind-Tier -Name 'ABC123.260824.10515511.dat' -Key 'ABC123') 'tier: file carries a stamp the key lacks'
Assert-Equal 'stripped' (FileFind-Tier -Name 'ABC123.dat' -Key 'ABC123.260824.10515511') 'tier: key carries a stamp the file lacks'
Assert-Equal 'fullwidth' (FileFind-Tier -Name ($fwA + 'BC123.dat') -Key 'ABC123') 'tier: full-width folded'
Assert-Equal 'case' (FileFind-Tier -Name 'abc123.dat' -Key 'ABC123') 'tier: case folded'
Assert-Equal '' (FileFind-Tier -Name 'ABC124.dat' -Key 'ABC123') 'tier: no match'
Assert-Equal 'glob' (FileFind-Tier -Name 'anything.dat' -Key '') 'tier: no key = glob'
$rk = FileFind-Rank -Names @('ABC123.260824.10515511.dat', 'ABC123.dat', 'ABC123.txt') -Key 'ABC123' -Ext 'dat'
Assert-True ($rk['tier'] -eq 'exact' -and @($rk['names']).Count -eq 1) 'rank: exact beats stamped; ext filter drops .txt'
$rk = FileFind-Rank -Names @('ABC123.260824.10515511.dat', 'ABC123.260824.10515533.dat') -Key 'ABC123' -Ext '.dat'
Assert-True ($rk['tier'] -eq 'stripped' -and @($rk['names']).Count -eq 2) 'rank: two stamped reruns both kept'
$rk = FileFind-Rank -Names @() -Key 'X'
Assert-Equal '' $rk['tier'] 'rank: nothing'
$cands = FileFind-Candidates -Files @(@{ name = 'a'; path = 'p'; modifiedAt = '2026-09-28 09:53:40'; size = 12 }, @{ name = 'b'; path = 'q'; modifiedAt = '2026-09-28 09:51:02'; size = 12 }) -Dir 'D' -Tier 'stripped'
Assert-True (@($cands['candidates']).Count -eq 2 -and $cands['candidates'][0]['id'] -eq 'c1' -and $cands['candidates'][1]['candidate'] -eq 'b' -and $cands['suggestion']['id'] -eq 'c1' -and $cands['candidates'][0]['evidence']['source'] -eq 'D') 'candidates: P0-R4 shape with ids, evidence, suggestion'

$dl = Join-Path $tmpRoot 'downloads'
New-Item -ItemType Directory -Path $dl -Force | Out-Null
foreach ($f in @('ABC123.260824.10515511.dat', 'ABC123.260824.10515533.dat', 'DEF456.dat', ($fwA + 'BC789.dat'), 'GHI000.txt')) { [System.IO.File]::WriteAllBytes((Join-Path $dl $f), [byte[]](1)) }
[System.IO.File]::SetLastWriteTimeUtc((Join-Path $dl 'ABC123.260824.10515533.dat'), (Get-Date).ToUniversalTime())
[System.IO.File]::SetLastWriteTimeUtc((Join-Path $dl 'ABC123.260824.10515511.dat'), (Get-Date).ToUniversalTime().AddMinutes(-3))
$ff = Get-EbiStep -Registry $reg -Use 'file.find'
function Find { param([hashtable]$With) $w = @{ dir = 'downloads'; key = ''; ext = ''; glob = '*'; recurse = $false; expect = 'one' }; foreach ($k in $With.Keys) { $w[$k] = $With[$k] }; return (& $ff['Invoke'] $w (New-TestCtx)) }
$ret = Find @{ key = 'DEF456'; ext = 'dat' }
Assert-True ($ret['ok'] -and $ret['matchedBy'] -eq 'exact' -and $ret['found'] -eq 1 -and $ret['path'].EndsWith('DEF456.dat') -and $null -eq $ret['candidates']) 'find: exact'
$ret = Find @{ key = 'ABC123'; ext = 'dat' }
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'ambiguous' -and $ret['found'] -eq 2 -and @($ret['candidates']['candidates']).Count -eq 2) 'find: two stamped reruns -> ambiguous with candidates'
Assert-True ($ret['candidates']['candidates'][0]['candidate'] -eq 'ABC123.260824.10515533.dat' -and $ret['candidates']['suggestion']['id'] -eq 'c1') 'find: newest is c1 and the suggestion; the step still does not pick'
$ret = Find @{ key = 'ABC123'; ext = 'dat'; expect = 'any' }
Assert-True ($ret['ok'] -and $ret['found'] -eq 2 -and $ret['path'].EndsWith('10515533.dat') -and @($ret['files']).Count -eq 2) 'find: expect=any returns all, newest first'
$ret = Find @{ key = 'ABC123.260824.10515511'; ext = 'dat' }
Assert-True ($ret['ok'] -and $ret['matchedBy'] -eq 'exact') 'find: the stamped key hits its own file exactly'
$ret = Find @{ key = 'ABC789'; ext = 'dat' }
Assert-True ($ret['ok'] -and $ret['matchedBy'] -eq 'fullwidth' -and @($ret['warnings']).Count -eq 1 -and $ret['warnings'][0]['code'] -eq 'full_width_name') 'find: full-width file name found, warned'
$ret = Find @{ key = 'ZZZ'; ext = 'dat' }
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'find: nothing -> file_not_found'
$ret = Find @{ key = 'GHI000'; ext = 'dat' }
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'find: ext filter excludes the .txt'
$ret = Find @{ glob = '*.txt' }
Assert-True ($ret['ok'] -and $ret['matchedBy'] -eq 'glob' -and $ret['found'] -eq 1) 'find: glob only'
$ret = Find @{ dir = 'nowhere' }
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'dir_not_found') 'find: missing dir'

# --- 8. file.assert_exists ---------------------------------------------------------
Write-Host '  -- file.assert_exists'
$ae = Get-EbiStep -Registry $reg -Use 'file.assert_exists'
$ret = & $ae['Invoke'] @{ path = 'downloads/DEF456.dat'; kind = 'file' } (New-TestCtx)
Assert-True ($ret['ok'] -and $ret['exists']) 'assert_exists: file'
$ret = & $ae['Invoke'] @{ path = 'downloads'; kind = 'file' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'assert_exists: a dir is not a file'
$ret = & $ae['Invoke'] @{ path = 'downloads'; kind = 'dir' } (New-TestCtx)
Assert-True $ret['ok'] 'assert_exists: dir'
$ret = & $ae['Invoke'] @{ path = 'downloads/none.dat'; kind = 'any' } (New-TestCtx)
Assert-True (-not $ret['ok'] -and $ret['failure'] -eq 'file_not_found') 'assert_exists: missing fails'
$ctx = New-TestCtx -DryRun $true
$ret = & $ae['Invoke'] @{ path = 'downloads/none.dat'; kind = 'any' } $ctx
Assert-True ($ret['ok'] -and -not $ret['exists'] -and @($ret['warnings']).Count -eq 1 -and ($ctx['Log'].Lines[0] -like 'warn:would fail*')) 'assert_exists: dry run warns instead of failing'

# --- 8b. P1-27: no step compares keys on its own ---------------------------------
Write-Host '  -- P1-27 guard'
$offenders = New-Object System.Collections.ArrayList
foreach ($f in @(Get-ChildItem -LiteralPath (Join-Path $repoRoot 'modules') -Filter '*.ps1' -File -Recurse | Where-Object { $_.Name -match '^[a-z]+\.[a-z_]+\.ps1$' })) {
    $n = 0
    foreach ($line in (Get-Content -LiteralPath $f.FullName)) {
        $n++
        $code = ($line -replace '#.*$', '')
        # an emptiness test ($key -eq '' / $null) is not a comparison of two keys
        if ($code -match '\$(key|Key|keySafe|correl|CorrelId)\b\s*-(c?i?eq|c?i?ne)\s+(?!''''|""|\$null)' -or $code -match '(?<!''''|""|\$null)\s+-(c?i?eq|c?i?ne)\s*\$(key|Key|keySafe|correl|CorrelId)\b') { [void]$offenders.Add($f.Name + ':' + $n) }
    }
}
Assert-Equal 0 $offenders.Count ('no step compares a key with -eq / -ne of its own (kernel/Key.ps1 is the one rule set)' + $(if ($offenders.Count) { ': ' + ($offenders.ToArray() -join ', ') } else { '' }))

# --- 9. the four Invoke-CropPng copies are gone ---------------------------------
Write-Host '  -- P1-20 retirement'
$copies = 0
foreach ($f in @(Get-ChildItem -LiteralPath $repoRoot -Filter '*.ps1' -File)) { if ((Get-Content -LiteralPath $f.FullName -Raw) -match '(?m)^function Invoke-CropPng') { $copies++ } }
Assert-Equal 0 $copies 'no `function Invoke-CropPng` left in the repo root'
foreach ($f in @('HmSnap.ps1', 'MqSnap.ps1', 'JenkinsSnap.ps1', 'Crop-Snap.ps1')) {
    $t = Get-Content -LiteralPath (Join-Path $repoRoot $f) -Raw
    Assert-True ($t -match 'kernel/Image\.ps1' -and $t -match 'Invoke-EbiCropPng') ($f + ' dot-sources kernel/Image.ps1 and calls Invoke-EbiCropPng')
}

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
