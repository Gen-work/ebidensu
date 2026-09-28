# Test-GrammarTune.ps1 -- P2-02 ebi grammar tune: the view, the edits, the
# save (grammar + fixture + expected.json, behind the mask gate) and the
# loop driven by a scripted reader. ASCII source; no param() block.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'
. (Join-Path $PSScriptRoot '_TestCommon.ps1')
$repoRoot = Split-Path -Parent $PSScriptRoot
. (Join-Path $repoRoot 'kernel/GrammarTune.ps1')

Reset-Tests 'GrammarTune'
$tmpRoot = Join-Path ([System.IO.Path]::GetTempPath()) ('ebi-tune-' + [Guid]::NewGuid().ToString('N').Substring(0, 8))
New-Item -ItemType Directory -Path $tmpRoot -Force | Out-Null
$tab = "`t"
$text = "Header line`r`n100234${tab}ABC123${tab}OK${tab}2026/08/24 9:50:03`r`n100235${tab}ABC124${tab}OK${tab}2026/08/24 10:02:11`r`nTotal 2`r`n"

Write-Host '  -- view'
$g = @{ parser = 'delimited'; delimiter = "`t"; rowWhen = @{ field = 0; matches = '^\d+$' }; fields = @('jobNo', 'key', 'status', 'time') }
$parsed = ConvertFrom-EbiGrammar -Text $text -Grammar $g
$view = @(Format-EbiTuneView -Grammar $g -Parsed $parsed -Rows 5)
Assert-True ($view[0].StartsWith('grammar: delimited, delimiter=') -and $view[0].Contains('rowWhen=field[0] matches ^\d+$')) 'view: grammar summary line'
Assert-True (($view -join "`n") -like '*records: 2 shown of 2;  unrecognised: 2*') 'view: counts'
Assert-True (($view -join "`n") -like '*| jobNo  | key    | status | time*') 'view: header row with the grammar fields'
Assert-True (($view -join "`n") -like '*unrecognised lines (every one*   1: Header line*   4: Total 2*') 'view: EVERY unrecognised line is listed with its number'
$bad = ConvertFrom-EbiGrammar -Text $text -Grammar @{ parser = 'regex'; pattern = '(' }
Assert-True ((@(Format-EbiTuneView -Grammar @{ parser = 'regex'; pattern = '(' } -Parsed $bad) -join "`n") -like '*! bad pattern*') 'view: a grammar error is shown, not thrown'

Write-Host '  -- edits'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'i ^Total '
Assert-True ($e['ok'] -and @($e['grammar']['ignore']).Count -eq 1 -and -not $g.Contains('ignore')) 'edit i: adds an ignore pattern on a copy'
$e = Edit-EbiTuneGrammar -Grammar $e['grammar'] -Command 'i ^Header'
Assert-True ($e['ok'] -and @($e['grammar']['ignore']).Count -eq 2) 'edit i: appends'
$p2 = ConvertFrom-EbiGrammar -Text $text -Grammar $e['grammar']
Assert-Equal 0 @($p2['unrecognized']).Count 'after two ignores nothing is unrecognised'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'c a,b,c,d'
Assert-True ($e['ok'] -and $e['grammar']['fields'][3] -eq 'd') 'edit c: field names'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'r key ^ABC'
Assert-True ($e['ok'] -and $e['grammar']['rowWhen']['field'] -eq 'key' -and $e['grammar']['rowWhen']['matches'] -eq '^ABC') 'edit r: row rule by field name'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'd ws'
Assert-Equal 'ws' $e['grammar']['delimiter'] 'edit d: delimiter'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'p regex'
Assert-Equal 'regex' $e['grammar']['parser'] 'edit p: parser'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'p csv'
Assert-True (-not $e['ok']) 'edit p: unknown parser refused'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'x ^(?<key>\S+)'
Assert-Equal '^(?<key>\S+)' $e['grammar']['pattern'] 'edit x: pattern'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'l key'
Assert-Equal 'key' $e['grammar']['lastNonEmpty'] 'edit l: lastNonEmpty'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'h Name & Time'
Assert-True ($e['ok'] -and @($e['grammar']['headerLine']['contains']).Count -eq 2) 'edit h: header contains'
$e = Edit-EbiTuneGrammar -Grammar $g -Command 'zzz'
Assert-True (-not $e['ok']) 'edit: unknown letter refused'

Write-Host '  -- save + loop'
$prof = Join-Path $tmpRoot 'prof'
New-Item -ItemType Directory -Path $prof -Force | Out-Null
$s = Save-EbiTuneResult -ProfileDir $prof -Page 'listPage' -Grammar $g -Text $text -FixtureName 'sample' -RecordCount 2
Assert-True ($s['ok'] -and (Test-Path -LiteralPath (Join-Path $prof 'grammar.json')) -and (Test-Path -LiteralPath (Join-Path $prof 'fixtures/listPage/sample.txt'))) 'save: grammar.json + fixture written'
$gj = (Read-EbiJson -Path (Join-Path $prof 'grammar.json'))['value']
Assert-True ($gj.Contains('listPage') -and $gj['listPage']['parser'] -eq 'delimited') 'save: grammar under the page name'
$ej = (Read-EbiJson -Path (Join-Path $prof 'fixtures/listPage/expected.json'))['value']
Assert-True ($ej.Contains('sample.txt') -and $ej['sample.txt']['records'] -eq 2) 'save: expected.json entry with the record count'
$s = Save-EbiTuneResult -ProfileDir $prof -Page 'listPage' -Grammar $g -Text ($text + "mail root@corp.example.com`r`n") -FixtureName 'leak' -RecordCount 2
Assert-True (-not $s['ok'] -and $s['message'] -like '*sensitive*' -and -not (Test-Path -LiteralPath (Join-Path $prof 'fixtures/listPage/leak.txt'))) 'save: the mask gate refuses a text with a mail address; nothing written'
$script:answers = New-Object System.Collections.ArrayList
foreach ($a in @('i ^Header', 'i ^Total ', 'u', 'bogus', 's looped', 'q')) { [void]$script:answers.Add($a) }
$reader = { if ($script:answers.Count -eq 0) { return 'q' }; $a = $script:answers[0]; $script:answers.RemoveAt(0); return $a }
$r = Invoke-EbiGrammarTune -Text $text -ProfileDir $prof -Page 'listPage' -Grammar $g -Reader $reader
Assert-True ($r['ok'] -and $r['quit'] -and $r['saved'] -and @($r['grammar']['ignore']).Count -eq 2) 'loop: edits apply, u and a bad command do not break it, s saves, q quits'
Assert-True (Test-Path -LiteralPath (Join-Path $prof 'fixtures/listPage/looped.txt')) 'loop: the fixture named after s'
$gj = (Read-EbiJson -Path (Join-Path $prof 'grammar.json'))['value']
Assert-Equal 2 @($gj['listPage']['ignore']).Count 'loop: grammar.json holds the tuned grammar'

Remove-Item -LiteralPath $tmpRoot -Recurse -Force -ErrorAction SilentlyContinue
$failed = Complete-Tests
exit $failed
