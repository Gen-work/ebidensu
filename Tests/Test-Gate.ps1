#Requires -Version 5.1
# Test-Gate.ps1 -- kernel/Gate.ps1 (P1-05): the one ASCII gate panel.
# The card's criteria: no line wraps in an 80-column console; r / s / q / m
# all work. Reading is injected, so nothing here blocks.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Gate.ps1')

Reset-Tests 'Gate'

# ---------------------------------------------------------------- wrapping
Assert-Equal 'abc' (@(ConvertTo-EbiGateWrapped -Text 'abc' -Width 10) -join '|') 'wrap: short text is one line'
Assert-Equal 'aaa bbb|ccc' (@(ConvertTo-EbiGateWrapped -Text 'aaa bbb ccc' -Width 8) -join '|') 'wrap: breaks at a space'
Assert-Equal 'abcde|fghij|k' (@(ConvertTo-EbiGateWrapped -Text 'abcdefghijk' -Width 5) -join '|') 'wrap: hard-breaks a word longer than the width'
Assert-Equal 'a|b' (@(ConvertTo-EbiGateWrapped -Text "a`nb" -Width 5) -join '|') 'wrap: existing line breaks are kept'
Assert-Equal 1 @(ConvertTo-EbiGateWrapped -Text '' -Width 5).Count 'wrap: empty text is one empty line'

# ---------------------------------------------------------------- the panel
$actions = @(@{ key = 'r'; label = 'retry' }, @{ key = 's'; label = 'skip' }, @{ key = 'q'; label = 'quit' }, @{ key = 'm'; label = 'note' })
$long = ('x' * 200) + ' ' + ('word ' * 40)
$panel = @(Format-EbiGatePanel -Title 'FAILED' -What @('each[A1 / JOB_A]/wait (browser.wait_for)', 'timeout: no match after 12s', $long) -Next @('r: run the step again', 's: leave the item pending') -Evidence @('capture/before_ts/A1.png', 'capture/before_ts/A1.txt') -Actions $actions)
Assert-True ($panel.Count -gt 8) 'panel: has lines'
$tooWide = @($panel | Where-Object { $_.Length -gt 80 })
Assert-Equal 0 $tooWide.Count 'panel: no line is wider than 80 columns (even with a 200-char word inside)'
$narrow = @($panel | Where-Object { $_.Length -ne 80 })
Assert-Equal 0 $narrow.Count 'panel: every line is exactly 80 wide (a closed box)'
Assert-True ($panel[0].StartsWith('+ FAILED ')) 'panel: the title sits in the top border'
Assert-True (($panel -join "`n").Contains('| WHAT HAPPENED')) 'panel: WHAT HAPPENED section'
Assert-True (($panel -join "`n").Contains('| NEXT')) 'panel: NEXT section'
Assert-True (($panel -join "`n").Contains('| EVIDENCE')) 'panel: EVIDENCE section'
Assert-True (($panel -join "`n").Contains('r=retry   s=skip   q=quit   m=note')) 'panel: the actions line lists key=label'
Assert-True (($panel -join "`n").Contains('capture/before_ts/A1.png')) 'panel: evidence paths are shown'
$noEv = @(Format-EbiGatePanel -Title 'T' -What 'w' -Actions $actions)
Assert-True (-not (($noEv -join "`n").Contains('EVIDENCE'))) 'panel: an empty section is left out'
Assert-True (-not (($noEv -join "`n").Contains('| NEXT'))) 'panel: ... NEXT too'
$narrowPanel = @(Format-EbiGatePanel -Title 'T' -What ('y' * 100) -Actions $actions -Width 40)
Assert-Equal 0 @($narrowPanel | Where-Object { $_.Length -ne 40 }).Count 'panel: an explicit width is honoured'
$longTitle = @(Format-EbiGatePanel -Title ('t' * 120) -What 'w' -Actions @())
Assert-Equal 80 $longTitle[0].Length 'panel: an over-long title is cut, the border stays 80'

# ---------------------------------------------------------------- answers
$a = Read-EbiGateAnswer -Text 'r' -Actions $actions
Assert-True ($a['ok'] -and $a['action'] -eq 'r') 'answer: r'
Assert-True ((Read-EbiGateAnswer -Text ' S ' -Actions $actions)['action'] -eq 's') 'answer: case and whitespace do not matter'
$a = Read-EbiGateAnswer -Text 'm the page was blank' -Actions $actions
Assert-True ($a['ok'] -and $a['action'] -eq 'm' -and $a['note'] -eq 'the page was blank') 'answer: m carries the note'
Assert-True (-not (Read-EbiGateAnswer -Text 'x' -Actions $actions)['ok']) 'answer: an unoffered key is not ok'
Assert-True (-not (Read-EbiGateAnswer -Text 's oops' -Actions $actions)['ok']) 'answer: a note after a non-m key is a typo, not a skip'
Assert-True (-not (Read-EbiGateAnswer -Text '' -Actions $actions)['ok']) 'answer: Enter with no default is not an answer'
$a = Read-EbiGateAnswer -Text '' -Actions $actions -Default 'r'
Assert-True ($a['ok'] -and $a['action'] -eq 'r') 'answer: Enter takes the default'
Assert-True (-not (Read-EbiGateAnswer -Text 'm' -Actions @(@{ key = 'r'; label = 'x' }))['ok']) 'answer: m only when offered'
$a = Read-EbiGateAnswer -Text '2' -Actions @('1=first', '2=second')
Assert-True ($a['ok'] -and $a['action'] -eq '2') 'answer: numbered choices given as "key=label" strings'
Assert-True (-not (Read-EbiGateAnswer -Text 'm' -Actions $actions)['ok'] -or (Read-EbiGateAnswer -Text 'm' -Actions $actions)['note'] -eq '') 'answer: a bare m is a note with no text'

# ---------------------------------------------------------------- Show-EbiGate with an injected reader
$script:typed = New-Object System.Collections.ArrayList
function Set-Typed { param($List) $script:typed.Clear(); foreach ($t in @($List)) { [void]$script:typed.Add($t) } }
$reader = { if ($script:typed.Count -eq 0) { return '' }; $t = $script:typed[0]; $script:typed.RemoveAt(0); return $t }
Set-Typed @('zz', 'm  looks wrong ')
$r = Show-EbiGate -Title 'T' -What 'w' -Actions $actions -Reader $reader
Assert-True ($r['action'] -eq 'm' -and $r['note'] -eq 'looks wrong' -and -not $r['auto']) 'show: an invalid answer is asked again; m returns the note'
Set-Typed @('q')
Assert-Equal 'q' (Show-EbiGate -Title 'T' -What 'w' -Actions $actions -Reader $reader)['action'] 'show: q'
Set-Typed @('')
Assert-Equal 's' (Show-EbiGate -Title 'T' -What 'w' -Actions $actions -Default 's' -Reader $reader)['action'] 'show: Enter takes the default'
$r = Show-EbiGate -Title 'T' -What 'w' -Actions $actions -Auto 's' -DryRun $true
Assert-True ($r['action'] -eq 's' -and $r['auto']) 'show: under DryRun the Auto action is taken without reading'
Set-Typed @('', '', '')
$r = Show-EbiGate -Title 'T' -What 'w' -Actions $actions -Auto 'q' -Reader $reader
Assert-True ($r['action'] -eq 'q' -and $r['auto']) 'show: with no valid answer ever, the Auto action is taken instead of looping forever'

# ---------------------------------------------------------------- the runner handler
Set-Typed @('r')
Assert-Equal 'r' (Invoke-EbiGateAsk -Question @{ kind = 'error'; section = 'each'; id = 'wait'; use = 'browser.wait_for'; key = 'A1'; group = ''; failure = 'timeout'; message = 'no match'; attempt = 2; transient = $true; policy = 'ask'; evidence = @{ path = 'a.png'; rect = @{ X = 1 } } } -Reader $reader) 'ask: error -> r'
Set-Typed @('n')
Assert-Equal 'n' (Invoke-EbiGateAsk -Question @{ kind = 'confirm'; section = 'each'; id = 'rep'; use = 'excel.replace_sheet'; key = 'A1'; group = ''; effects = 'destructive'; with = @{ sheet = 'x' } } -Reader $reader) 'ask: confirm -> n'
Assert-Equal 'y' (Invoke-EbiGateAsk -Question @{ kind = 'confirm'; section = 'each'; id = 'rep'; use = 'x.y'; key = ''; group = ''; effects = 'destructive'; with = @{} } -DryRun $true) 'ask: confirm under DryRun is y'
Assert-Equal 's' (Invoke-EbiGateAsk -Question @{ kind = 'error'; section = 'setup'; id = 'e'; use = 'x.y'; key = ''; group = ''; failure = 'boom'; message = 'm'; attempt = 1; transient = $false; policy = 'ask' } -DryRun $true) 'ask: error under DryRun is s'

$rc = Complete-Tests
exit $rc
