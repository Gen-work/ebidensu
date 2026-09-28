# modules/human/human.choose.ps1
# Lay out every candidate with all its evidence and let a person pick
# (P1-34, PROFILE-SCHEMA 6.3): the ONE renderer of the P0-R4 candidate
# shape that table.key / file.find / verify.match_record return. Never
# shows only the "best" one; the suggestion and the doubts are printed
# next to the list, not instead of it.

. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'human.choose'
  group      = 'human'
  summary    = 'Show every candidate with its evidence; the operator picks one, none, or skips'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    candidates = @{ type='map';    required=$true; desc='the P0-R4 shape: { candidates: [ { id, candidate, evidence } ], suggestion?, doubts? }' }
    question   = @{ type='string'; default='Which one is it?' }
    key        = @{ type='string'; default=''; desc='which item; empty = the current item' }
  }
  outputs    = @{
    action    = @{ type='string'; desc='chosen | none | skip' }
    id        = @{ type='string'; desc='the chosen candidate id (c1, c2, ...) or empty' }
    index     = @{ type='int';    desc='1-based position of the choice; 0 when none' }
    candidate = @{ type='string'; desc='the chosen candidate text' }
    learn     = @{ type='bool';   desc='the operator wants this decision kept as a rule (persisting it is a later card)' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'human.choose'; with = @{ candidates = '{{steps.find.out.candidates}}'; question = 'Which file is this run''s?' } }
  notes      = 'Under DryRun or without a console the suggestion is taken when there is one, else none. n = none of them (the item becomes unknown), s = skip (leave pending), q = quit.'
}

function HumanChoose-Lines {
    # PURE. The candidate shape -> the WHAT lines: "#1 text  [k=v, k=v]".
    param($Shape)
    $L = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($c in @($Shape['candidates'])) {
        $i++
        $ev = New-Object System.Collections.ArrayList
        if ($c -is [System.Collections.IDictionary] -and $c.Contains('evidence') -and ($c['evidence'] -is [System.Collections.IDictionary])) {
            foreach ($k in ($c['evidence'].Keys | Sort-Object)) { $v = $c['evidence'][$k]; if ($null -ne $v -and [string]$v -ne '') { [void]$ev.Add([string]$k + '=' + [string]$v) } }
        }
        $text = if ($c -is [System.Collections.IDictionary] -and $c.Contains('candidate')) { [string]$c['candidate'] } else { [string]$c }
        [void]$L.Add('#' + $i + ' ' + $text + $(if ($ev.Count) { '   [' + ($ev.ToArray() -join ', ') + ']' } else { '   [no evidence]' }))
    }
    return $L.ToArray()
}

function HumanChoose-SuggestIndex {
    # PURE. 1-based index of the suggestion, 0 when none.
    param($Shape)
    if (-not ($Shape -is [System.Collections.IDictionary]) -or -not $Shape.Contains('suggestion') -or -not ($Shape['suggestion'] -is [System.Collections.IDictionary]) -or -not $Shape['suggestion'].Contains('id')) { return 0 }
    $i = 0
    foreach ($c in @($Shape['candidates'])) { $i++; if ($c -is [System.Collections.IDictionary] -and $c.Contains('id') -and [string]$c['id'] -eq [string]$Shape['suggestion']['id']) { return $i } }
    return 0
}

function Invoke-Step {
    param($In, $Ctx)
    $shape = $In['candidates']
    if ($null -eq $shape -or -not ($shape -is [System.Collections.IDictionary]) -or -not $shape.Contains('candidates')) { return @{ ok = $false; failure = 'input_invalid'; message = 'candidates must be the P0-R4 shape with a "candidates" list'; action = 'none'; id = ''; index = 0; candidate = ''; learn = $false } }
    $list = @($shape['candidates'])
    if ($list.Count -eq 0) { return @{ ok = $true; action = 'none'; id = ''; index = 0; candidate = ''; learn = $false; warnings = @( @{ code = 'no_candidates'; message = 'nothing to choose from'; data = @{} } ) } }
    $key = [string]$In['key']
    if ($key -eq '' -and $null -ne $Ctx['Item']) { $key = Get-EbiKeyDisplay -Item $Ctx['Item'] -KeyColumns @($Ctx['KeyColumns']) }
    $sugg = HumanChoose-SuggestIndex -Shape $shape
    $next = New-Object System.Collections.ArrayList
    if ($sugg -gt 0) { [void]$next.Add('suggestion: #' + $sugg + $(if ($shape['suggestion'].Contains('reason') -and [string]$shape['suggestion']['reason'] -ne '') { ' -- ' + [string]$shape['suggestion']['reason'] } else { '' })) }
    if ($shape.Contains('doubts') -and [string]$shape['doubts'] -ne '') { [void]$next.Add('doubt: ' + [string]$shape['doubts']) }
    [void]$next.Add('1-' + $list.Count + ': pick one   n: none of them (unknown)   s: skip, leave pending   q: quit')
    $actions = New-Object System.Collections.ArrayList
    for ($i = 1; $i -le $list.Count; $i++) { [void]$actions.Add(@{ key = [string]$i; label = ('#' + $i) }) }
    [void]$actions.Add(@{ key = 'n'; label = 'none' }); [void]$actions.Add(@{ key = 's'; label = 'skip' }); [void]$actions.Add(@{ key = 'q'; label = 'quit' })
    $auto = if ($sugg -gt 0) { [string]$sugg } else { 'n' }
    $what = @([string]$In['question']) + @(HumanChoose-Lines -Shape $shape)
    $r = Show-EbiGate -Title ('CHOOSE ' + $key) -What $what -Next $next.ToArray() -Actions $actions.ToArray() -Default '' -Auto $auto -DryRun ([bool]$Ctx['DryRun'])
    $a = [string]$r['action']
    if ($r['auto']) { $Ctx.Log.Info(('nobody to ask (dry run or no console): ' + $(if ($sugg -gt 0) { 'took the suggestion #' + $sugg } else { 'none chosen' }))) }
    if ($a -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q while choosing'; action = 'none'; id = ''; index = 0; candidate = ''; learn = $false } }
    if ($a -eq 'n') { return @{ ok = $true; action = 'none'; id = ''; index = 0; candidate = ''; learn = $false } }
    if ($a -eq 's') { return @{ ok = $true; action = 'skip'; id = ''; index = 0; candidate = ''; learn = $false } }
    $idx = [int]$a
    $c = $list[$idx - 1]
    $learn = $false
    if (-not $r['auto']) {
        $l = Show-EbiGate -Title 'LEARN' -What ('keep "' + [string]$c['candidate'] + '" as a rule for next time?') -Actions @(@{ key = 'y'; label = 'yes' }, @{ key = 'n'; label = 'no (Enter)' }) -Default 'n' -Auto 'n' -DryRun ([bool]$Ctx['DryRun'])
        $learn = ([string]$l['action'] -eq 'y')
    }
    return @{ ok = $true; action = 'chosen'; id = $(if ($c -is [System.Collections.IDictionary] -and $c.Contains('id')) { [string]$c['id'] } else { 'c' + $idx }); index = $idx; candidate = $(if ($c -is [System.Collections.IDictionary] -and $c.Contains('candidate')) { [string]$c['candidate'] } else { [string]$c }); learn = $learn }
}
