# modules/human/human.pick.ps1
# Let the operator pick today's work from named sets (10/20, DETOUR-GFIX-
# 001.10). Two levels so the usual answer is one key:
#     first   a few recommended combinations (presets): Enter = 1
#     more    every set, numbered, with a note per name (m from the first)
# On both levels a name, or a part of one, can be typed instead of a
# number ("DBC" = every offered name holding DBC), "-NAME" drops one, and
# "h NAME ..." puts names on hold instead (they wait for someone else).
# Whatever is typed is read back ("-> 4 jobs: ...") and Enter confirms it,
# so a typo never runs anything.

. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'human.pick'
  group      = 'human'
  summary    = 'Pick names from recommended combinations or numbered sets, or by typing (part of) names'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    sets      = @{ type='list';   required=$true; desc='@{ n (0 = shown, not pickable); id; label; names; notes (name -> text); default } (table.day_sets)' }
    presets   = @{ type='list';   default=@(); desc='@{ label; names; hint } shown first, numbered from 1; empty = the default sets as one preset' }
    title     = @{ type='string'; default='PICK' }
    info      = @{ type='list';   default=@(); desc='lines shown above the choices' }
    noun      = @{ type='string'; default='job'; desc='what a name is, in the read-back ("4 jobs")' }
    holdReason = @{ type='string'; default='waiting'; desc='what Enter means when asked why names go on hold' }
    answer     = @{ type='list';   default=@(); desc='names given before the panel (ebi.ps1 run <plan> -Only A,B -> {{run.only}}): taken as the typed answer, read back and confirmed; a name no set offers shows the panel instead' }
    autoPick   = @{ type='bool';   default=$false; desc='true: take choice 1 without asking (the caller already knows it is right, e.g. verify.time_candidates certain); the panel is not shown' }
  }
  outputs    = @{
    picked     = @{ type='list';   desc='the picked names, in set order' }
    pickedCount = @{ type='int';   desc='how many were picked' }
    any         = @{ type='bool';   desc='at least one picked (a later step''s when: any == true is false too when this step was skipped)' }
    hold       = @{ type='list';   desc='names the operator put on hold (h ...)' }
    holdReason = @{ type='string' }
    answer     = @{ type='string'; desc='the confirmed answer as typed' }
    auto       = @{ type='bool';   desc='nobody could answer (dry run / no console): preset 1 was taken' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'human.pick'; with = @{ title = 'PLAN today'; sets = '{{steps.sets.out.sets}}'; presets = '{{steps.sets.out.presets}}'; info = '{{steps.sets.out.info}}' } }
  notes      = 'Runs in setup, once per run. Under DryRun or without a console the first preset is taken (auto=true). q stops the whole run (operator_quit); 0 picks nothing.'
}

function HumanPick-Norm {
    param($Value)
    if ($null -eq $Value) { return '' }
    return (ConvertTo-EbiHalfWidth -Value ([string]$Value)).Trim().ToUpperInvariant()
}

function HumanPick-Presets {
    # PURE. The presets to show: the given ones, else one made of the default sets.
    param($Sets, $Presets)
    $list = @(@($Presets) | Where-Object { $_ -is [System.Collections.IDictionary] })
    if ($list.Count -gt 0) { return $list }
    $keys = New-Object System.Collections.ArrayList
    foreach ($s in @($Sets)) { if ($s -is [System.Collections.IDictionary] -and [bool]$s['default']) { foreach ($k in @($s['names'])) { [void]$keys.Add([string]$k) } } }
    return @(@{ label = 'the usual'; names = $keys.ToArray(); hint = '' })
}

function HumanPick-Offered {
    # PURE. Every pickable name, in set order, once.
    param($Sets)
    $out = New-Object System.Collections.ArrayList
    $seen = @{}
    foreach ($s in @($Sets)) {
        if (-not ($s -is [System.Collections.IDictionary]) -or [int]$s['n'] -le 0) { continue }
        foreach ($k in @($s['names'])) { $n = HumanPick-Norm $k; if ($n -ne '' -and -not $seen.Contains($n)) { $seen[$n] = $true; [void]$out.Add([string]$k) } }
    }
    return $out.ToArray()
}

function HumanPick-Resolve {
    <#
      PURE. One typed line -> what it means.
        Level   'first' (numbers = presets) or 'more' (numbers = set n)
      Returns @{ ok; kind = pick | hold | more | back | none | quit;
      keys; message }. A pick's keys keep the offered order.
    #>
    param([string]$Text, $Sets, $Presets, [string]$Level = 'first')
    $t = (ConvertTo-EbiHalfWidth -Value $(if ($null -eq $Text) { '' } else { $Text })).Trim()
    $empty = @{ ok = $true; kind = 'pick'; keys = @(); message = '' }
    $presetList = @(HumanPick-Presets -Sets $Sets -Presets $Presets)
    if ($t -eq '') {
        if ($Level -eq 'first') { $t = '1' } else { return @{ ok = $false; kind = ''; keys = @(); message = 'type set numbers or names (b = back)' } }
    }
    $low = $t.ToLowerInvariant()
    if ($low -eq 'q') { return @{ ok = $true; kind = 'quit'; keys = @(); message = '' } }
    if ($low -eq 'm') { return @{ ok = $true; kind = 'more'; keys = @(); message = '' } }
    if ($low -eq 'b') { return @{ ok = $true; kind = 'back'; keys = @(); message = '' } }
    if ($low -eq '0' -or $low -eq 'n') { return @{ ok = $true; kind = 'none'; keys = @(); message = '' } }
    $tokens = @($t -split '[\s,;]+' | Where-Object { $_ -ne '' })
    $kind = 'pick'
    if ($tokens.Count -gt 0 -and $tokens[0].ToLowerInvariant() -eq 'h') { $kind = 'hold'; $tokens = @($tokens | Select-Object -Skip 1) }
    if ($kind -eq 'hold' -and $tokens.Count -eq 0) { return @{ ok = $false; kind = ''; keys = @(); message = 'h needs names: h KJODWDBC, or h DBC for every name with DBC' } }
    $offered = @(HumanPick-Offered -Sets $Sets)
    $chosen = @{}
    $bad = New-Object System.Collections.ArrayList
    foreach ($tok in $tokens) {
        $drop = $false
        $w = $tok
        if ($w.StartsWith('-') -and $w.Length -gt 1) { $drop = $true; $w = $w.Substring(1) }
        $hits = New-Object System.Collections.ArrayList
        # a number is a choice; one with a leading zero, or past the last
        # choice, is a part of a name (file names are digits: F202609110026)
        $maxNum = if ($Level -eq 'first') { $presetList.Count } else { @(@($Sets) | ForEach-Object { [int]$_['n'] } | Measure-Object -Maximum).Maximum }
        if ($w -match '^[1-9]\d*$' -and $kind -eq 'pick' -and [int64]$w -le [int64]$maxNum) {
            $num = [int]$w
            if ($Level -eq 'first') {
                if ($num -lt 1 -or $num -gt $presetList.Count) { [void]$bad.Add(($w + ' (no such choice)')); continue }
                foreach ($k in @($presetList[$num - 1]['names'])) { [void]$hits.Add([string]$k) }
            } else {
                $set = $null
                foreach ($s in @($Sets)) { if ($s -is [System.Collections.IDictionary] -and [int]$s['n'] -eq $num) { $set = $s } }
                if ($null -eq $set) { [void]$bad.Add(($w + ' (no such set)')); continue }
                foreach ($k in @($set['names'])) { [void]$hits.Add([string]$k) }
            }
        } else {
            $nw = HumanPick-Norm $w
            foreach ($k in $offered) { if ((HumanPick-Norm $k) -eq $nw) { [void]$hits.Add($k) } }
            if ($hits.Count -eq 0) { foreach ($k in $offered) { if ((HumanPick-Norm $k).Contains($nw)) { [void]$hits.Add($k) } } }
            if ($hits.Count -eq 0) { [void]$bad.Add(($w + ' (not in any set)')); continue }
        }
        foreach ($k in $hits) { $n = HumanPick-Norm $k; if ($drop) { [void]$chosen.Remove($n) } else { $chosen[$n] = $true } }
    }
    if ($bad.Count -gt 0) { return @{ ok = $false; kind = ''; keys = @(); message = ('not understood: ' + ($bad.ToArray() -join ', ')) } }
    $keys = @(foreach ($k in $offered) { if ($chosen.Contains((HumanPick-Norm $k))) { $k } })
    # a preset may hold a name no set offers (should not happen; keep it)
    foreach ($n in $chosen.Keys) { if (-not (@($keys | ForEach-Object { HumanPick-Norm $_ }) -contains $n)) { $keys += $n } }
    return @{ ok = $true; kind = $kind; keys = $keys; message = '' }
}

function HumanPick-KeyText {
    # PURE. Names as one line, with a note after each when asked: A B(note) C.
    param($Keys, $Notes = $null, [int]$Max = 0)
    $parts = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($k in @($Keys)) {
        $i++
        if ($Max -gt 0 -and $i -gt $Max) { [void]$parts.Add(('+' + (@($Keys).Count - $Max) + ' more')); break }
        $n = if ($null -ne $Notes -and ($Notes -is [System.Collections.IDictionary]) -and $Notes.Contains([string]$k)) { [string]$Notes[[string]$k] } else { '' }
        [void]$parts.Add($(if ($n -ne '') { [string]$k + ' (' + $n + ')' } else { [string]$k }))
    }
    return ($parts.ToArray() -join '  ')
}

function HumanPick-Lines {
    # PURE. The choice lines of one level (what the panel's WHAT HAPPENED shows).
    param($Sets, $Presets, [string]$Level, [string]$Noun = 'job')
    $L = New-Object System.Collections.ArrayList
    if ($Level -eq 'first') {
        $i = 0
        foreach ($p in @(HumanPick-Presets -Sets $Sets -Presets $Presets)) {
            $i++
            $c = @($p['names']).Count
            $head = ('' + $i + '  ' + [string]$p['label'] + $(if ($i -eq 1) { '  <- Enter' } else { '' }))
            [void]$L.Add($head)
            $body = if ($c -eq 0) { '(nothing)' } else { ('' + $c + ' ' + $Noun + $(if ($c -ne 1) { 's' } else { '' }) + ': ' + (HumanPick-KeyText -Keys $p['names'] -Max 12)) }
            [void]$L.Add(('   ' + $body + $(if ([string]$p['hint'] -ne '') { '  -- ' + [string]$p['hint'] } else { '' })))
        }
        return $L.ToArray()
    }
    foreach ($s in @($Sets)) {
        if (-not ($s -is [System.Collections.IDictionary])) { continue }
        $num = if ([int]$s['n'] -gt 0) { ([string][int]$s['n']) } else { '-' }
        $c = @($s['names']).Count
        [void]$L.Add(($num + '  ' + [string]$s['label'] + '  (' + $c + ')' + $(if ([int]$s['n'] -le 0) { '  not pickable' } else { '' })))
        if ($c -gt 0) { [void]$L.Add(('   ' + (HumanPick-KeyText -Keys $s['names'] -Notes $s['notes']))) }
    }
    return $L.ToArray()
}

function Invoke-Step {
    param($In, $Ctx)
    $sets = @(@($In['sets']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $presets = @(HumanPick-Presets -Sets $sets -Presets @($In['presets']))
    $noun = [string]$In['noun']
    $reader = if ($null -ne $Ctx['Reader']) { $Ctx['Reader'] } else { $null }
    $level = 'first'
    $info = @(@($In['info']) | Where-Object { [string]$_ -ne '' } | ForEach-Object { [string]$_ })
    $held = @(); $heldReason = ''
    if ([bool]$In['autoPick']) {
        $keys = if ($presets.Count -gt 0) { @($presets[0]['names']) } else { @() }
        $Ctx.Log.Info(('taken without asking: ' + ($keys -join ', ')))
        return @{ ok = $true; picked = $keys; pickedCount = @($keys).Count; any = (@($keys).Count -gt 0); hold = @(); holdReason = ''; answer = ''; auto = $true }
    }
    $again = $null          # an answer typed at the read-back instead of Enter
    $given = @(@($In['answer']) | ForEach-Object { ([string]$_).Trim() } | Where-Object { $_ -ne '' })
    if ($given.Count -gt 0) {
        $pre = HumanPick-Resolve -Text ($given -join ' ') -Sets $sets -Presets $presets -Level 'first'
        if ($pre['ok'] -and [string]$pre['kind'] -eq 'pick' -and @($pre['keys']).Count -gt 0) {
            $noConsole = [bool]$Ctx['DryRun'] -or ([string]$env:EBI_NO_ASK -eq '1' -and $null -eq $reader)
            if (-not $noConsole -and $null -eq $reader) { try { if ([Console]::IsInputRedirected) { $noConsole = $true } } catch { } }
            if ($noConsole) {
                $Ctx.Log.Info(('given (-Only): ' + (@($pre['keys']) -join ', ')))
                return @{ ok = $true; picked = @($pre['keys']); pickedCount = @($pre['keys']).Count; any = $true; hold = @(); holdReason = ''; answer = ($given -join ' '); auto = $true }
            }
            $again = ($given -join ' ')
        } else {
            Write-Host ('  -Only ' + ($given -join ',') + ': ' + $(if ($pre['ok']) { 'nothing to take' } else { [string]$pre['message'] }) + ' -- pick below') -ForegroundColor DarkYellow
        }
    }
    $tries = 0
    while ($true) {
        $tries++
        if ($tries -gt 40) { return @{ ok = $false; failure = 'operator_quit'; message = 'no usable answer after 40 tries'; picked = @(); pickedCount = 0; any = $false; hold = $held; holdReason = $heldReason; answer = ''; auto = $false } }
        if ($null -ne $again) { $typed = $again; $again = $null }
        else {
            $what = @($info) + @(HumanPick-Lines -Sets $sets -Presets $presets -Level $level -Noun $noun)
            $next = if ($level -eq 'first') {
                @('Enter = 1.  A number = that choice.  Or type names, or part of one: DBC = every name with DBC; -NAME drops one (2 -SJODWDBC)', 'm = more sets (late / ahead / others / mapping)   h NAME = put on hold   0 = nothing today   q = stop')
            } else {
                @('Set numbers (1,3), names or parts of names, -NAME drops one', 'b = back to the recommended choices   h NAME = put on hold   0 = nothing today   q = stop')
            }
            $r = Show-EbiGate -Title ([string]$In['title'] + $(if ($level -eq 'more') { ': every set' } else { '' })) -What $what -Next $next -Actions @(@{ key = 'i'; label = 'answer (Enter = 1)' }, @{ key = 'q'; label = 'quit' }) -Default 'i' -Auto 'i' -DryRun ([bool]$Ctx['DryRun']) -Reader $reader -Raw
            if ($r['auto']) {
                $keys = if ($presets.Count -gt 0) { @($presets[0]['names']) } else { @() }
                $Ctx.Log.Info(('no one to ask: took choice 1 (' + @($keys).Count + ' ' + $noun + 's)'))
                return @{ ok = $true; picked = $keys; pickedCount = @($keys).Count; any = (@($keys).Count -gt 0); hold = @(); holdReason = ''; answer = ''; auto = $true }
            }
            $typed = if ([string]$r['action'] -eq 'q') { 'q' } else { [string]$r['note'] }
        }
        $res = HumanPick-Resolve -Text $typed -Sets $sets -Presets $presets -Level $level
        if (-not $res['ok']) { Write-Host ('  ' + $res['message']) -ForegroundColor DarkYellow; continue }
        # (if, not switch: a continue inside switch would only leave the switch)
        if ([string]$res['kind'] -eq 'quit') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q at the pick'; picked = @(); pickedCount = 0; any = $false; hold = $held; holdReason = $heldReason; answer = 'q'; auto = $false } }
        if ([string]$res['kind'] -eq 'more') { $level = 'more'; continue }
        if ([string]$res['kind'] -eq 'back') { $level = 'first'; continue }
        $keys = @($res['keys'])
        $isHold = ([string]$res['kind'] -eq 'hold')
        $c = $keys.Count
        $plural = $(if ($c -ne 1) { 's' } else { '' })
        $readBack = if ([string]$res['kind'] -eq 'none') { '-> nothing today' }
            elseif ($isHold) { ('-> put on hold, ' + $c + ' ' + $noun + $plural + ': ' + ($keys -join ' ')) }
            elseif ($c -eq 0) { '-> nothing (that choice is empty)' }
            else { ('-> ' + $c + ' ' + $noun + $plural + ' today: ' + ($keys -join ' ')) }
        foreach ($w in @(ConvertTo-EbiGateWrapped -Text $readBack -Width 76)) { Write-Host ('  ' + $w) -ForegroundColor Green }
        $reason = ''
        if ($isHold) {
            Write-Host ('  why? (Enter = ' + [string]$In['holdReason'] + ') ') -NoNewline
            $why = if ($null -ne $reader) { [string](& $reader) } else { [string](Read-Host) }
            $reason = if ($why.Trim() -ne '') { $why.Trim() } else { [string]$In['holdReason'] }
        }
        Write-Host '  Enter = yes, or type another answer: ' -ForegroundColor Magenta -NoNewline
        $ok = if ($null -ne $reader) { [string](& $reader) } else { [string](Read-Host) }
        if ($ok.Trim() -ne '') { $again = $ok; continue }
        if ($isHold) {
            # held names leave every set; then the pick is asked again
            $held = @($held) + @($keys); $heldReason = $reason
            $Ctx.Log.Info(('hold: ' + ($keys -join ', ') + ' (' + $reason + ')'))
            $sets = @(foreach ($s in $sets) { $copy = @{}; foreach ($kk in $s.Keys) { $copy[$kk] = $s[$kk] }; $copy['names'] = @(@($s['names']) | Where-Object { -not (@($held) -contains [string]$_) }); $copy })
            $presets = @(foreach ($p in $presets) { $copy = @{}; foreach ($kk in $p.Keys) { $copy[$kk] = $p[$kk] }; $copy['names'] = @(@($p['names']) | Where-Object { -not (@($held) -contains [string]$_) }); $copy })
            $info = @($info) + @(('on hold now: ' + ($keys -join ' ') + ' (' + $reason + ')'))
            continue
        }
        $final = if ([string]$res['kind'] -eq 'none') { @() } else { $keys }
        return @{ ok = $true; picked = $final; pickedCount = @($final).Count; any = (@($final).Count -gt 0); hold = $held; holdReason = $heldReason; answer = $typed; auto = $false }
    }
}
