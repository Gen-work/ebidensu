# modules/human/human.paste.ps1
# The operator copies a chat message (Ctrl+C in Teams) and presses Enter;
# the step reads the clipboard, checks the message with a regex and returns
# its named groups plus the moment it was taken. The moment is the point:
# the HOST team's "job XXX starts now" message is posted when the job
# starts, so the time
# the operator copies it is that job's start -- the scheduled time no
# spreadsheet has to carry. A time typed instead of Enter (10:30) is used
# in place of "now" (the leader's schedule, a message copied late).
#
# The panel stays until the clipboard holds a matching message for the
# expected job, the operator goes on without one (k: the time is now), or
# skips the item (s -> operator_skip, which a workflow maps to policy skip)
# or quits (q).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Get-EbiClipboardText
. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')      # ConvertTo-EbiHalfWidth

$Manifest = @{
  id         = 'human.paste'
  group      = 'human'
  summary    = 'Operator copies a chat message; read it from the clipboard, check it, note the time'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    message     = @{ type='string'; required=$true; desc='what to copy (shown on the panel)' }
    pattern     = @{ type='string'; required=$true; desc='.NET regex searched in the clipboard text; named groups become fields' }
    expect      = @{ type='string'; default=''; desc='value the expectGroup must have (full-width folded, case-insensitive); empty = any' }
    expectGroup = @{ type='string'; default='job' }
  }
  outputs    = @{
    text    = @{ type='string'; desc='the clipboard text used ("" when going on without one)' }
    fields  = @{ type='map';    desc='every named group of the pattern ("" when it did not match / no message)' }
    clock   = @{ type='string'; desc='HH:mm:ss: when Enter was pressed, or the time typed' }
    date    = @{ type='string'; desc='yyyy-MM-dd of that moment' }
    source  = @{ type='string'; desc='now | typed | without | auto' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
    @{ id = 'operator_skip'; transient = $false }
  )
  example    = @{ use = 'human.paste'; with = @{ message = 'Copy the Teams start message of the job and press Enter'; pattern = '\u30B8\u30E7\u30D6[:\uFF1A]\s*(?<job>[A-Z0-9]+)'; expect = '{{item.Excel_NAME}}' } }
  notes      = 'Under DryRun or without a console nobody is asked: source=auto, the time is now, fields are empty. Map operator_skip to policy skip in the workflow (onError.byFailure) so s skips the item without a second question.'
}

function HumanPaste-EmptyFields {
    # PURE. Every named group of the pattern -> '' (so a workflow can always
    # template steps.<id>.out.fields.<group>, matched or not).
    param([string]$Pattern)
    $f = @{}
    try { foreach ($n in ([regex]::new($Pattern)).GetGroupNames()) { if ($n -notmatch '^\d+$') { $f[$n] = '' } } } catch { }
    return $f
}

function HumanPaste-Check {
    <#
      PURE. Clipboard text -> @{ ok; fields; reason }. ok when the pattern
      matches and (expect empty or the expectGroup equals expect after
      full-width folding, case-insensitive).
    #>
    param([string]$Text, [string]$Pattern, [string]$Expect, [string]$ExpectGroup)
    $t = if ($null -eq $Text) { '' } else { $Text }
    $re = [regex]::new($Pattern)
    $m = $re.Match((ConvertTo-EbiHalfWidth -Value $t))
    if (-not $m.Success) { $m = $re.Match($t) }
    if (-not $m.Success) {
        $peek = ($t -replace '\s+', ' ').Trim()
        if ($peek.Length -gt 60) { $peek = $peek.Substring(0, 60) + '...' }
        return @{ ok = $false; fields = (HumanPaste-EmptyFields -Pattern $Pattern); reason = $(if ($peek -eq '') { 'the clipboard is empty' } else { 'not the expected message: "' + $peek + '"' }) }
    }
    $fields = @{}
    foreach ($n in $re.GetGroupNames()) { if ($n -notmatch '^\d+$') { $fields[$n] = $(if ($m.Groups[$n].Success) { $m.Groups[$n].Value.Trim() } else { '' }) } }
    if (-not [string]::IsNullOrWhiteSpace($Expect)) {
        $got = if ($fields.Contains($ExpectGroup)) { [string]$fields[$ExpectGroup] } else { '' }
        $a = (ConvertTo-EbiHalfWidth -Value $got).Trim().ToUpperInvariant()
        $b = (ConvertTo-EbiHalfWidth -Value $Expect).Trim().ToUpperInvariant()
        if ($a -ne $b) { return @{ ok = $false; fields = $fields; reason = ('the message is for ' + $got + ', this item is ' + $Expect) } }
    }
    return @{ ok = $true; fields = $fields; reason = '' }
}

function HumanPaste-Typed {
    # PURE. A typed answer -> @{ kind = enter|k|s|time|other; clock }.
    param([string]$Answer)
    $a = if ($null -eq $Answer) { '' } else { (ConvertTo-EbiHalfWidth -Value $Answer).Trim() }
    if ($a -eq '') { return @{ kind = 'enter'; clock = '' } }
    if ($a -eq 'k' -or $a -eq 's') { return @{ kind = $a; clock = '' } }
    if ($a -match '^([01]?\d|2[0-3])[:.]?([0-5]\d)(?:[:.]([0-5]\d))?$') {
        $sec = if ($Matches[3]) { [int]$Matches[3] } else { 0 }
        return @{ kind = 'time'; clock = ('{0:00}:{1:00}:{2:00}' -f [int]$Matches[1], [int]$Matches[2], $sec) }
    }
    return @{ kind = 'other'; clock = '' }
}

function Invoke-Step {
    param($In, $Ctx)
    $now = Get-Date
    $expect = [string]$In['expect']
    $what = New-Object System.Collections.ArrayList
    [void]$what.Add([string]$In['message'])
    if ($expect -ne '') { [void]$what.Add(('this item: ' + $expect)) }
    $next = @('Enter: read the clipboard; the start time is NOW',
              '10:30 (a time): read the clipboard; the start time is the time typed',
              'k: go on without a message (start time = now)',
              's: skip this item (leave it pending)',
              'q: cancel the whole run')
    while ($true) {
        $r = Show-EbiGate -Title 'PASTE' -What $what.ToArray() -Next $next -Actions @(@{ key = 'c'; label = 'Enter / 10:30 / k / s / q' }) -Default 'c' -Auto 'c' -DryRun ([bool]$Ctx['DryRun']) -Raw
        $now = Get-Date
        if ($r['auto']) {
            if ($Ctx['DryRun']) { $Ctx.Log.Info('would read the start message from the clipboard here (dry run: start time = now)') }
            return @{ ok = $true; text = ''; fields = (HumanPaste-EmptyFields -Pattern ([string]$In['pattern'])); clock = $now.ToString('HH:mm:ss'); date = $now.ToString('yyyy-MM-dd'); source = 'auto' }
        }
        if ([string]$r['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q'; text = ''; fields = @{}; clock = ''; date = ''; source = '' } }
        $typed = HumanPaste-Typed -Answer ([string]$r['note'])
        if ($typed['kind'] -eq 's') { return @{ ok = $false; failure = 'operator_skip'; message = 'operator skipped this item'; text = ''; fields = @{}; clock = ''; date = ''; source = '' } }
        if ($typed['kind'] -eq 'k') { return @{ ok = $true; text = ''; fields = (HumanPaste-EmptyFields -Pattern ([string]$In['pattern'])); clock = $now.ToString('HH:mm:ss'); date = $now.ToString('yyyy-MM-dd'); source = 'without' } }
        if ($typed['kind'] -eq 'other') {
            $what = New-Object System.Collections.ArrayList
            [void]$what.Add([string]$In['message'])
            [void]$what.Add(('not understood: "' + [string]$r['note'] + '" -- Enter, a time like 10:30, k, s or q'))
            continue
        }
        $text = Get-EbiClipboardText
        $c = HumanPaste-Check -Text $text -Pattern ([string]$In['pattern']) -Expect $expect -ExpectGroup ([string]$In['expectGroup'])
        if ($c['ok']) {
            $clock = if ($typed['kind'] -eq 'time') { $typed['clock'] } else { $now.ToString('HH:mm:ss') }
            $src = if ($typed['kind'] -eq 'time') { 'typed' } else { 'now' }
            $Ctx.Log.Info(('start message for {0} taken; start time {1} ({2})' -f $(if ($c['fields'].Contains([string]$In['expectGroup'])) { $c['fields'][[string]$In['expectGroup']] } else { '?' }), $clock, $src))
            return @{ ok = $true; text = $text; fields = $c['fields']; clock = $clock; date = $now.ToString('yyyy-MM-dd'); source = $src }
        }
        $what = New-Object System.Collections.ArrayList
        [void]$what.Add([string]$In['message'])
        if ($expect -ne '') { [void]$what.Add(('this item: ' + $expect)) }
        [void]$what.Add(('clipboard: ' + $c['reason']))
    }
}
