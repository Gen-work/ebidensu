#Requires -Version 5.1
# ============================================================
#  kernel/Gate.ps1
#
#  The one ASCII gate panel (P1-05). Dot-source only (no param() block,
#  ASCII source, no class).
#
#  Every place the old tool stopped to ask a person wrote its own
#  Read-Host (27 files, 77 prompts); this file replaces all of them with
#  one panel of four fixed sections --
#
#      WHAT HAPPENED / NEXT / EVIDENCE / ACTIONS
#
#  -- rendered at most 80 columns wide, never wrapping in an 80-column
#  console, with one answer parser that understands the standard keys:
#      r retry   s skip   q quit   m <note> (a note to keep)
#  plus whatever extra keys the caller offers (Enter / n / y / 1..N).
#
#  Split of concerns:
#      Format-EbiGatePanel   PURE: text -> lines
#      Read-EbiGateAnswer    PURE: what the operator typed -> action + note
#      Show-EbiGate          the impure loop: render, read, repeat until
#                            the answer is one of the offered actions.
#                            The reader is injectable (tests) and, when
#                            no console can answer (DryRun, redirected
#                            stdin, or EBI_NO_ASK=1 as the test runner
#                            sets it), the default action is taken and
#                            the line says so.
#      Invoke-EbiGateAsk     the runner's -AskHandler built on the above
#                            (the two question shapes of kernel/Runner.ps1's
#                            Invoke-EbiDefaultAsk).
# ============================================================

. (Join-Path $PSScriptRoot 'Native.ps1')   # console window back to the front before a question

function Get-EbiGateWidth { return 80 }

function Save-EbiConsoleWindow {
    # Remember the window the operator started the run from (the one in
    # front right now: they just pressed Enter in it), so every question
    # can bring it back after a step put a browser / DF / Excel in front.
    # Windows only; never throws.
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return }
    try {
        [void](Get-EbiNative)
        $h = [EbiNative]::GetForegroundWindow()
        if ($h -eq [IntPtr]::Zero) { $h = [EbiNative]::GetConsoleWindow() }
        $global:EbiConsoleHwnd = $h
    } catch { }
}

function Set-EbiWindowFront {
    <#
      The runner's foreground rule (office feedback 2026-10-09: "only the
      page that is needed in front, the console otherwise, so I can do
      other things while it runs"), called before every step:
        -FrontHWnd  the call's "front" window (a session name the workflow
                    gave, e.g. a screen capture that must see the page)
                    -> bring it in front
        -NeedsForeground  the step brings its own window -> leave it
        otherwise   a window a previous step put in front -> console back
      Never throws; a no-op off Windows.
    #>
    param([IntPtr]$FrontHWnd = [IntPtr]::Zero, [bool]$NeedsForeground = $false)
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return }
    try {
        if ($FrontHWnd -ne [IntPtr]::Zero) { [void](Set-EbiForeground -HWnd $FrontHWnd -SettleMs 300); return }
        if ($NeedsForeground) { return }
        if ($global:EbiWindowOut) { Restore-EbiConsoleWindow -Quiet; $global:EbiWindowOut = $false }
    } catch { }
}

function Restore-EbiConsoleWindow {
    # Bring the remembered console back before reading an answer. The
    # first office run: after the browser steps the panel waited in a
    # console hidden behind Edge. Best effort; never throws. -Quiet: no
    # blink, no beep (the runner putting the console back between steps).
    param([switch]$Quiet)
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return }
    try {
        $h = $global:EbiConsoleHwnd
        if ($null -eq $h -or $h -eq [IntPtr]::Zero) { return }
        if ([EbiNative]::GetForegroundWindow() -ne $h) {
            $fg = Set-EbiForeground -HWnd $h -SettleMs 100
            # still behind: blink its taskbar button so it can be found
            if ($Quiet) { return }
            if (-not $fg['ok']) { [void][EbiNative]::FlashWindow($h, $true) }
            try { [Console]::Beep(880, 120) } catch { }   # the console was behind: say a question is waiting
        }
    } catch { }
}

function ConvertTo-EbiGateWrapped {
    # PURE. Wrap one text to at most $Width characters, breaking at spaces
    # when possible and hard-breaking a word longer than the width. Empty
    # text is one empty line. Existing line breaks are respected.
    param([string]$Text, [int]$Width)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Text) { $Text = '' }
    if ($Width -lt 1) { $Width = 1 }
    foreach ($para in ($Text -split "`r?`n")) {
        if ($para.Length -le $Width) { [void]$out.Add($para); continue }
        $line = ''
        foreach ($word in ($para -split ' ')) {
            while ($word.Length -gt $Width) {
                if ($line -ne '') { [void]$out.Add($line); $line = '' }
                [void]$out.Add($word.Substring(0, $Width))
                $word = $word.Substring($Width)
            }
            if ($line -eq '') { $line = $word }
            elseif (($line.Length + 1 + $word.Length) -le $Width) { $line = $line + ' ' + $word }
            else { [void]$out.Add($line); $line = $word }
        }
        [void]$out.Add($line)
    }
    return $out.ToArray()
}

function Format-EbiGatePanel {
    <#
      PURE. The panel as an array of lines, each at most Width characters.
        Title     one line, shown in the top border
        What      what happened (string or string[])
        Next      what will happen after each action (string or string[])
        Evidence  paths / values worth looking at (string or string[])
        Actions   ordered list of @{ key; label } (rendered "key=label")
      A section with nothing to say is left out.
    #>
    param([string]$Title = '', $What = $null, $Next = $null, $Evidence = $null, $Actions = @(), [int]$Width = 0)
    if ($Width -le 0) { $Width = Get-EbiGateWidth }
    $inner = $Width - 4          # "| " + text + " |"
    $lines = New-Object System.Collections.ArrayList

    $t = if ([string]::IsNullOrWhiteSpace($Title)) { '' } else { ' ' + $Title.Trim() + ' ' }
    if ($t.Length -gt ($Width - 4)) { $t = $t.Substring(0, $Width - 7) + '...' }
    [void]$lines.Add('+' + $t + ('-' * ($Width - 2 - $t.Length)) + '+')

    function Add-Section {
        param([string]$Label, $Body)
        $items = New-Object System.Collections.ArrayList
        foreach ($b in @($Body)) { if ($null -ne $b -and [string]$b -ne '') { [void]$items.Add([string]$b) } }
        if ($items.Count -eq 0) { return }
        [void]$lines.Add('| ' + $Label.PadRight($inner) + ' |')
        foreach ($item in $items) {
            foreach ($w in @(ConvertTo-EbiGateWrapped -Text $item -Width ($inner - 2))) {
                [void]$lines.Add('|   ' + $w.PadRight($inner - 2) + ' |')
            }
        }
    }
    Add-Section -Label 'WHAT HAPPENED' -Body $What
    Add-Section -Label 'NEXT' -Body $Next
    Add-Section -Label 'EVIDENCE' -Body $Evidence

    $acts = New-Object System.Collections.ArrayList
    foreach ($a in @($Actions)) {
        if ($a -is [System.Collections.IDictionary]) { [void]$acts.Add([string]$a['key'] + '=' + [string]$a['label']) }
        elseif ($null -ne $a) { [void]$acts.Add([string]$a) }
    }
    if ($acts.Count -gt 0) {
        [void]$lines.Add('+' + ('-' * ($Width - 2)) + '+')
        foreach ($w in @(ConvertTo-EbiGateWrapped -Text ($acts.ToArray() -join '   ') -Width $inner)) {
            [void]$lines.Add('| ' + $w.PadRight($inner) + ' |')
        }
    }
    [void]$lines.Add('+' + ('-' * ($Width - 2)) + '+')
    return $lines.ToArray()
}

function Read-EbiGateAnswer {
    <#
      PURE. What the operator typed -> @{ ok; action; note } or ok=$false.
        ''            -> the Default action (when one is offered)
        'r' 's' 'q'   -> that action, if offered
        'm some text' -> action 'm' with note 'some text' (if 'm' offered)
        '2'           -> action '2' for numbered choices
      Keys are matched case-insensitively; anything not offered is ok=$false.
    #>
    param([string]$Text, $Actions, [string]$Default = '')
    $keys = New-Object System.Collections.ArrayList
    foreach ($a in @($Actions)) {
        if ($a -is [System.Collections.IDictionary]) { [void]$keys.Add(([string]$a['key']).ToLowerInvariant()) }
        elseif ($null -ne $a) { [void]$keys.Add((([string]$a -split '=')[0]).ToLowerInvariant()) }
    }
    $t = if ($null -eq $Text) { '' } else { $Text.Trim() }
    if ($t -eq '') {
        if ($Default -ne '' -and $keys -contains $Default.ToLowerInvariant()) { return @{ ok = $true; action = $Default.ToLowerInvariant(); note = '' } }
        return @{ ok = $false; action = ''; note = '' }
    }
    $m = [regex]::Match($t, '^([A-Za-z0-9]+)(?:\s+(.*))?$')
    if (-not $m.Success) { return @{ ok = $false; action = ''; note = '' } }
    $key = $m.Groups[1].Value.ToLowerInvariant()
    $note = if ($m.Groups[2].Success) { $m.Groups[2].Value.Trim() } else { '' }
    if (-not ($keys -contains $key)) { return @{ ok = $false; action = ''; note = '' } }
    if ($key -ne 'm' -and $note -ne '') { return @{ ok = $false; action = ''; note = '' } }   # "s something" is a typo, not a skip
    return @{ ok = $true; action = $key; note = $note }
}

function Show-EbiGate {
    <#
      Render the panel, read an answer, repeat until it is one of the
      offered actions. Returns @{ action; note; auto }.
        Actions   ordered @{ key; label } list
        Default   the action Enter means ('' = Enter is not an answer)
        Auto      the action to take when nobody can answer (DryRun,
                  redirected stdin); '' = still try to read
        Reader    scriptblock returning the typed line (tests); default
                  Read-Host
        Raw       free-text mode: the typed line is the answer (note under
                  the Default action); only a bare 'q' is an action
      Nothing here decides what an action MEANS -- the caller does.
    #>
    param([string]$Title = '', $What = $null, $Next = $null, $Evidence = $null, $Actions = @(), [string]$Default = '', [string]$Auto = '', [scriptblock]$Reader = $null, [bool]$DryRun = $false, [switch]$Raw)
    foreach ($line in @(Format-EbiGatePanel -Title $Title -What $What -Next $Next -Evidence $Evidence -Actions $Actions)) {
        Write-Host ('  ' + $line) -ForegroundColor Yellow
    }
    $noConsole = $DryRun
    $why = 'dry run'
    if (-not $noConsole -and $null -eq $Reader) {
        try { if ([Console]::IsInputRedirected) { $noConsole = $true; $why = 'no console to ask' } } catch { }
        # EBI_NO_ASK=1: nobody is meant to answer (Tests\Run-Tests.ps1 sets
        # it), so a real console must not stop the run either.
        if (-not $noConsole -and [string]$env:EBI_NO_ASK -eq '1') { $noConsole = $true; $why = 'EBI_NO_ASK' }
    }
    if ($noConsole -and $Auto -ne '') {
        Write-Host ('  (' + $why + ': ' + $Auto + ')') -ForegroundColor DarkGray
        return @{ action = $Auto; note = ''; auto = $true }
    }
    $read = if ($null -ne $Reader) { $Reader } else { { Read-Host } }
    if ($null -eq $Reader) { Restore-EbiConsoleWindow }
    $tries = 0
    while ($true) {
        $tries++
        Write-Host '  > ' -ForegroundColor Magenta -NoNewline
        $typed = [string](& $read)
        if ($Raw) {
            # free text (human.input, P2-07): 'q' alone quits, anything else --
            # including nothing -- comes back as the note under the Default action
            if ($typed.Trim() -eq 'q') { return @{ action = 'q'; note = ''; auto = $false } }
            return @{ action = $Default; note = $typed.Trim(); auto = $false }
        }
        $ans = Read-EbiGateAnswer -Text $typed -Actions $Actions -Default $Default
        if ($ans['ok']) { return @{ action = $ans['action']; note = $ans['note']; auto = $false } }
        if ($tries -ge 20 -and $Auto -ne '') { Write-Host ('  (no valid answer after 20 tries: ' + $Auto + ')') -ForegroundColor DarkGray; return @{ action = $Auto; note = ''; auto = $true } }
        Write-Host ('  not an option: "' + $typed + '"') -ForegroundColor DarkYellow
    }
}

function Invoke-EbiGateAsk {
    <#
      The runner's -AskHandler on top of the panel (the two question
      shapes documented on kernel/Runner.ps1's Invoke-EbiDefaultAsk):
        error   -> 'r' | 's' | 'q'   (auto: s)
        confirm -> 'y' | 'n' | 'q'   (auto: y)
    #>
    param([hashtable]$Question, [bool]$DryRun = $false, [scriptblock]$Reader = $null)
    $where = if ([string]$Question['key'] -ne '') { $Question['section'] + '[' + $Question['key'] + ']/' + $Question['id'] } else { $Question['section'] + '/' + $Question['id'] }
    if ([string]$Question['kind'] -eq 'confirm') {
        $withText = ''
        if ($Question.Contains('with') -and ($Question['with'] -is [System.Collections.IDictionary])) {
            $parts = New-Object System.Collections.ArrayList
            foreach ($k in ($Question['with'].Keys | Sort-Object)) { [void]$parts.Add([string]$k + '=' + [string]$Question['with'][$k]) }
            $withText = $parts.ToArray() -join ' '
        }
        $r = Show-EbiGate -Title 'CONFIRM' `
            -What @(('{0} ({1}) is a destructive step' -f $where, $Question['use']), $withText) `
            -Next @('y: run it', 'n: skip this item, leave it pending', 'q: cancel the whole run (teardown still runs)') `
            -Actions @(@{ key = 'y'; label = 'do it' }, @{ key = 'n'; label = 'skip item' }, @{ key = 'q'; label = 'quit' }) `
            -Auto 'y' -Reader $Reader -DryRun $DryRun
        return $r['action']
    }
    $evidence = New-Object System.Collections.ArrayList
    if ($Question.Contains('evidence') -and ($Question['evidence'] -is [System.Collections.IDictionary])) {
        foreach ($k in ($Question['evidence'].Keys | Sort-Object)) { $v = $Question['evidence'][$k]; if ($null -ne $v -and -not ($v -is [System.Collections.IDictionary]) -and -not ($v -is [System.Collections.IList])) { [void]$evidence.Add([string]$k + ': ' + [string]$v) } }
    }
    $attemptText = if ($Question.Contains('attempt') -and [int]$Question['attempt'] -gt 1) { (' (attempt ' + [string]$Question['attempt'] + ')') } else { '' }
    $r = Show-EbiGate -Title 'FAILED' `
        -What @(('{0} ({1}){2}' -f $where, $Question['use'], $attemptText), ([string]$Question['failure'] + ': ' + [string]$Question['message'])) `
        -Next @('r: run this step again', 's: skip this item, leave it pending', 'q: cancel the whole run (teardown still runs)') `
        -Evidence $evidence.ToArray() `
        -Actions @(@{ key = 'r'; label = 'retry' }, @{ key = 's'; label = 'skip item' }, @{ key = 'q'; label = 'quit' }) `
        -Auto 's' -Reader $Reader -DryRun $DryRun
    return $r['action']
}
