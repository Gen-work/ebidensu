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
      Returns $false only when a front window was asked for and did not
      come to the front. Never throws; a no-op ($true) off Windows.
    #>
    param([IntPtr]$FrontHWnd = [IntPtr]::Zero, [bool]$NeedsForeground = $false)
    if ([System.Environment]::OSVersion.Platform -ne [System.PlatformID]::Win32NT) { return $true }
    try {
        if ($FrontHWnd -ne [IntPtr]::Zero) { return [bool](Set-EbiForeground -HWnd $FrontHWnd -SettleMs 300)['ok'] }
        if ($NeedsForeground) { return $true }
        if ($global:EbiWindowOut) { Restore-EbiConsoleWindow -Quiet; $global:EbiWindowOut = $false }
    } catch { return ($FrontHWnd -eq [IntPtr]::Zero) }
    return $true
}

function Get-EbiSessionHandle {
    # PURE. The window handle of a $Ctx.Session entry (@{ kind; value; ... }
    # -- the runner's own record, not the bare value a step receives) or of
    # a bare value; Zero when there is none. 3560785 passed the whole entry
    # to ConvertTo-EbiHandle, which turned every "front" into Zero.
    param($Entry)
    if ($Entry -is [System.Collections.IDictionary]) { if (-not $Entry.Contains('value')) { return [IntPtr]::Zero }; $Entry = $Entry['value'] }
    try { return (ConvertTo-EbiHandle $Entry) } catch { return [IntPtr]::Zero }
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

function Get-EbiGateLineColor {
    # PURE. Yellow, or Red when the line holds one of the Alert fragments.
    param([string]$Line, [string[]]$Alert = @())
    foreach ($a in @($Alert)) { if (-not [string]::IsNullOrEmpty($a) -and $Line.IndexOf($a, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { return 'Red' } }
    return 'Yellow'
}

function Get-EbiFailureHints {
    <#
      PURE. The FAILED panel's "what usually causes this / what to do"
      lines, so the operator does not have to look the code up (10/15:
      "less time to rtfm"). Two tables, both checked:
        - message patterns (a COM HRESULT says more than the failure id)
        - failure ids (the STEP-CONTRACT failure table's ids)
      Returns string[] (empty when nothing is known). Add a row here when
      a new failure turns up on the office PC -- and a line to
      docs/ebi-dance/KNOWLEDGE.md saying where it came from.
    #>
    param([string]$Failure, [string]$Message = '', [string]$Use = '')
    $out = New-Object System.Collections.ArrayList
    $byMessage = @(
        @{ pattern = '0x800AC472|0x80010001|0x8001010A|RPC_E_CALL_REJECTED|RPC_E_SERVERCALL_RETRYLATER'
           hint = 'Excel is busy and refused the call: a cell is being edited, a dialog is open, or someone clicked in Excel while the tool was writing. Press Esc in Excel, keep your hands off it, then r.' },
        @{ pattern = '0x800A03EC'
           hint = 'Excel rejected the value or the range (protected sheet, merged cells, a formula error). Look at the sheet named above, then r or s.' },
        @{ pattern = 'read-only|ReadOnly|has it open'
           hint = 'Someone else has the workbook open. Ask them to close it (the name is above when Excel knows it), then r.' }
    )
    foreach ($row in $byMessage) { if ($Message -match $row['pattern']) { [void]$out.Add($row['hint']) } }
    $byId = @{
        'not_found'         = $(if ($Use -eq 'screen.find_window') { 'The window title is the ACTIVE TAB''s title. Give the page its own browser window, make it the active tab, wait until it has loaded, then r.' } else { 'Nothing matched. Check the item / page named above; r after fixing it, s to leave the item pending.' })
        'foreground_lost'   = 'Another window came to the front (a popup, a Teams toast, a click). Close it, leave the mouse alone, then r.'
        'window_gone'       = 'The window was closed or replaced (browser restarted?). Open it again, then r.'
        'no_window'         = 'The window is not open. Open it, then r.'
        'no_browser_window' = 'The browser window is not open. Open it, then r.'
        'file_not_found'    = $(if ($Message -match 'ebi\.local\.json|paths\.') { 'A setting file is missing: the message names the ebi.local.json entry. Put the file there (or fix the path in ebi.local.json), then r.' } else { 'An earlier flow did not produce this file for this job (track / logs), or a path in ebi.local.json is wrong. s, then run that flow with -Only <job>.' })
        'dir_not_found'     = 'The folder does not exist: check the paths in ebi.local.json (and the share is reachable), then r.'
        'write_failed'      = 'Writing failed. If Excel shows a dialog or a cell in edit mode, close it (Esc), then r.'
        'save_failed'       = 'Saving failed: the workbook may be open read-only by someone else, or the share is unreachable. r after fixing it.'
        'copy_failed'       = 'Copy / paste through the clipboard failed (another program used the clipboard?). r.'
        'clipboard_error'   = 'The clipboard was busy (another program held it). r.'
        'session_missing'   = 'A step before this one was skipped or failed, so what it opens is not there. s (the item stays pending).'
        'sheet_not_found'   = 'The workbook has no sheet with that name: wrong workbook, or the sheet was renamed. Check it, then s.'
        'wrong_page'        = 'The page is not the expected one (filter, columns, tab). Fix the page as the PREPARE panel says, then r.'
        'page_loading'      = 'The page is still loading. Wait, then r.'
        'timeout'           = 'Nothing came in time (page slow, file not there yet). r to wait again, s to leave the item pending.'
        'read_only'         = 'Someone else has the workbook open. Ask them to close it, then r.'
        'open_failed'       = 'Excel could not open the file (path, lock, or a broken file). Check it, then r.'
        'download_failed'   = 'The download did not arrive (session expired? proxy?). Check the browser, then r.'
        'input_invalid'     = 'The workflow passed a wrong value to this step: not your screen. q and report it (the message names the parameter).'
        'contract_violation' = 'The step definition is wrong: not your screen. q and report it.'
        'ambiguous'         = 'Several candidates matched; the item cannot be decided automatically. s, then run it with -Only <job> after fixing the data.'
    }
    if ($byId.Contains($Failure)) { [void]$out.Add([string]$byId[$Failure]) }
    return $out.ToArray()
}

function Test-EbiStopRequested {
    <#
      Was q (or Esc) pressed in the console while a step was working?
      Reads -- and so drops -- every key typed ahead; true when one of them
      was q / Q / Esc. The runner asks before the next step (STOP? panel).
      False with no console to read (redirected stdin, EBI_NO_ASK=1 as the
      test runner sets it, Linux CI). Never throws.
    #>
    if ([string]$env:EBI_NO_ASK -eq '1') { return $false }
    $hit = $false
    try {
        if ([Console]::IsInputRedirected) { return $false }
        while ([Console]::KeyAvailable) {
            $k = [Console]::ReadKey($true)
            if ($k.Key -eq [ConsoleKey]::Escape -or $k.KeyChar -eq 'q' -or $k.KeyChar -eq 'Q') { $hit = $true }
        }
    } catch { }
    return $hit
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
        Alert     text fragments: a panel line holding one of them is
                  drawn red instead of yellow (10/15: "a run from another
                  day", "the picture shows another job") -- the panel text
                  itself is unchanged, so tests and logs read the same
      Nothing here decides what an action MEANS -- the caller does.
    #>
    param([string]$Title = '', $What = $null, $Next = $null, $Evidence = $null, $Actions = @(), [string]$Default = '', [string]$Auto = '', [scriptblock]$Reader = $null, [bool]$DryRun = $false, [switch]$Raw, [string[]]$Alert = @())
    foreach ($line in @(Format-EbiGatePanel -Title $Title -What $What -Next $Next -Evidence $Evidence -Actions $Actions)) {
        Write-Host ('  ' + $line) -ForegroundColor (Get-EbiGateLineColor -Line $line -Alert $Alert)
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
        error   -> 'r' | 's' | 'q'   (auto: s; Enter = r when the failure is
                                      transient, else s -- 10/15)
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
    # Enter: r when the failure is transient (window not open yet, Excel
    # busy -- trying again is what a person does), s otherwise (a missing
    # track.json does not appear by retrying). 10/14 Enter = s; 10/15 split.
    $transient = $Question.Contains('transient') -and [bool]$Question['transient']
    $enter = if ($transient) { 'r' } else { 's' }
    $what = New-Object System.Collections.ArrayList
    [void]$what.Add(('{0} ({1}){2}' -f $where, $Question['use'], $attemptText))
    [void]$what.Add(([string]$Question['failure'] + ': ' + [string]$Question['message']))
    foreach ($h in @(Get-EbiFailureHints -Failure ([string]$Question['failure']) -Message ([string]$Question['message']) -Use ([string]$Question['use']))) { [void]$what.Add('hint: ' + $h) }
    $r = Show-EbiGate -Title 'FAILED' `
        -What $what.ToArray() `
        -Next @($(if ($enter -eq 'r') { 'r / Enter: run this step again' } else { 'r: run this step again' }), $(if ($enter -eq 's') { 's / Enter: skip this item, leave it pending' } else { 's: skip this item, leave it pending' }), 'q: cancel the whole run (teardown still runs)') `
        -Evidence $evidence.ToArray() `
        -Actions @(@{ key = 'r'; label = $(if ($enter -eq 'r') { 'retry (Enter)' } else { 'retry' }) }, @{ key = 's'; label = $(if ($enter -eq 's') { 'skip item (Enter)' } else { 'skip item' }) }, @{ key = 'q'; label = 'quit' }) `
        -Default $enter -Auto 's' -Reader $Reader -DryRun $DryRun
    return $r['action']
}
