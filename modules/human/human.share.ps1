# modules/human/human.share.ps1
# Put a message -- text lines and pictures -- on the clipboard for the
# operator to paste into a chat or mail, then wait until they say it is
# sent. The step never pastes or sends by itself: what leaves the machine
# is what a person looked at.
#   rich      one paste: HTML (pictures as data: URIs) + RTF (\pngblip) +
#             plain text in one clipboard object (kernel/RichClip.ps1),
#             the way a mixed selection copied out of Word pastes
#   sequence  the text first, then each picture on its own, Enter between
#             -- the fallback for a target that drops pictures from a
#             rich paste

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\RichClip.ps1')

$Manifest = @{
  id         = 'human.share'
  group      = 'human'
  summary    = 'Put text and pictures on the clipboard for the operator to paste and send'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    lines    = @{ type='list';   default=@(); desc='text lines, in order' }
    images   = @{ type='list';   default=@(); desc='PNG paths, in order after the text; a nested list is flattened (a path, then a step''s list of paths)' }
    mode     = @{ type='string'; default='rich'; enum=@('rich', 'sequence') }
    where    = @{ type='string'; default='the chat'; desc='where to paste, for the prompt' }
    maxWidth = @{ type='int';    default=0; desc='rich mode: show pictures at most this wide (px); 0 = natural size' }
  }
  outputs    = @{
    action = @{ type='string'; desc='sent | skip' }
    note   = @{ type='string'; desc='free text typed with m' }
    pieces = @{ type='int';    desc='clipboard loads made' }
  }
  failures   = @(
    @{ id = 'operator_quit';   transient = $false }
    @{ id = 'file_not_found';  transient = $false }
    @{ id = 'clipboard_error'; transient = $true  }
  )
  example    = @{ use = 'human.share'; with = @{ lines = @('{{profile.messages.done}}'); images = '{{steps.shots.out.paths}}'; where = 'Teams' } }
  notes      = 'Answers: Enter = pasted and sent, s = not sent (the item stays pending for whatever checkpoint follows), m <text> = sent with a note, q = cancel the run. Dry run / no console answers Enter. Foreground is the console afterwards.'
}

function HumanShare-Actions { return @(@{ key = 'c'; label = 'sent (Enter)' }, @{ key = 's'; label = 'not sent' }, @{ key = 'm'; label = 'sent + note' }, @{ key = 'q'; label = 'quit' }) }

function HumanShare-Picture {
    param([string]$Path, [int]$MaxWidth)
    $bytes = [System.IO.File]::ReadAllBytes($Path)
    $info = Get-EbiPngInfo -Bytes $bytes
    $w = [int]$info['width']; $h = [int]$info['height']
    if ($MaxWidth -gt 0 -and $w -gt $MaxWidth) { $h = [int][Math]::Round($h * $MaxWidth / [double]$w); $w = $MaxWidth }
    return @{ bytes = $bytes; width = $w; height = $h }
}

function HumanShare-Ask {
    param([string]$What, [string[]]$Evidence, $Ctx)
    return (Show-EbiGate -Title 'SHARE' -What $What -Evidence $Evidence -Next @('Enter: pasted and sent', 's: not sent (leave the item pending)', 'm <text>: sent, with a note', 'q: cancel the whole run') -Actions (HumanShare-Actions) -Default 'c' -Auto 'c' -DryRun ([bool]$Ctx['DryRun']))
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $lines = @(@($In['lines']) | ForEach-Object { [string]$_ })
    $flat = New-Object System.Collections.ArrayList
    foreach ($x in @($In['images'])) { if ($x -is [System.Collections.IList] -and -not ($x -is [string])) { foreach ($y in $x) { [void]$flat.Add($y) } } else { [void]$flat.Add($x) } }
    $imgs = @($flat | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | ForEach-Object { Resolve-EbiWorkPath -PathValue ([string]$_) -WorkDir $work })
    $where = [string]$In['where']
    if ($Ctx['DryRun']) {
        $Ctx.Log.Info(('would put {0} line(s) + {1} picture(s) on the clipboard ({2}) for {3}' -f $lines.Count, $imgs.Count, $In['mode'], $where))
        return @{ ok = $true; action = 'sent'; note = ''; pieces = 0 }
    }
    foreach ($p in $imgs) { if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $p; action = ''; note = ''; pieces = 0 } } }
    $text = $lines -join "`r`n"
    $pieces = 0
    if ([string]$In['mode'] -eq 'sequence') {
        $steps = New-Object System.Collections.ArrayList
        if ($lines.Count -gt 0) { [void]$steps.Add(@{ kind = 'text'; value = $text }) }
        foreach ($p in $imgs) { [void]$steps.Add(@{ kind = 'image'; value = $p }) }
        $i = 0
        foreach ($s in $steps) {
            $i++
            $r = if ($s['kind'] -eq 'text') { try { Set-EbiClipboardText -Text $s['value']; @{ ok = $true; message = '' } } catch { @{ ok = $false; message = $_.Exception.Message } } } else { Set-EbiClipboardImage -Path $s['value'] }
            if (-not $r['ok']) { return @{ ok = $false; failure = 'clipboard_error'; message = $r['message']; action = ''; note = ''; pieces = $pieces } }
            $pieces++
            $last = ($i -eq $steps.Count)
            $what = ('piece {0}/{1} is on the clipboard ({2}): paste it into {3}{4}' -f $i, $steps.Count, $s['kind'], $where, $(if ($last) { ', then send' } else { ' (do not send yet), then Enter for the next piece' }))
            $evText = [string]$s['value']
            if ($evText.Length -gt 200) { $evText = $evText.Substring(0, 200) }
            $ans = HumanShare-Ask -What $what -Evidence @($evText) -Ctx $Ctx
            if ($ans['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q'; action = ''; note = ''; pieces = $pieces } }
            if ($ans['action'] -eq 's') { return @{ ok = $true; action = 'skip'; note = ''; pieces = $pieces } }
            if ($last) { return @{ ok = $true; action = 'sent'; note = [string]$ans['note']; pieces = $pieces } }
        }
        return @{ ok = $true; action = 'sent'; note = ''; pieces = $pieces }
    }
    $pics = @(foreach ($p in $imgs) { HumanShare-Picture -Path $p -MaxWidth ([int]$In['maxWidth']) })
    $html = New-EbiCfHtml -Fragment (New-EbiShareHtml -Lines $lines -Pictures $pics)
    $rtf = New-EbiShareRtf -Lines $lines -Pictures $pics
    $r = Set-EbiClipboardRich -CfHtml $html -Rtf $rtf -Text $text
    if (-not $r['ok']) { return @{ ok = $false; failure = 'clipboard_error'; message = $r['message']; action = ''; note = ''; pieces = 0 } }
    $ev = New-Object System.Collections.ArrayList
    foreach ($l in $lines) { [void]$ev.Add($l) }
    foreach ($p in $imgs) { [void]$ev.Add($p) }
    $ans = HumanShare-Ask -What ('text + ' + $imgs.Count + ' picture(s) are on the clipboard: paste (Ctrl+V) into ' + $where + ', check it, send') -Evidence $ev.ToArray() -Ctx $Ctx
    if ($ans['action'] -eq 'q') { return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q'; action = ''; note = ''; pieces = 1 } }
    if ($ans['action'] -eq 's') { return @{ ok = $true; action = 'skip'; note = ''; pieces = 1 } }
    return @{ ok = $true; action = 'sent'; note = [string]$ans['note']; pieces = 1 }
}
