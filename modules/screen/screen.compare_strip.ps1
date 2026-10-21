# modules/screen/screen.compare_strip.ps1
# Do two pictures of the same rows show the same thing in one column?
# (10/21, the operator: picture 1 of the GFIX receive-result sheet -- the
# new snap of the overview list -- must show the same job as picture 1 of
# the GIFT receive-result sheet; compare the job column's pixels, not OCR,
# which kept misreading the job name.) The second picture is scaled to
# the first one's height (then +-6 %), the strip of the first is searched
# for across the second (10/21 FJODWWBV: compared at a fixed place, the
# same job read 52 -- the two pictures have other margins), and the least
# mean grey difference decides:
#   ok       score <= okMax: the same cells
#   ng       score >= ngMin: different text there (another job, other rows)
#   unknown  in between, a blank strip, a picture missing, sizes that
#            do not fit -- a person looks
# A CHECK, never a stop: the workflow puts `alert` on the gate in red.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Image.ps1')    # Compare-EbiPngStrip

$Manifest = @{
  id         = 'screen.compare_strip'
  group      = 'screen'
  summary    = 'Compare the same column strip of two pictures of the same rows (pixels); ok / ng / unknown + alert'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    a        = @{ type='any';    default=''; desc='the picture the strip is measured on (a path, or a list: its first path is taken)' }
    b        = @{ type='any';    default=''; desc='the picture compared with it (a path or a list)' }
    fromFrac = @{ type='any';    default=0;  desc='left edge of the strip, 0..1 of a''s width (excel.snap_list jobFrom)' }
    toFrac   = @{ type='any';    default=1;  desc='right edge, 0..1 (jobTo)' }
    bodyFrom = @{ type='any';    default=0;  desc='top of the compared rows, 0..1 of a''s height (excel.snap_list bodyFrom: under the headers, which are the same on every job)' }
    okMax    = @{ type='int';    default=8;  desc='score at or below this = the same' }
    ngMin    = @{ type='int';    default=25; desc='score at or above this = different' }
    minInk   = @{ type='any';    default=0.005; desc='a strip with less dark ink than this share says nothing (unknown)' }
    label    = @{ type='string'; default='the pictures'; desc='for the reason / alert wording' }
    saveBase = @{ type='path';   default=''; desc='also save what was compared: <saveBase>.a.png (a''s strip) and .b.png (the best place in b); empty = none' }
  }
  outputs    = @{
    code   = @{ type='string'; desc='ok | ng | unknown' }
    score  = @{ type='any';    desc='mean grey difference 0..255; -1 when not measured' }
    reason = @{ type='string'; desc='one line for the gate' }
    alert  = @{ type='string'; desc='the reason when code is not ok, else empty' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'screen.compare_strip'; with = @{ a = 'capture/{{item.keySafe}}/pic1gfix/pic1.png'; b = 'capture/{{item.keySafe}}/pic1gift/pic1.png'; fromFrac = 0.8; toFrac = 1; label = 'picture 1 job column' } }
  notes      = 'okMax / ngMin are first guesses (10/21); the gate and gate.json show the score so they can be set from real runs.'
}

function ScreenCompareStrip-First {
    # PURE. A path or a list of paths -> the first non-empty one, or ''.
    param($Value)
    foreach ($v in @($Value)) { if (-not [string]::IsNullOrWhiteSpace([string]$v)) { return [string]$v } }
    return ''
}

function ScreenCompareStrip-Verdict {
    # PURE. score / ink -> ok | ng | unknown.
    param([double]$Score, [double]$Ink, [double]$OkMax, [double]$NgMin, [double]$MinInk)
    if ($Score -lt 0 -or $Ink -lt $MinInk) { return 'unknown' }
    if ($Score -le $OkMax) { return 'ok' }
    if ($Score -ge $NgMin) { return 'ng' }
    return 'unknown'
}

function Invoke-Step {
    param($In, $Ctx)
    $work = [string]$Ctx['WorkDir']
    $a = ScreenCompareStrip-First $In['a']; $b = ScreenCompareStrip-First $In['b']
    $label = [string]$In['label']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would compare the strip of {0} with {1}' -f $a, $b)); return @{ ok = $true; code = 'unknown'; score = -1; reason = ''; alert = '' } }
    if ($a -eq '' -or $b -eq '') {
        $why = ('{0}: nothing to compare ({1} missing)' -f $label, $(if ($a -eq '') { 'the new picture' } else { 'the other picture' }))
        return @{ ok = $true; code = 'unknown'; score = -1; reason = $why; alert = $why }
    }
    $pa = Resolve-EbiWorkPath -PathValue $a -WorkDir $work; $pb = Resolve-EbiWorkPath -PathValue $b -WorkDir $work
    $r = Compare-EbiPngStrip -PathA $pa -PathB $pb -FromFrac ([double]$In['fromFrac']) -ToFrac ([double]$In['toFrac']) -BodyFrom ([double]$In['bodyFrom']) -SaveBase $(if ([string]$In['saveBase'] -ne '') { Resolve-EbiWorkPath -PathValue ([string]$In['saveBase']) -WorkDir $work } else { '' })
    if (-not $r['ok']) {
        $why = ('{0}: not compared -- {1}' -f $label, $r['message'])
        return @{ ok = $true; code = 'unknown'; score = -1; reason = $why; alert = $why }
    }
    $code = ScreenCompareStrip-Verdict -Score ([double]$r['score']) -Ink ([double]$r['ink']) -OkMax ([double]$In['okMax']) -NgMin ([double]$In['ngMin']) -MinInk ([double]$In['minInk'])
    $why = switch ($code) {
        'ok'    { '{0}: the same (difference {1}, {2})' -f $label, $r['score'], $r['how'] }
        'ng'    { '{0}: DIFFERENT (difference {1} at {2}, the same would be <= {3}): another job''s picture?' -f $label, $r['score'], $r['how'], $In['okMax'] }
        default { if ([double]$r['ink'] -lt [double]$In['minInk']) { '{0}: the strip is blank -- look by eye' -f $label } else { '{0}: not sure (difference {1}, between {2} and {3}) -- look by eye' -f $label, $r['score'], $In['okMax'], $In['ngMin'] } }
    }
    $Ctx.Log.Info($why)
    return @{ ok = $true; code = $code; score = $r['score']; reason = $why; alert = $(if ($code -eq 'ok') { '' } else { $why }) }
}
