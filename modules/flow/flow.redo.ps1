# modules/flow/flow.redo.ps1
# Send the current item back to an earlier step (10/23 track: the start
# time was wrong and another job's file came in; a file was put right by
# hand). A gate answer (human.gate action retime / recheck) is looked up in
# map; when it is there, redo names the step the item goes back to and the
# runner jumps there (only for this step's use, at most 5 times an item).
# answer is handed to the first step after the jump as if typed there
# (human.paste reads it as the start time), so nobody is asked twice.

$Manifest = @{
  id         = 'flow.redo'
  group      = 'flow'
  summary    = 'Send the item back to an earlier step when an answer says so (the runner jumps)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    action = @{ type='string'; default=''; desc='the answer that may ask for it (human.gate action)' }
    map    = @{ type='map';    default=@{}; desc='action -> id of the earlier step to go back to, e.g. retime -> start, recheck -> gift' }
    answer = @{ type='string'; default=''; desc='handed to the first step after the jump as its typed answer (a start time)' }
  }
  outputs    = @{
    redo   = @{ type='string'; desc='the step id the item goes back to; empty = go on' }
    answer = @{ type='string'; desc='what the step after the jump gets as typed' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'flow.redo'; with = @{ action = '{{steps.gate.out.action}}'; map = @{ retime = 'start'; recheck = 'gift' }; answer = '{{steps.gate.out.clock}}' } }
  notes      = 'The jump is the runner''s (kernel/Runner.ps1): only to an EARLIER step of the same each list; the steps in between run again with fresh outputs. More than 5 jumps for one item are refused (warning redo_limit) so a loop cannot run forever.'
}

function FlowRedo-Target {
    # PURE. action + map -> the step id to go back to, or ''.
    param([string]$Action, $Map)
    if ($Action -eq '' -or -not ($Map -is [System.Collections.IDictionary])) { return '' }
    if ($Map.Contains($Action) -and $null -ne $Map[$Action]) { return ([string]$Map[$Action]).Trim() }
    return ''
}

function Invoke-Step {
    param($In, $Ctx)
    $to = FlowRedo-Target -Action ([string]$In['action']) -Map $In['map']
    if ($to -eq '') { return @{ ok = $true; redo = ''; answer = '' } }
    $Ctx.Log.Info(('back to step "{0}" ({1}){2}' -f $to, [string]$In['action'], $(if ([string]$In['answer'] -ne '') { ', answer ' + [string]$In['answer'] } else { '' })))
    return @{ ok = $true; redo = $to; answer = [string]$In['answer'] }
}
