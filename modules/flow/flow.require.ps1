# modules/flow/flow.require.ps1
# Stop an item early, cleanly, when something it needs is not so: the
# value is not what the rest of the item requires -> failure not_met,
# which the workflow maps to its skip policy (onError.byFailure.not_met:
# skip), so the item is left pending with one line saying why instead of
# twenty steps failing one after another on missing inputs. 2026-10-13:
# logs / evidence for a job the track judged ng (failed on the OPEN side)
# have nothing to collect.

$Manifest = @{
  id         = 'flow.require'
  group      = 'flow'
  summary    = 'Fail with not_met unless a value equals what the rest of the item needs'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    value   = @{ type='string'; default=''; desc='the value to check (a template)' }
    equals  = @{ type='string'; default='ok'; desc='what it must be (text, case-sensitive)' }
    message = @{ type='string'; default=''; desc='why the item cannot go on; shown with the actual value' }
  }
  outputs    = @{
    met   = @{ type='bool' }
    value = @{ type='string'; desc='the value as seen' }
  }
  failures   = @(
    @{ id = 'not_met'; transient = $false }
  )
  example    = @{ use = 'flow.require'; with = @{ value = '{{steps.track.out.data.verdict}}'; equals = 'ok'; message = 'the track judged this job ng: nothing to collect' } }
  notes      = 'Pair it with onError.byFailure.not_met = skip in the workflow; under DryRun the check always passes (a dry run has no real values).'
}

function Invoke-Step {
    param($In, $Ctx)
    $v = if ($null -eq $In['value']) { '' } else { [string]$In['value'] }
    $want = [string]$In['equals']
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would require the value to be "' + $want + '" (dry run: passes)')); return @{ ok = $true; met = $true; value = $v } }
    if ($v -ceq $want) { return @{ ok = $true; met = $true; value = $v } }
    $why = [string]$In['message']; if ($why -eq '') { $why = 'the value is not "' + $want + '"' }
    return @{ ok = $false; failure = 'not_met'; message = ($why + ' (value: "' + $v + '", wanted "' + $want + '")'); met = $false; value = $v }
}
