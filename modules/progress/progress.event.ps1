# modules/progress/progress.event.ps1
# Append one event to the run trace (P1-29, over P0-03's kernel/Trace.ps1):
# what a workflow wants a person to see later without opening the worklist
# -- the successor of ProgressLog.ps1's status\progress.jsonl.

. (Join-Path $PSScriptRoot '..\..\kernel\Trace.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'progress.event'
  group      = 'progress'
  summary    = 'Append a workflow-authored event to run/<runId>/trace.jsonl'
  tier       = 'core'
  effects    = 'write'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    action  = @{ type='string'; required=$true; desc='short verb, e.g. captured, verdict, delivered' }
    status  = @{ type='string'; default='info'; enum=@('ok', 'fail', 'skip', 'info', 'start'); desc='same words the runner uses' }
    message = @{ type='string'; default='' }
    key     = @{ type='string'; default=''; desc='the item; empty = the current item, if any' }
    data    = @{ type='map';    default=@{}; desc='payload, JSON-serializable' }
  }
  outputs    = @{ key = @{ type='string' } }
  failures   = @(
    @{ id = 'internal_error'; transient = $true }
  )
  example    = @{ use = 'progress.event'; with = @{ action = 'verdict'; status = 'ok'; message = '{{steps.verdict.out.message}}' } }
  notes      = 'Not idempotent by declaration (two calls are two events), though a replayed step on resume is not re-run, so the trace holds one event per item per run.'
}

function Invoke-Step {
    param($In, $Ctx)
    $key = [string]$In['key']
    if ($key -eq '' -and $null -ne $Ctx['Item']) { $key = Get-EbiKeyDisplay -Item $Ctx['Item'] -KeyColumns @($Ctx['KeyColumns']) }
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would record event {0}/{1} for {2}' -f $In['action'], $In['status'], $(if ($key -ne '') { $key } else { '(run)' }))); return @{ ok = $true; key = $key } }
    try {
        Write-TraceEvent -WorkDir ([string]$Ctx['WorkDir']) -RunId ([string]$Ctx['RunId']) -Phase 'workflow' -Key $key -Tags @{ step = 'progress.event' } -Action ([string]$In['action']) -Status ([string]$In['status']) -Message ([string]$In['message']) -Data $In['data']
    } catch { return @{ ok = $false; failure = 'internal_error'; message = $_.Exception.Message; key = $key } }
    return @{ ok = $true; key = $key }
}
