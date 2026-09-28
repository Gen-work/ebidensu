# modules/human/human.prepare.ps1
# Block until the operator says the screen is ready (Enter), or quits (q).
# Ported from Common.ps1 Wait-PagePrepared (P0-08); on kernel/Gate.ps1's
# panel since P1-34. The old function called `exit` on q; a step never
# does that -- it returns operator_quit and the runner turns it into
# cancelled (STEP-CONTRACT.md 3.1, WORKFLOW-SCHEMA 5.2).

. (Join-Path $PSScriptRoot '..\..\kernel\Gate.ps1')

$Manifest = @{
  id         = 'human.prepare'
  group      = 'human'
  summary    = 'Show a message and wait for Enter (ready) or q (quit)'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    message = @{ type='string'; required=$true;  desc='what the operator must have ready before continuing' }
    url     = @{ type='string'; default='';      desc='optional URL or location hint printed under the message' }
  }
  outputs    = @{
    action  = @{ type='string'; desc='enter | quit' }
  }
  failures   = @(
    @{ id = 'operator_quit'; transient = $false }
  )
  example    = @{
    use  = 'human.prepare'
    with = @{ message = 'Open the page to capture, then press Enter'; url = 'https://example.invalid/list' }
  }
  notes      = 'DryRun (or no console) answers Enter on the operator''s behalf, so a dry run never blocks. After this step the foreground is the console (P0-R12): the next browser step brings its window back.'
}

function HumanPrepare-Actions { return @(@{ key = 'c'; label = 'continue (Enter)' }, @{ key = 'q'; label = 'quit' }) }

function Invoke-Step {
    param($In, $Ctx)
    $message = [string]$In['message']
    $url     = if ($In.Contains('url')) { [string]$In['url'] } else { '' }
    $r = Show-EbiGate -Title 'PREPARE' -What $message -Evidence $(if ($url -ne '') { @($url) } else { $null }) -Next @('Enter: the screen is ready, go on', 'q: cancel the whole run (teardown still runs)') -Actions (HumanPrepare-Actions) -Default 'c' -Auto 'c' -DryRun ([bool]$Ctx['DryRun'])
    if ($r['auto'] -and $Ctx['DryRun']) { $Ctx.Log.Info('would wait for Enter here (dry run answers Enter)') }
    if ([string]$r['action'] -eq 'q') {
        return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q'; action = 'quit' }
    }
    return @{ ok = $true; action = 'enter' }
}
