# modules/human/human.prepare.ps1
# Block until the operator says the screen is ready (Enter), or quits (q).
# Ported from Common.ps1 Wait-PagePrepared (P0-08). The old function called
# `exit` on q; a step never does that -- it returns operator_quit and the
# runner's onError policy decides (STEP-CONTRACT.md 3.1).
#
# P1-34 wires this to kernel/Gate.ps1's panel; the inputs/outputs here are
# the ones that panel will render, so workflows written now keep working.

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
  notes      = 'DryRun prints the prompt and answers Enter on the operator''s behalf, so a dry run never blocks.'
}

function HumanPrepare-Show {
    param([string]$Message, [string]$Url)
    Write-Host ''
    Write-Host ('  ' + $Message) -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace($Url)) {
        Write-Host ('  ' + $Url) -ForegroundColor DarkYellow
    }
}

function Invoke-Step {
    param($In, $Ctx)

    $message = [string]$In['message']
    $url     = if ($In.Contains('url')) { [string]$In['url'] } else { '' }

    HumanPrepare-Show -Message $message -Url $url

    if ($Ctx['DryRun']) {
        $Ctx.Log.Info('would wait for Enter here (dry run answers Enter)')
        return @{ ok = $true; action = 'enter' }
    }

    Write-Host '  Enter=OK / q=quit : ' -ForegroundColor Magenta -NoNewline
    $resp = Read-Host
    if ([string]$resp -eq 'q') {
        return @{ ok = $false; failure = 'operator_quit'; message = 'operator answered q'; action = 'quit' }
    }
    return @{ ok = $true; action = 'enter' }
}
