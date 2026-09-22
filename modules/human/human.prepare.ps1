# modules/human/human.prepare.ps1
# Show a message and block until the operator presses Enter.
# Source: Common.ps1 Wait-PagePrepared, minus the exit-on-q (a step reports
# quit as an output; the runner turns it into the reserved 'cancelled').

$Manifest = @{
  id         = 'human.prepare'
  group      = 'human'
  summary    = 'Show a message and block until the operator presses Enter'
  tier       = 'core'
  effects    = 'ui'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    message   = @{ type='string'; required=$true; desc='what to ask the operator to prepare before continuing' }
    url       = @{ type='string'; default='';     desc='shown under the message when non-empty (which page to open)' }
    allowQuit = @{ type='bool';   default=$true;  desc='accept q as a request to abort the run' }
  }
  outputs    = @{
    action = @{ type='string'; desc='enter | quit' }
  }
  failures   = @(
    @{ id = 'no_console'; transient = $false }
  )
  example    = @{ use='human.prepare'; with=@{ message='Open the target page in the browser, then press Enter.' } }
  notes      = 'Takes the console foreground (effects=ui, P0-R12): the next key-sending step must re-activate its own window. action=quit is turned into the reserved cancelled failure by the runner (WORKFLOW-SCHEMA 5.2).'
}

function Invoke-Step {
    param($In, $Ctx)

    if ($Ctx.DryRun) {
        $Ctx.Log.Info(('would prompt: ' + [string]$In.message))
        return @{ ok = $true; action = 'enter' }
    }

    Write-Host ''
    Write-Host ([string]$In.message) -ForegroundColor Yellow
    if (-not [string]::IsNullOrWhiteSpace([string]$In.url)) {
        Write-Host ('  ' + [string]$In.url) -ForegroundColor DarkGray
    }
    if ([bool]$In.allowQuit) {
        Write-Host 'Enter=OK / q=quit : ' -ForegroundColor Magenta -NoNewline
    } else {
        Write-Host 'Enter=OK : ' -ForegroundColor Magenta -NoNewline
    }

    $resp = $null
    try {
        $resp = Read-Host
    } catch {
        return @{ ok = $false; failure = 'no_console'; message = ('cannot read from the console: ' + $_.Exception.Message) }
    }

    if ([bool]$In.allowQuit -and ([string]$resp).Trim() -eq 'q') {
        return @{ ok = $true; action = 'quit' }
    }
    return @{ ok = $true; action = 'enter' }
}
