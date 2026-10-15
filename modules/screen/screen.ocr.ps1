# modules/screen/screen.ocr.ps1
# Read the text in a PNG with Windows OCR (kernel/Ocr.ps1). A CHECK step:
# when OCR is not available on this PC (no recognizer language, an old
# Windows, Linux) it returns ok with available = false and a warning, so a
# workflow that only uses the text to raise a flag keeps going (10/15: the
# picture-1 job name check marks the gate red, it never stops the run).

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Ocr.ps1')

$Manifest = @{
  id         = 'screen.ocr'
  group      = 'screen'
  summary    = 'OCR a PNG (Windows OCR); no OCR on this PC = ok with available=false'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    path     = @{ type='path';   default=''; desc='the PNG to read; relative paths resolve under the work dir. Empty = nothing to read (an earlier export that did not happen): ok, available=false, no warning' }
    language = @{ type='string'; default='en-US'; desc='recognizer tried first (BCP-47); then en, then the Windows profile languages. en reads job / file names (O/0, I/1) better than ja' }
  }
  outputs    = @{
    text      = @{ type='string'; desc='every line, joined with a newline; empty when nothing was read' }
    lines     = @{ type='list';   desc='the lines OCR found, top to bottom' }
    language  = @{ type='string'; desc='the recognizer that read it; empty when none ran' }
    available = @{ type='bool';   desc='false: OCR could not run here (the text is empty for that reason, not because the picture is blank)' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $false }
  )
  example    = @{ use = 'screen.ocr'; with = @{ path = 'capture/gfix/{{item.keySafe}}/pic1.png' } }
  notes      = 'OCR misreads look-alikes (O/0, I/1, S/5, B/8): compare with verify.find_token, which folds them. Windows only; on another host available=false.'
}

function Invoke-Step {
    param($In, $Ctx)
    if ([string]::IsNullOrWhiteSpace([string]$In['path'])) { $Ctx.Log.Info('no image to read'); return @{ ok = $true; text = ''; lines = @(); language = ''; available = $false } }
    $p = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would OCR {0}' -f $p)); return @{ ok = $true; text = ''; lines = @(); language = ''; available = $false } }
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = ('no image at ' + $p); text = ''; lines = @(); language = ''; available = $false } }
    $r = Invoke-EbiOcrFile -Path $p -Language ([string]$In['language'])
    if (-not $r['ok']) {
        return @{ ok = $true; text = ''; lines = @(); language = ''; available = $false; warnings = @(@{ code = 'ocr_unavailable'; message = ('OCR did not run: ' + [string]$r['message']) }) }
    }
    $Ctx.Log.Info(('read {0} line(s) from {1} ({2})' -f @($r['lines']).Count, $p, $r['language']))
    return @{ ok = $true; text = [string]$r['text']; lines = @($r['lines']); language = [string]$r['language']; available = $true }
}
