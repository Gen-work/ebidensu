# modules/verify/verify.compare_sets.ps1
# Do two lists name the same things? (Today's jobs by the plan vs by the
# operator's own sheet, files on one side vs the other.) Values are
# trimmed and full-width folded before comparing (kernel/Key.ps1), so a
# full-width digit typed into a sheet does not make a false difference.

. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')   # ConvertTo-EbiHalfWidth

$Manifest = @{
  id         = 'verify.compare_sets'
  group      = 'verify'
  summary    = 'Compare two lists as sets; ok when they hold the same values'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    left       = @{ type='list';   required=$true }
    right      = @{ type='list';   required=$true }
    leftLabel  = @{ type='string'; default='left' }
    rightLabel = @{ type='string'; default='right' }
  }
  outputs    = @{
    code      = @{ type='string'; desc='ok | unknown (a value on one side only)' }
    both      = @{ type='list' }
    onlyLeft  = @{ type='list' }
    onlyRight = @{ type='list' }
    reason    = @{ type='string' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.compare_sets'; with = @{ left = '{{steps.wbs.out.plucked}}'; right = '{{steps.map.out.plucked}}'; leftLabel = 'WBS'; rightLabel = 'mapping' } }
}

function VerifyCompareSets-Norm {
    param($Value)
    if ($null -eq $Value) { return '' }
    return (ConvertTo-EbiHalfWidth -Value ([string]$Value)).Trim().ToUpperInvariant()
}

function Invoke-Step {
    param($In, $Ctx)
    $L = [ordered]@{}; $R = [ordered]@{}
    foreach ($v in @($In['left']))  { $n = VerifyCompareSets-Norm $v; if ($n -ne '' -and -not $L.Contains($n)) { $L[$n] = [string]$v } }
    foreach ($v in @($In['right'])) { $n = VerifyCompareSets-Norm $v; if ($n -ne '' -and -not $R.Contains($n)) { $R[$n] = [string]$v } }
    $both = @(foreach ($k in $L.Keys) { if ($R.Contains($k)) { $L[$k] } })
    $ol = @(foreach ($k in $L.Keys) { if (-not $R.Contains($k)) { $L[$k] } })
    $or = @(foreach ($k in $R.Keys) { if (-not $L.Contains($k)) { $R[$k] } })
    $code = if ($ol.Count -eq 0 -and $or.Count -eq 0) { 'ok' } else { 'unknown' }
    $parts = New-Object System.Collections.ArrayList
    if ($ol.Count -gt 0) { [void]$parts.Add(('only in ' + $In['leftLabel'] + ': ' + ($ol -join ', '))) }
    if ($or.Count -gt 0) { [void]$parts.Add(('only in ' + $In['rightLabel'] + ': ' + ($or -join ', '))) }
    $reason = if ($parts.Count -gt 0) { $parts -join '; ' } else { ('' + $both.Count + ' value(s) on both sides') }
    return @{ ok = $true; code = $code; both = $both; onlyLeft = $ol; onlyRight = $or; reason = $reason }
}
