# modules/table/table.select.ps1
# Filter worklist rows (P1-26). The SAME function as the runner's
# source.select -- kernel/Worklist.ps1 Select-EbiWorklistRows -- so "ng is
# still pending" and the profile's verdict.values translation hold in both
# places (WORKFLOW-SCHEMA 3.1 / 3.2). The old Get-PendingRows treated any
# non-0 as done and hid NG rows; that rule is not here.

. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'table.select'
  group      = 'table'
  summary    = 'Rows whose field is pending under one of the five pendingWhen forms'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist    = @{ type='session'; sessionKind='worklist'; required=$true }
    field       = @{ type='string'; default=''; desc='the column to test (not needed for always)' }
    pendingWhen = @{ type='string'; default='!= ok'; desc='empty | != ok | == ng | bit !<name> | always (WORKFLOW-SCHEMA 3.1)' }
    only        = @{ type='list'; default=@(); desc='key display strings; keep only these rows' }
    limit       = @{ type='int';  default=0 }
  }
  outputs    = @{
    rows     = @{ type='list'; desc='the selected rows (copies of the hashtables)' }
    keyList  = @{ type='list'; desc='their key display strings' }
    selected = @{ type='int' }
    total    = @{ type='int' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'table.select'; with = @{ worklist = 'wl'; field = 'before_transferStatus'; pendingWhen = '!= ok' } }
  notes      = 'The five forms are checked by kernel/Worklist.ps1 (the same parser the runner uses); anything else is input_invalid naming the form list.'
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $field = [string]$In['field']; $pw = [string]$In['pendingWhen']
    $keyColumns = @($Ctx['KeyColumns'])
    if ($keyColumns.Count -eq 0) { $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    $spec = Get-EbiWorklistColumnSpec -Profile $Ctx['Profile'] -Field $field
    $only = @(@($In['only']) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' })
    $sel = Select-EbiWorklistRows -Rows $wl['rows'] -Field $field -PendingWhen $pw -Spec $spec -Only $(if ($only.Count) { $only } else { $null }) -KeyColumns $keyColumns -Limit ([int]$In['limit'])
    if (-not $sel['ok']) { return @{ ok = $false; failure = 'input_invalid'; message = $sel['message']; rows = @(); keyList = @(); selected = 0; total = 0 } }
    $copies = New-Object System.Collections.ArrayList
    $keys = New-Object System.Collections.ArrayList
    foreach ($r in @($sel['rows'])) {
        $c = @{}; foreach ($k in $r.Keys) { $c[[string]$k] = $r[$k] }
        [void]$copies.Add($c)
        [void]$keys.Add((Get-EbiKeyDisplay -Item $r -KeyColumns $keyColumns))
    }
    return @{ ok = $true; rows = $copies.ToArray(); keyList = $keys.ToArray(); selected = $sel['selected']; total = $sel['total'] }
}
