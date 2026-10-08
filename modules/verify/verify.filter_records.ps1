# modules/verify/verify.filter_records.ps1
# Keep the records that satisfy EVERY condition (the same twelve ops as
# verify.assert, kernel/Rules.ps1) and say where they were in the list.
# The list order is the page order, so the 1-based positions tell a later
# step which screen rows they are. Optional extras the evidence needs
# without arithmetic in the workflow: one field plucked from every kept
# record (plucked), and the time span the kept records cover (span).

. (Join-Path $PSScriptRoot '..\..\kernel\Rules.ps1')   # Test-EbiRuleHolds, ConvertTo-EbiDateTime

$Manifest = @{
  id         = 'verify.filter_records'
  group      = 'verify'
  summary    = 'Keep the records matching every condition; positions, plucked values, time span'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    records   = @{ type='list';   required=$true; desc='verify.parse_text output' }
    where     = @{ type='list';   default=@(); desc='conditions @{ field; op; value } (verify.assert ops); empty keeps all' }
    pluck     = @{ type='string'; default=''; desc='field whose value of every kept record goes to plucked' }
    spanFrom  = @{ type='string'; default=''; desc='time field whose earliest value starts span' }
    spanTo    = @{ type='string'; default=''; desc='time field whose latest value ends span (default spanFrom)' }
    padBeforeSec = @{ type='int'; default=0; desc='span.from moved this much earlier' }
    padAfterSec  = @{ type='int'; default=0; desc='span.to moved this much later' }
    expect    = @{ type='int';    default=-1; desc='how many records should match; -1 = any number. A different count is a warning, not a failure' }
  }
  outputs    = @{
    records = @{ type='list'; desc='the kept records, in list order' }
    indexes = @{ type='list'; desc='their 1-based positions in the input list' }
    first   = @{ type='int';  desc='position of the first kept record; 0 when none' }
    matched = @{ type='int' }
    plucked = @{ type='list'; desc='the pluck field of each kept record' }
    span    = @{ type='map';  desc='@{ from; to } ISO, empty map when spanFrom is blank or nothing matched' }
  }
  failures   = @(
    @{ id = 'rules_invalid'; transient = $false }
    @{ id = 'not_found';     transient = $true  }
  )
  example    = @{ use = 'verify.filter_records'; with = @{ records = '{{steps.rec.out.records}}'; where = @(@{ field = 'folder'; op = 'equals'; value = '/in' }); pluck = 'key' } }
  notes      = 'not_found when nothing matches (transient: the page may not show the new rows yet -- retry after a refresh). A condition with op within takes value = @{ from; to } (verify.time_window output).'
}

function Invoke-Step {
    param($In, $Ctx)
    $records = @(@($In['records']) | Where-Object { $null -ne $_ })
    $conds = @(@($In['where']) | Where-Object { $null -ne $_ })
    $ops = Get-EbiRuleOps
    $i = 0
    foreach ($c in $conds) {
        $i++
        if (-not ($c -is [System.Collections.IDictionary]) -or -not $c.Contains('field') -or -not $c.Contains('op') -or -not ($ops -contains [string]$c['op'])) {
            return @{ ok = $false; failure = 'rules_invalid'; message = ('condition ' + $i + ' needs field and op (one of ' + ($ops -join ', ') + ')'); records = @(); indexes = @(); first = 0; matched = 0; plucked = @(); span = @{} }
        }
    }
    $kept = New-Object System.Collections.ArrayList
    $idx = New-Object System.Collections.ArrayList
    for ($k = 0; $k -lt $records.Count; $k++) {
        $all = $true
        foreach ($c in $conds) { if (-not (Test-EbiRuleHolds -Record $records[$k] -Rule $c)) { $all = $false; break } }
        if ($all) { [void]$kept.Add($records[$k]); [void]$idx.Add($k + 1) }
    }
    $pf = [string]$In['pluck']
    $values = @(if ($pf -ne '') { foreach ($r in $kept) { if ($r -is [System.Collections.IDictionary] -and $r.Contains($pf)) { [string]$r[$pf] } else { '' } } })
    $span = @{}
    $sf = [string]$In['spanFrom']; $st = [string]$In['spanTo']; if ($st -eq '') { $st = $sf }
    if ($sf -ne '' -and $kept.Count -gt 0) {
        $lo = $null; $hi = $null
        foreach ($r in $kept) {
            $a = ConvertTo-EbiDateTime -Text ([string]$r[$sf]); if ($a['ok'] -and ($null -eq $lo -or $a['value'] -lt $lo)) { $lo = $a['value'] }
            $b = ConvertTo-EbiDateTime -Text ([string]$r[$st]); if ($b['ok'] -and ($null -eq $hi -or $b['value'] -gt $hi)) { $hi = $b['value'] }
        }
        if ($null -ne $lo -and $null -ne $hi) {
            $span = @{ from = $lo.AddSeconds(-1 * [int]$In['padBeforeSec']).ToString('yyyy-MM-ddTHH:mm:ss'); to = $hi.AddSeconds([int]$In['padAfterSec']).ToString('yyyy-MM-ddTHH:mm:ss') }
        }
    }
    $first = if ($idx.Count -gt 0) { [int]$idx[0] } else { 0 }
    $out = @{ ok = $true; records = $kept.ToArray(); indexes = $idx.ToArray(); first = $first; matched = $kept.Count; plucked = $values; span = $span }
    if ($kept.Count -eq 0) {
        if ($Ctx['DryRun']) { return $out }
        $out['ok'] = $false; $out['failure'] = 'not_found'; $out['message'] = ('none of ' + $records.Count + ' record(s) matched ' + $conds.Count + ' condition(s)')
        return $out
    }
    if ([int]$In['expect'] -ge 0 -and $kept.Count -ne [int]$In['expect']) {
        $out['warnings'] = @(@{ code = 'count_mismatch'; message = ('expected ' + $In['expect'] + ' record(s), matched ' + $kept.Count); data = @{ expected = [int]$In['expect']; matched = $kept.Count } })
    }
    return $out
}
