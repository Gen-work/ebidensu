# modules/progress/progress.status.ps1
# The ASCII progress table (P1-29, ported from VerifyTool.ps1 Show-Status):
# per field, how many rows are done / pending / ng, judged through the same
# pendingWhen + verdict.values machinery the runner selects by, so the
# numbers here are the numbers `ebi run` will act on.

. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')

$Manifest = @{
  id         = 'progress.status'
  group      = 'progress'
  summary    = 'Print done / pending / ng counts per field as an ASCII table'
  tier       = 'core'
  effects    = 'read'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    worklist = @{ type='session'; sessionKind='worklist'; required=$true }
    field    = @{ type='string'; default=''; desc='one field; empty = every verdict / bitmask column the profile declares' }
    groupBy  = @{ type='string'; default=''; desc='also count per value of this column' }
  }
  outputs    = @{
    lines   = @{ type='list'; desc='the rendered table' }
    summary = @{ type='map';  desc='field -> { total; done; pending; ng }' }
  }
  failures   = @(
    @{ id = 'column_missing'; transient = $false }
  )
  example    = @{ use = 'progress.status'; with = @{ worklist = 'wl'; field = 'before_transferStatus' } }
}

function ProgressStatus-Fields {
    # PURE. Which columns to report: the input, else the profile's verdict / bitmask columns.
    param([string]$Field, $Profile, $Columns)
    if ($Field -ne '') { return @($Field) }
    $out = New-Object System.Collections.ArrayList
    if ($null -ne $Profile -and ($Profile -is [System.Collections.IDictionary]) -and $Profile.Contains('worklist') -and ($Profile['worklist'] -is [System.Collections.IDictionary]) -and $Profile['worklist'].Contains('columns') -and $null -ne $Profile['worklist']['columns']) {
        foreach ($c in $Profile['worklist']['columns']) {
            if ($c -is [System.Collections.IDictionary] -and $c.Contains('name') -and $c.Contains('role') -and ([string]$c['role'] -in @('verdict', 'bitmask')) -and (@($Columns) -contains [string]$c['name'])) { [void]$out.Add([string]$c['name']) }
        }
    }
    return $out.ToArray()
}

function ProgressStatus-Count {
    # PURE. Rows + one field -> @{ total; done; pending; ng; bits (map or null) }.
    param($Rows, [string]$Field, $Spec)
    $isBits = ($null -ne $Spec -and ($Spec -is [System.Collections.IDictionary]) -and $Spec.Contains('role') -and [string]$Spec['role'] -eq 'bitmask')
    $total = 0; $done = 0; $pending = 0; $ng = 0
    $bitDone = @{}
    $notOk = ConvertFrom-EbiPendingWhen -Text '!= ok'
    $isNg = ConvertFrom-EbiPendingWhen -Text '== ng'
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $total++
        if ($isBits) {
            $allSet = $true
            if ($Spec.Contains('bits') -and ($Spec['bits'] -is [System.Collections.IDictionary])) {
                foreach ($bn in $Spec['bits'].Keys) {
                    $p = ConvertFrom-EbiPendingWhen -Text ('bit !' + [string]$bn)
                    $isPending = Test-EbiRowPending -Row $row -Field $Field -Parsed $p -Spec $Spec
                    if (-not $bitDone.Contains([string]$bn)) { $bitDone[[string]$bn] = 0 }
                    if (-not $isPending) { $bitDone[[string]$bn]++ } else { $allSet = $false }
                }
            } else { $allSet = $false }
            if ($allSet) { $done++ } else { $pending++ }
            continue
        }
        if (Test-EbiRowPending -Row $row -Field $Field -Parsed $notOk -Spec $Spec) {
            $pending++
            if (Test-EbiRowPending -Row $row -Field $Field -Parsed $isNg -Spec $Spec) { $ng++ }
        } else { $done++ }
    }
    return @{ total = $total; done = $done; pending = $pending; ng = $ng; bits = $(if ($isBits) { $bitDone } else { $null }) }
}

function ProgressStatus-Lines {
    # PURE. summary map (field -> counts) in the given order -> ASCII lines.
    param($Fields, $Summary, $GroupCounts = $null)
    $L = New-Object System.Collections.ArrayList
    $w = 0
    foreach ($f in @($Fields)) { if (([string]$f).Length -gt $w) { $w = ([string]$f).Length } }
    if ($w -lt 5) { $w = 5 }
    $hdr = 'field'.PadRight($w) + '  total   done  pending     ng'
    [void]$L.Add($hdr)
    [void]$L.Add('-' * $hdr.Length)
    foreach ($f in @($Fields)) {
        $s = $Summary[[string]$f]
        [void]$L.Add(([string]$f).PadRight($w) + '  ' + ([string]$s['total']).PadLeft(5) + '  ' + ([string]$s['done']).PadLeft(5) + '  ' + ([string]$s['pending']).PadLeft(7) + '  ' + ([string]$s['ng']).PadLeft(5))
        if ($null -ne $s['bits']) {
            foreach ($bn in ($s['bits'].Keys | Sort-Object)) { [void]$L.Add(('  bit ' + [string]$bn).PadRight($w) + '  ' + ''.PadLeft(5) + '  ' + ([string]$s['bits'][$bn]).PadLeft(5)) }
        }
    }
    if ($null -ne $GroupCounts -and $GroupCounts.Count -gt 0) {
        [void]$L.Add('')
        foreach ($g in ($GroupCounts.Keys | Sort-Object)) { [void]$L.Add('  ' + ([string]$g).PadRight($w) + '  ' + ([string]$GroupCounts[$g]).PadLeft(5)) }
    }
    return $L.ToArray()
}

function Invoke-Step {
    param($In, $Ctx)
    $wl = $In['worklist']
    $columns = @($wl['columns']); $rows = @($wl['rows'])
    $fields = @(ProgressStatus-Fields -Field ([string]$In['field']) -Profile $Ctx['Profile'] -Columns $columns)
    foreach ($f in $fields) { if (-not ($columns -contains $f)) { return @{ ok = $false; failure = 'column_missing'; message = ('column "' + $f + '" is not in the worklist'); lines = @(); summary = @{} } } }
    $summary = @{}
    foreach ($f in $fields) { $summary[$f] = ProgressStatus-Count -Rows $rows -Field $f -Spec (Get-EbiWorklistColumnSpec -Profile $Ctx['Profile'] -Field $f) }
    $groups = $null
    $gb = [string]$In['groupBy']
    if ($gb -ne '' -and ($columns -contains $gb)) {
        $groups = @{}
        foreach ($r in $rows) { $g = if ($null -ne $r -and $r.Contains($gb) -and $null -ne $r[$gb]) { [string]$r[$gb] } else { '' }; if (-not $groups.Contains($g)) { $groups[$g] = 0 }; $groups[$g]++ }
    }
    $lines = @(ProgressStatus-Lines -Fields $fields -Summary $summary -GroupCounts $groups)
    $warnings = @()
    if ($fields.Count -eq 0) { $warnings = @( @{ code = 'no_fields'; message = 'no field given and the profile declares no verdict / bitmask columns'; data = @{} } ) }
    Write-Host ''
    Write-Host ('  worklist: ' + $rows.Count + ' row(s)' + $(if ([string]$wl['path'] -ne '') { '  ' + [string]$wl['path'] } else { '' })) -ForegroundColor Cyan
    foreach ($l in $lines) { Write-Host ('  ' + $l) }
    return @{ ok = $true; lines = $lines; summary = $summary; warnings = $warnings }
}
