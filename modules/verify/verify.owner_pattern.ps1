# modules/verify/verify.owner_pattern.ps1
# The WBS owner cell -> the regex that says "this row is mine" (the
# OwnerFilter.ps1 arrow rule: the bare name, "name<-other" (taken over by
# me), "other->name" (handed to me); "name->other" / "other<-name" are NOT
# mine). The owner comes from the profile (ebi.local.json
# { "operator": { "owner": "..." } }); an EMPTY owner means nobody filters:
# every row of the day is kept, and the step says so.

. (Join-Path $PSScriptRoot '..\..\kernel\Worklist.ps1')   # Get-EbiOwnerPattern

$Manifest = @{
  id         = 'verify.owner_pattern'
  group      = 'verify'
  summary    = 'Build the WBS owner-cell regex for one operator (empty owner = everyone)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    owner = @{ type='string'; default=''; desc='the operator as the WBS writes it (e.g. a family name); empty = no owner filter' }
  }
  outputs    = @{
    pattern  = @{ type='string'; desc='regex for an excel.read_rows where { op: matches }: the name alone, name<-other, other->name; "" (matches every cell) when owner is empty' }
    everyone = @{ type='bool';   desc='true when owner is empty: the plan takes every owner''s rows' }
    owner    = @{ type='string'; desc='the owner, trimmed' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.owner_pattern'; with = @{ owner = '{{profile.operator.owner}}' } }
  notes      = 'The arrows are the full-width ones the WBS uses (U+2190 / U+2192). An owner that contains regex characters is escaped.'
}

function VerifyOwnerPattern-Build {
    # PURE. owner -> regex ('' when owner is blank). The rule itself lives in
    # kernel/Worklist.ps1 (Get-EbiOwnerPattern): the runner filters the day
    # workflows' rows with the same pattern.
    param([string]$Owner)
    return (Get-EbiOwnerPattern -Owner $Owner)
}

function Invoke-Step {
    param($In, $Ctx)
    $o = ([string]$In['owner']).Trim()
    $p = VerifyOwnerPattern-Build -Owner $o
    if ($p -eq '') { $Ctx.Log.Info('no operator.owner set: the owner column is not filtered (every owner''s rows of the day). Set { "operator": { "owner": "<name>" } } in ebi.local.json to keep only yours.') }
    else { $Ctx.Log.Info(('owner filter: ' + $o + ' (the bare name, ' + $o + [char]0x2190 + 'other, other' + [char]0x2192 + $o + ')')) }
    return @{ ok = $true; pattern = $p; everyone = ($p -eq ''); owner = $o }
}
