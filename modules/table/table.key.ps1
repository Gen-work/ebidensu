# modules/table/table.key.ps1
# The step face of kernel/Key.ps1 (P1-27): which of some candidate texts is
# this key? One hit at the best tier -> match; several -> ambiguous with the
# P0-R4 candidate shape for human.choose; none -> key_not_found. The rules
# (suffix / fullwidth / case-insensitive) are the profile's confirmedRules,
# grown through the ambiguity panel; this file holds no comparison of its
# own.

. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')

$Manifest = @{
  id         = 'table.key'
  group      = 'table'
  summary    = 'Match a key against candidate texts through the one rule set'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    key   = @{ type='string'; required=$true; desc='the key to find (display form for a composite key)' }
    among = @{ type='list';   required=$true; desc='candidate texts, or records (maps) whose key columns are compared' }
    rules = @{ type='list';   default=@(); desc='override of profile.worklist.key.confirmedRules' }
  }
  outputs    = @{
    match      = @{ type='any';    desc='the one matching entry (text or record); null otherwise' }
    index      = @{ type='int';    desc='its 0-based position in among; -1 otherwise' }
    matchedBy  = @{ type='string'; desc='exact | stripped | fullwidth | case | empty' }
    found      = @{ type='int';    desc='how many entries matched at that tier' }
    candidates = @{ type='map';    desc='P0-R4 candidate shape when ambiguous, else null' }
  }
  failures   = @(
    @{ id = 'key_not_found'; transient = $false }
    @{ id = 'ambiguous';     transient = $false }
  )
  example    = @{ use = 'table.key'; with = @{ key = '{{item.key}}'; among = '{{steps.rec.out.names}}' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $rules = @($In['rules'])
    if ($rules.Count -eq 0) { $rules = @(Get-EbiKeyRules -Profile $Ctx['Profile']) }
    $keyColumns = @($Ctx['KeyColumns'])
    if ($keyColumns.Count -eq 0) { $keyColumns = @(Get-EbiWorklistKeyColumns -Profile $Ctx['Profile']) }
    $among = @($In['among'])
    $m = Find-EbiKeyMatches -Records $among -Key ([string]$In['key']) -KeyColumns $keyColumns -Rules $rules
    $n = @($m['indexes']).Count
    if ($n -eq 0) { return @{ ok = $false; failure = 'key_not_found'; message = ('"' + [string]$In['key'] + '" matches none of ' + $among.Count + ' candidate(s)'); match = $null; index = -1; matchedBy = ''; found = 0; candidates = $null } }
    if ($n -gt 1) {
        $items = New-Object System.Collections.ArrayList
        foreach ($i in @($m['indexes'])) {
            $rec = $among[$i]
            $text = if ($rec -is [System.Collections.IDictionary]) { Get-EbiKeyOfRecord -Record $rec -KeyColumns $keyColumns } else { [string]$rec }
            [void]$items.Add(@{ candidate = $text; evidence = @{ position = ($i + 1); matchedBy = $m['tier'] } })
        }
        $cand = New-EbiCandidateList -Items $items.ToArray() -SuggestIndex -1 -Reason '' -Doubts ('' + $n + ' entries are the same key under the "' + $m['tier'] + '" rule; nothing here tells them apart')
        return @{ ok = $false; failure = 'ambiguous'; message = ('' + $n + ' candidates match "' + [string]$In['key'] + '" (' + $m['tier'] + ')'); match = $null; index = -1; matchedBy = $m['tier']; found = $n; candidates = $cand }
    }
    $idx = [int]$m['indexes'][0]
    return @{ ok = $true; match = $among[$idx]; index = $idx; matchedBy = $m['tier']; found = 1; candidates = $null }
}
