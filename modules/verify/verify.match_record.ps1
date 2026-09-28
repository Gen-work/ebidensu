# modules/verify/verify.match_record.ps1
# The record for a key among parsed records (P1-32, ported from
# SnapVerify.ps1 Get-MatchedRowIndex / Select-JenkinsFileCandidate). Key
# comparison is kernel/Key.ps1's (no -eq here). Several hits: tieBreak
# 'newest' picks the latest by a time field (inside the run window first
# when one is given, as the old code did), 'first' the first listed,
# 'none' refuses with the P0-R4 candidate shape so a person decides.

. (Join-Path $PSScriptRoot '..\..\kernel\Key.ps1')
. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')

$Manifest = @{
  id         = 'verify.match_record'
  group      = 'verify'
  summary    = 'Find the record for a key; newest / first / ask on several'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    records   = @{ type='list';   required=$true; desc='verify.parse_text output' }
    key       = @{ type='string'; required=$true; desc='what to find (a column value, or the key display form)' }
    field     = @{ type='string'; default='key'; desc='the record field that holds the key' }
    tieBreak  = @{ type='string'; default='newest'; enum=@('newest', 'first', 'none') }
    timeField = @{ type='string'; default='time'; desc='the record field with the time, for newest' }
    window    = @{ type='map';    default=@{}; desc='{ from, to } ISO; with newest, a hit inside the window beats one outside' }
  }
  outputs    = @{
    record     = @{ type='map';    desc='the chosen record (null when none)' }
    index      = @{ type='int';    desc='1-based position among the records; 0 when none' }
    found      = @{ type='int';    desc='how many records matched the key' }
    matchedBy  = @{ type='string'; desc='exact | stripped | fullwidth | case' }
    candidates = @{ type='map';    desc='P0-R4 candidate shape when ambiguous with tieBreak none' }
  }
  failures   = @(
    @{ id = 'record_not_found'; transient = $true  }
    @{ id = 'ambiguous';        transient = $false }
  )
  example    = @{ use = 'verify.match_record'; with = @{ records = '{{steps.rec.out.records}}'; key = '{{item.Correl_ID_S}}'; tieBreak = 'newest' } }
  notes      = 'record_not_found is transient: the row is usually still arriving. The key input is a raw column value ({{item.<col>}}), not {{item.key}}, unless the records carry the composite key in one field.'
}

function VerifyMatchRecord-Pick {
    # kernel/Parse.ps1 Select-EbiNewestRecord under the step prefix.
    param($Hits, [string]$TieBreak, [string]$TimeField, $Window)
    return (Select-EbiNewestRecord -Hits $Hits -TieBreak $TieBreak -TimeField $TimeField -Window $Window)
}

function Invoke-Step {
    param($In, $Ctx)
    $records = @($In['records'])
    $field = [string]$In['field']; $key = [string]$In['key']; $tie = [string]$In['tieBreak']; $tf = [string]$In['timeField']
    $rules = @(Get-EbiKeyRules -Profile $Ctx['Profile'])
    $values = @(foreach ($r in $records) { if ($r -is [System.Collections.IDictionary] -and $r.Contains($field)) { [string]$r[$field] } else { '' } })
    $m = Find-EbiKeyMatches -Records $values -Key $key -Rules $rules
    $n = @($m['indexes']).Count
    if ($n -eq 0) { return @{ ok = $false; failure = 'record_not_found'; message = ('no record whose "' + $field + '" is "' + $key + '" among ' + $records.Count); record = $null; index = 0; found = 0; matchedBy = ''; candidates = $null } }
    $hits = @(foreach ($i in @($m['indexes'])) { @{ index = ($i + 1); record = $records[$i] } })
    if ($n -gt 1 -and $tie -eq 'none') {
        $items = @(foreach ($h in $hits) { @{ candidate = [string]$h['record'][$field]; evidence = @{ position = $h['index']; time = $(if ($h['record'].Contains($tf)) { [string]$h['record'][$tf] } else { '' }); matchedBy = $m['tier'] } } })
        $cand = New-EbiCandidateList -Items $items -SuggestIndex -1 -Doubts ('' + $n + ' records carry this key; tieBreak is none, so nobody picked')
        return @{ ok = $false; failure = 'ambiguous'; message = ('' + $n + ' records match "' + $key + '"'); record = $null; index = 0; found = $n; matchedBy = $m['tier']; candidates = $cand }
    }
    $pick = VerifyMatchRecord-Pick -Hits $hits -TieBreak $tie -TimeField $tf -Window $In['window']
    $warnings = @()
    if ($n -gt 1) { $warnings = @( @{ code = 'several_matches'; message = ('' + $n + ' records match "' + $key + '"; took #' + $pick['index'] + ' (' + $pick['reason'] + ')'); data = @{ found = $n; index = $pick['index'] } } ) }
    return @{ ok = $true; record = $pick['record']; index = $pick['index']; found = $n; matchedBy = $m['tier']; candidates = $null; warnings = $warnings }
}
