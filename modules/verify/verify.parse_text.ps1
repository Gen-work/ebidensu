# modules/verify/verify.parse_text.ps1
# Page text -> records through a profile grammar (P1-30 / P1-31): the four
# parsers of PROFILE-SCHEMA.md section 4 live in kernel/Parse.ps1; this is
# the step face. Unrecognised lines go out on the standard warnings channel
# with their line number and text (P0-R5) -- the runner traces them and the
# end-of-run summary shows them -- and are counted in outputs, so a page
# that silently lost rows cannot look like a page that had none.

. (Join-Path $PSScriptRoot '..\..\kernel\Parse.ps1')

$Manifest = @{
  id         = 'verify.parse_text'
  group      = 'verify'
  summary    = 'Parse page text into records with a delimited / labeled / columns / regex grammar'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    text        = @{ type='string'; required=$true; desc='the page text (browser.wait_for / read_text output)' }
    grammar     = @{ type='map';    required=$true; desc='{ parser: delimited|labeled|columns|regex, ... } (PROFILE-SCHEMA 4)' }
    maxWarnings = @{ type='int';    default=20; desc='unrecognised lines reported one by one up to this many, then one summary' }
  }
  outputs    = @{
    records      = @{ type='list'; desc='hashtables field -> text, plus _line' }
    recordCount  = @{ type='int' }
    unrecognized = @{ type='int';  desc='non-blank lines the grammar did not recognise' }
    names        = @{ type='list'; desc='the "key" field of every record when the grammar has one (for table.key)' }
  }
  failures   = @(
    @{ id = 'grammar_invalid'; transient = $false }
    @{ id = 'no_records';      transient = $true  }
  )
  example    = @{ use = 'verify.parse_text'; with = @{ text = '{{steps.wait.out.text}}'; grammar = '{{page.grammar}}' } }
  notes      = 'no_records is transient: the usual cause is a page still loading. A labeled grammar always yields one record; a label it could not find is a warning (label_missing) with the field blank, and verify.assert''s present/empty rules decide what that means.'
}

function Invoke-Step {
    param($In, $Ctx)
    $r = ConvertFrom-EbiGrammar -Text ([string]$In['text']) -Grammar $In['grammar']
    if (-not $r['ok']) { return @{ ok = $false; failure = 'grammar_invalid'; message = $r['message']; records = @(); recordCount = 0; unrecognized = 0; names = @() } }
    $warnings = New-Object System.Collections.ArrayList
    $max = [int]$In['maxWarnings']; if ($max -lt 0) { $max = 0 }
    $unrec = @($r['unrecognized'])
    $i = 0
    foreach ($u in $unrec) {
        $i++
        if ($i -gt $max) { break }
        [void]$warnings.Add(@{ code = 'unrecognized_line'; message = ('line ' + $u['line'] + ' not recognised: ' + $u['text']); data = @{ line = $u['line']; text = $u['text'] } })
    }
    if ($unrec.Count -gt $max) { [void]$warnings.Add(@{ code = 'unrecognized_lines'; message = ('' + $unrec.Count + ' line(s) not recognised (' + $max + ' listed above)'); data = @{ total = $unrec.Count } }) }
    foreach ($m in @($r['missing'])) { [void]$warnings.Add(@{ code = 'label_missing'; message = ('label for "' + $m + '" not found in the text'); data = @{ field = $m } }) }
    $records = @($r['records'])
    $names = @(foreach ($rec in $records) { if ($rec.Contains('key')) { [string]$rec['key'] } })
    if ($records.Count -eq 0) { return @{ ok = $false; failure = 'no_records'; message = ('the ' + $r['parser'] + ' grammar recognised no record in ' + (@(Get-EbiGrammarLines -Text ([string]$In['text'])) | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }).Count + ' non-blank line(s)'); records = @(); recordCount = 0; unrecognized = $unrec.Count; names = @(); warnings = $warnings.ToArray() } }
    return @{ ok = $true; records = $records; recordCount = $records.Count; unrecognized = $unrec.Count; names = $names; warnings = $warnings.ToArray() }
}
