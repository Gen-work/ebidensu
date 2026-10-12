# modules/verify/verify.derive_fields.ps1
# Add fields to every record: a constant, a copy of another field, or a
# regex rewrite of another field. What a workflow needs to turn rows it
# read into rows a worklist takes, without arithmetic in the workflow --
# e.g. the GFIX-side name from the GIFT-side job name (FJDSJM40 ->
# FJDSWM40: the 5th character J -> W) and today's date on every row.

$Manifest = @{
  id         = 'verify.derive_fields'
  group      = 'verify'
  summary    = 'Add fields to records: a constant, a copy, or a regex rewrite of another field'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    records = @{ type='list'; required=$true; desc='records (excel.read_rows / verify.parse_text output)' }
    set     = @{ type='list'; required=$true; desc='@{ to; value } | @{ to; from } | @{ to; from; pattern; replace } (.NET regex, $1 / ${name} in replace), applied in order' }
  }
  outputs    = @{
    records   = @{ type='list'; desc='copies of the input records with the fields set' }
    unchanged = @{ type='list'; desc='"<to> <- <from>: <value>" for every rewrite whose pattern did not match (the value was copied as is)' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.derive_fields'; with = @{ records = '{{steps.today.out.records}}'; set = @(@{ to = 'Excel_NAME'; from = 'job'; pattern = '^(.{4})J'; replace = '${1}W' }, @{ to = 'GFIX_DATE'; value = '{{run.date}}' }) } }
  notes      = 'A rule naming a field the record lacks reads it as empty. A pattern that does not match copies the value unchanged and lists it in unchanged (a warning), so a name that does not follow the rule is visible instead of silently wrong.'
}

function Invoke-Step {
    param($In, $Ctx)
    $rules = @(@($In['set']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    foreach ($r in $rules) {
        if ([string]::IsNullOrWhiteSpace([string]$r['to'])) { return @{ ok = $false; failure = 'input_invalid'; message = 'every set rule needs "to"'; records = @(); unchanged = @() } }
        if (-not ($r.Contains('value') -or $r.Contains('from'))) { return @{ ok = $false; failure = 'input_invalid'; message = ('set rule for "' + $r['to'] + '" needs "value" or "from"'); records = @(); unchanged = @() } }
        if ($r.Contains('pattern')) {
            try { [void][regex]::new([string]$r['pattern']) } catch { return @{ ok = $false; failure = 'input_invalid'; message = ('set rule for "' + $r['to'] + '": bad pattern: ' + $_.Exception.Message); records = @(); unchanged = @() } }
        }
    }
    $out = New-Object System.Collections.ArrayList
    $unchanged = New-Object System.Collections.ArrayList
    foreach ($rec in @($In['records'])) {
        if (-not ($rec -is [System.Collections.IDictionary])) { continue }
        $copy = @{}
        foreach ($k in $rec.Keys) { $copy[$k] = $rec[$k] }
        foreach ($r in $rules) {
            $to = [string]$r['to']
            if ($r.Contains('value')) { $copy[$to] = $r['value']; continue }
            $from = [string]$r['from']
            $v = if ($copy.Contains($from) -and $null -ne $copy[$from]) { [string]$copy[$from] } else { '' }
            if ($r.Contains('pattern')) {
                $re = [regex]::new([string]$r['pattern'])
                if ($re.IsMatch($v)) { $v = $re.Replace($v, [string]$r['replace'], 1) }
                else { [void]$unchanged.Add(($to + ' <- ' + $from + ': ' + $v)) }
            }
            $copy[$to] = $v
        }
        [void]$out.Add($copy)
    }
    $res = @{ ok = $true; records = $out.ToArray(); unchanged = $unchanged.ToArray() }
    if ($unchanged.Count -gt 0) { $res['warnings'] = @(@{ code = 'pattern_unmatched'; message = ('copied as is (pattern did not match): ' + ($unchanged -join '; ')) }) }
    return $res
}
