# modules/verify/verify.record_counts.ps1
# The HOST team's start message says how many records each file carries
# ("815, 46" for two files). Read that list, and -- once the files are
# here -- hold it against their line counts: a file with a header line has
# one line more, an empty transfer is an empty file or one empty line.
# What comes out is the count of files the message promised (how many
# Receive rows / downloads to expect) and the text the reply carries,
# with the file's own count in brackets where the two disagree.

$Manifest = @{
  id         = 'verify.record_counts'
  group      = 'verify'
  summary    = 'Parse the per-file record counts of a start message; compare with the files'' line counts'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    text    = @{ type='string'; default=''; desc='the counts as the message writes them (numbers, any separators, a unit after each); empty = no message' }
    files   = @{ type='list';   default=@(); desc='file.download / file.list files (lines per file, arrival order); empty = parse only' }
    label   = @{ type='string'; default=''; desc='text before the counts in suffix (e.g. the words for "received data:")' }
    unit    = @{ type='string'; default=''; desc='text after each count (the message''s counter word)' }
    joiner  = @{ type='string'; default=', '; desc='between the counts in text' }
    headerLines = @{ type='int'; default=1; desc='a file may have this many lines more than its record count (a header)' }
  }
  outputs    = @{
    expected  = @{ type='list';   desc='the counts in the message, in order' }
    fileCount = @{ type='int';    desc='how many files the message names; -1 when no message' }
    actual    = @{ type='list';   desc='line count of each file given, in order' }
    text      = @{ type='string'; desc='per file: "<count><unit>", or "<lines><unit>(<count><unit>)" when the file disagrees; the message''s counts alone when no files were given' }
    suffix    = @{ type='string'; desc='" (" + label + text + ")" ready to append to a reply; "" when no message' }
    matched   = @{ type='bool';   desc='every file agrees with its count (true when there is nothing to compare)' }
    reason    = @{ type='string' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.record_counts'; with = @{ text = '{{steps.start.out.fields.counts}}'; files = '{{steps.dl.out.files}}' } }
  notes      = 'A disagreement is a warning (count_disagree), never a failure: the line compare decides whether the files are right; this only says what the reply should show.'
}

function VerifyRecordCounts-Parse {
    # PURE. "815 ken, 46 ken" -> @(815, 46); '' -> @().
    param([string]$Text)
    $out = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrWhiteSpace($Text)) { return $out.ToArray() }
    foreach ($m in [regex]::Matches($Text, '\d+')) { [void]$out.Add([int]$m.Value) }
    return $out.ToArray()
}

function VerifyRecordCounts-Agree {
    # PURE. Does a file of Lines lines carry Count records? (header tolerance, empty transfer = 0 or 1 line)
    param([int]$Lines, [int]$Count, [int]$HeaderLines)
    if ($Lines -lt 0) { return $false }
    if ($Lines -eq $Count) { return $true }
    if ($Lines -gt $Count -and ($Lines - $Count) -le $HeaderLines) { return $true }
    if ($Count -eq 0 -and $Lines -le 1) { return $true }
    return $false
}

function Invoke-Step {
    param($In, $Ctx)
    $exp = @(VerifyRecordCounts-Parse -Text ([string]$In['text']))
    $files = @(@($In['files']) | Where-Object { $_ -is [System.Collections.IDictionary] })
    $unit = [string]$In['unit']; $join = [string]$In['joiner']
    $fileCount = if ([string]::IsNullOrWhiteSpace([string]$In['text'])) { -1 } else { $exp.Count }
    $actual = @(foreach ($f in $files) { if ($f.Contains('lines') -and $null -ne $f['lines']) { [int]$f['lines'] } else { -1 } })
    $parts = New-Object System.Collections.ArrayList
    $why = New-Object System.Collections.ArrayList
    $matched = $true
    if ($files.Count -eq 0) {
        foreach ($e in $exp) { [void]$parts.Add(('' + $e + $unit)) }
    } else {
        $n = [Math]::Max($exp.Count, $actual.Count)
        for ($i = 0; $i -lt $n; $i++) {
            $hasE = ($i -lt $exp.Count); $hasA = ($i -lt $actual.Count)
            if ($hasE -and $hasA -and (VerifyRecordCounts-Agree -Lines ([int]$actual[$i]) -Count ([int]$exp[$i]) -HeaderLines ([int]$In['headerLines']))) { [void]$parts.Add(('' + $exp[$i] + $unit)); continue }
            $matched = $false
            if ($hasE -and $hasA) { [void]$parts.Add(('' + $actual[$i] + $unit + '(' + $exp[$i] + $unit + ')')); [void]$why.Add(('file ' + ($i + 1) + ' (' + [string]$files[$i]['name'] + '): ' + $actual[$i] + ' line(s), the message says ' + $exp[$i])) }
            elseif ($hasA) { [void]$parts.Add(('' + $actual[$i] + $unit + '(?)')); [void]$why.Add(('file ' + ($i + 1) + ' (' + [string]$files[$i]['name'] + ') has no count in the message')) }
            else { [void]$parts.Add(('-(' + $exp[$i] + $unit + ')')); [void]$why.Add(('the message names a file ' + ($i + 1) + ' (' + $exp[$i] + ') that did not arrive')) }
        }
    }
    $text = ($parts -join $join)
    $suffix = if ($fileCount -ge 0 -or ($files.Count -gt 0 -and $exp.Count -gt 0)) { ' (' + [string]$In['label'] + $text + ')' } else { '' }
    $out = @{ ok = $true; expected = $exp; fileCount = $fileCount; actual = $actual; text = $text; suffix = $suffix; matched = $matched; reason = ($why -join '; ') }
    if (-not $matched) { $out['warnings'] = @(@{ code = 'count_disagree'; message = ('the files do not match the start message''s counts: ' + ($why -join '; ')); data = @{ expected = $exp; actual = $actual } }) }
    return $out
}
