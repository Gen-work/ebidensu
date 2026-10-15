# modules/verify/verify.find_token.ps1
# Is the expected name in an OCR'd text? (10/15: the Excel snap copied to
# the GFIX receive-result sheet's B3 was once another job's row; its last
# column, the job column, must be this job's J name.) OCR look-alikes are folded on both sides (O/0,
# I/1/l/|, S/5, B/8, Z/2) and spaces OCR puts inside a word are ignored.
#   ok       the name is there
#   ng       text was read, the name is not there, and no name-like word
#            is one character away from it: another job, most likely
#   unknown  nothing to compare (no text, no expected name), or a word one
#            character off (an OCR misread or another job -- a person looks)
# A CHECK, never a stop: the workflow puts `alert` on the gate in red.

$Manifest = @{
  id         = 'verify.find_token'
  group      = 'verify'
  summary    = 'Find an expected name in OCR text, folding look-alikes; ok / ng / unknown + alert'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    text    = @{ type='string'; default=''; desc='the text to search (screen.ocr text)' }
    expect  = @{ type='string'; default=''; desc='the name that must be in it, e.g. the job''s J name' }
    label   = @{ type='string'; default='the picture'; desc='what the text came from, for the reason / alert wording' }
    pattern = @{ type='string'; default='[A-Za-z0-9|]{5,16}'; desc='regex of a name-like word, for the candidates the reason lists' }
  }
  outputs    = @{
    code       = @{ type='string'; desc='ok | ng | unknown' }
    reason     = @{ type='string'; desc='one line for the gate' }
    alert      = @{ type='string'; desc='the reason when code is ng or a near miss, else empty: the gate draws it red' }
    found      = @{ type='bool' }
    candidates = @{ type='list';   desc='name-like words read from the text (as read, upper-cased)' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.find_token'; with = @{ text = '{{steps.pic1Ocr.out.text}}'; expect = '{{item.JOB}}'; label = 'picture 1 (Excel snap)' } }
  notes      = 'Pure. Folding is symmetric, so a correct name never fails for a look-alike; the cost is that two names differing only by a look-alike pair (O vs 0) are not told apart.'
}

function VerifyFindToken-Fold {
    # PURE. Upper-case, drop whitespace, map OCR look-alikes to one class.
    param([string]$Text)
    $u = ([string]$Text).ToUpperInvariant() -replace '\s+', ''
    $sb = New-Object System.Text.StringBuilder
    foreach ($ch in $u.ToCharArray()) {
        switch ($ch) {
            '0' { [void]$sb.Append('O') }
            '1' { [void]$sb.Append('I') }
            'L' { [void]$sb.Append('I') }
            '|' { [void]$sb.Append('I') }
            '5' { [void]$sb.Append('S') }
            '8' { [void]$sb.Append('B') }
            '2' { [void]$sb.Append('Z') }
            default { [void]$sb.Append($ch) }
        }
    }
    return $sb.ToString()
}

function VerifyFindToken-Distance {
    # PURE. Levenshtein distance of two short strings.
    param([string]$A, [string]$B)
    $n = $A.Length; $m = $B.Length
    if ($n -eq 0) { return $m }
    if ($m -eq 0) { return $n }
    $prev = New-Object 'int[]' ($m + 1)
    $cur = New-Object 'int[]' ($m + 1)
    for ($j = 0; $j -le $m; $j++) { $prev[$j] = $j }
    for ($i = 1; $i -le $n; $i++) {
        $cur[0] = $i
        for ($j = 1; $j -le $m; $j++) {
            $cost = if ($A[$i - 1] -eq $B[$j - 1]) { 0 } else { 1 }
            $cur[$j] = [Math]::Min([Math]::Min($prev[$j] + 1, $cur[$j - 1] + 1), $prev[$j - 1] + $cost)
        }
        $tmp = $prev; $prev = $cur; $cur = $tmp
    }
    return $prev[$m]
}

function VerifyFindToken-Judge {
    # PURE. text + expected name -> the step's outputs (without ok).
    param([string]$Text, [string]$Expect, [string]$Label = 'the picture', [string]$Pattern = '[A-Za-z0-9|]{5,16}')
    $cands = New-Object System.Collections.ArrayList
    foreach ($m in [regex]::Matches([string]$Text, $Pattern)) {
        $w = $m.Value.ToUpperInvariant()
        if ($w -match '[A-Z]' -and -not ($cands -contains $w)) { [void]$cands.Add($w) }
    }
    $e = ([string]$Expect).Trim()
    if ($e -eq '') { return @{ code = 'unknown'; reason = ('nothing to look for in ' + $Label + ' (no expected name on this row)'); alert = ''; found = $false; candidates = $cands.ToArray() } }
    if (([string]$Text).Trim() -eq '') { return @{ code = 'unknown'; reason = ('no text read from ' + $Label + ': check by eye that it shows ' + $e); alert = ''; found = $false; candidates = $cands.ToArray() } }
    $fe = VerifyFindToken-Fold $e
    if ((VerifyFindToken-Fold $Text).Contains($fe)) { return @{ code = 'ok'; reason = ($Label + ' shows ' + $e); alert = ''; found = $true; candidates = $cands.ToArray() } }
    $near = @($cands | Where-Object { (VerifyFindToken-Distance (VerifyFindToken-Fold $_) $fe) -le 1 })
    $seen = if ($cands.Count -gt 0) { (@($cands | Select-Object -First 6) -join ', ') } else { '(no name-like word)' }
    if ($near.Count -gt 0) {
        $r = ($Label + ' shows ' + $near[0] + ', one character off ' + $e + ': an OCR misread or another job -- look')
        return @{ code = 'unknown'; reason = $r; alert = $r; found = $false; candidates = $cands.ToArray() }
    }
    $r = ($Label + ' does NOT show ' + $e + ' (read: ' + $seen + '): another job''s picture?')
    return @{ code = 'ng'; reason = $r; alert = $r; found = $false; candidates = $cands.ToArray() }
}

function Invoke-Step {
    param($In, $Ctx)
    try { [void][regex]::new([string]$In['pattern']) } catch { return @{ ok = $false; failure = 'input_invalid'; message = ('pattern is not a regex: ' + $_.Exception.Message); code = 'unknown'; reason = ''; alert = ''; found = $false; candidates = @() } }
    $j = VerifyFindToken-Judge -Text ([string]$In['text']) -Expect ([string]$In['expect']) -Label ([string]$In['label']) -Pattern ([string]$In['pattern'])
    if ($j['alert'] -ne '') { $Ctx.Log.Warn($j['alert']) }
    return @{ ok = $true; code = $j['code']; reason = $j['reason']; alert = $j['alert']; found = [bool]$j['found']; candidates = @($j['candidates']) }
}
