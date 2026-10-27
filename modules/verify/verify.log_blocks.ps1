# modules/verify/verify.log_blocks.ps1
# A receive log pasted on an evidence sheet, checked as text: every
# "...Receiver END <id>" has its "...Receiver START <id>" and the other way
# round, and the files the log says it stored include the files the DF
# compares. Two misses of 10/26 - 10/27 the review caught by hand:
#   QJODWAF1 / SJODWAF1: GIFTReceive.log on the GIFT receive-result sheet carried both ENDs
#     but only one START (K-37);
#   LJODWAF1: the DF's GIFT side was F202609280003.csv, KJODWAF1's empty
#     file in the wrong folder, while the GIFT log stored ...0004 / ...0005
#     (an empty file compares equal to any other empty file, so the DF said
#     "the same").
# A CHECK, never a stop: the workflow puts `alert` on the gate in red.

$Manifest = @{
  id         = 'verify.log_blocks'
  group      = 'verify'
  summary    = 'Check a pasted receive log: START / END pairs, stored files include the compared files'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    lines        = @{ type='list';   default=@(); desc='the log lines: strings, or records (excel.read_rows) whose field holds the line' }
    field        = @{ type='string'; default='line'; desc='records: the field with the text' }
    startPattern = @{ type='string'; default='Receiver\s+START\s+(\S+)'; desc='regex; group 1 = the transfer id' }
    endPattern   = @{ type='string'; default='Receiver\s+END\s+(\S+)'; desc='regex; group 1 = the transfer id' }
    filePattern  = @{ type='string'; default='file stored : \[[^\]]*?([^/\\\]]+)\]'; desc='regex; group 1 = a stored file name' }
    pairs        = @{ type='list';   default=@(); desc='the DF pairs (track.json pairs); their side file names must be among the stored files' }
    side         = @{ type='string'; default='left'; enum=@('left', 'right'); desc='which file of a pair the log stored (left = GIFT)' }
    label        = @{ type='string'; default='the log'; desc='what the lines are, for the reason / alert wording' }
  }
  outputs    = @{
    code     = @{ type='string'; desc='ok | ng | unknown (no lines, or no START / END / stored file read)' }
    reason   = @{ type='string'; desc='one line for the gate' }
    alert    = @{ type='string'; desc='the reason when code is ng, else empty: the gate draws it red' }
    starts   = @{ type='list';   desc='transfer ids with a START line, in order' }
    ends     = @{ type='list';   desc='transfer ids with an END line, in order' }
    unpaired = @{ type='list';   desc='"END <id> without START" / "START <id> without END"' }
    files    = @{ type='list';   desc='file names the log stored, in order' }
    missing  = @{ type='list';   desc='compared file names the log did not store' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.log_blocks'; with = @{ lines = '{{steps.giftLogRows.out.records}}'; pairs = '{{steps.track.out.data.pairs}}'; label = 'GIFTReceive.log (GIFT receive result)' } }
  notes      = 'Pure. A pair whose side has no name (no leftName / left path) is not checked, nor one whose extension the log stored no file of (a ZIP transfer stores the .zip). Stored files are compared by file name, case-insensitive.'
}

function VerifyLogBlocks-Text {
    # PURE. A line: the string, or the record's field.
    param($Line, [string]$Field)
    if ($Line -is [System.Collections.IDictionary]) { if ($Line.Contains($Field) -and $null -ne $Line[$Field]) { return [string]$Line[$Field] }; return '' }
    if ($null -eq $Line) { return '' }
    return [string]$Line
}

function VerifyLogBlocks-PairName {
    # PURE. The side's file name of a pair: <side>Name, else the leaf of <side>.
    param($Pair, [string]$Side)
    if (-not ($Pair -is [System.Collections.IDictionary])) { return '' }
    $n = $Side + 'Name'
    if ($Pair.Contains($n) -and -not [string]::IsNullOrWhiteSpace([string]$Pair[$n])) { return ([string]$Pair[$n]).Trim() }
    if ($Pair.Contains($Side) -and -not [string]::IsNullOrWhiteSpace([string]$Pair[$Side])) { return (([string]$Pair[$Side]) -split '[\\/]')[-1].Trim() }
    return ''
}

function VerifyLogBlocks-Judge {
    # PURE. -> the step's outputs (without ok).
    param($Lines, [string]$Field, [string]$StartPattern, [string]$EndPattern, [string]$FilePattern, $Pairs, [string]$Side, [string]$Label)
    $starts = New-Object System.Collections.ArrayList
    $ends = New-Object System.Collections.ArrayList
    $files = New-Object System.Collections.ArrayList
    $n = 0
    foreach ($l in @($Lines)) {
        $t = VerifyLogBlocks-Text -Line $l -Field $Field
        if ($t.Trim() -eq '') { continue }
        $n++
        $m = [regex]::Match($t, $StartPattern); if ($m.Success) { [void]$starts.Add($m.Groups[1].Value) }
        $m = [regex]::Match($t, $EndPattern); if ($m.Success) { [void]$ends.Add($m.Groups[1].Value) }
        $m = [regex]::Match($t, $FilePattern); if ($m.Success -and -not ($files -contains $m.Groups[1].Value)) { [void]$files.Add($m.Groups[1].Value) }
    }
    $unpaired = New-Object System.Collections.ArrayList
    foreach ($e in $ends) { if (-not ($starts -contains $e)) { [void]$unpaired.Add('END ' + $e + ' without START') } }
    foreach ($s in $starts) { if (-not ($ends -contains $s)) { [void]$unpaired.Add('START ' + $s + ' without END') } }
    $missing = New-Object System.Collections.ArrayList
    if ($files.Count -gt 0) {
        foreach ($p in @($Pairs)) {
            $name = VerifyLogBlocks-PairName -Pair $p -Side $Side
            if ($name -eq '' -or $missing -contains $name) { continue }
            # a ZIP transfer stores the .zip, the DF compares what was in it:
            # only names whose extension the log stored are checked
            $ext = [System.IO.Path]::GetExtension($name)
            if (@($files | Where-Object { [System.IO.Path]::GetExtension($_) -ieq $ext }).Count -eq 0) { continue }
            if (@($files | Where-Object { $_ -ieq $name }).Count -eq 0) { [void]$missing.Add($name) }
        }
    }
    $out = @{ starts = $starts.ToArray(); ends = $ends.ToArray(); unpaired = $unpaired.ToArray(); files = $files.ToArray(); missing = $missing.ToArray() }
    if ($n -eq 0) { return $out + @{ code = 'unknown'; reason = ('no lines in ' + $Label); alert = '' } }
    $bad = New-Object System.Collections.ArrayList
    if ($unpaired.Count -gt 0) { [void]$bad.Add(($Label + ': ' + ($unpaired.ToArray() -join '; ') + ' -- a START / END line is missing (K-37)')) }
    if ($missing.Count -gt 0) { [void]$bad.Add(($Label + ' stored ' + ($files.ToArray() -join ', ') + ' but the DF compares ' + ($missing.ToArray() -join ', ') + ': another job''s file in the DATA folder? (an empty file matches any empty file)')) }
    if ($bad.Count -gt 0) { $r = $bad.ToArray() -join ' / '; return $out + @{ code = 'ng'; reason = $r; alert = $r } }
    if ($starts.Count -eq 0 -or $files.Count -eq 0) { return $out + @{ code = 'unknown'; reason = ($Label + ': ' + $(if ($starts.Count -eq 0) { 'no START / END line read' } else { 'no stored file read' }) + ' -- check by eye'); alert = '' } }
    return $out + @{ code = 'ok'; reason = ($Label + ': ' + $starts.Count + ' START / END pair(s), stored ' + ($files.ToArray() -join ', ')); alert = '' }
}

function Invoke-Step {
    param($In, $Ctx)
    foreach ($k in @('startPattern', 'endPattern', 'filePattern')) {
        try { [void][regex]::new([string]$In[$k]) } catch { return @{ ok = $false; failure = 'input_invalid'; message = ($k + ' is not a regex: ' + $_.Exception.Message); code = 'unknown'; reason = ''; alert = ''; starts = @(); ends = @(); unpaired = @(); files = @(); missing = @() } }
    }
    $j = VerifyLogBlocks-Judge -Lines $In['lines'] -Field ([string]$In['field']) -StartPattern ([string]$In['startPattern']) -EndPattern ([string]$In['endPattern']) -FilePattern ([string]$In['filePattern']) -Pairs $In['pairs'] -Side ([string]$In['side']) -Label ([string]$In['label'])
    if ($j['alert'] -ne '') { $Ctx.Log.Warn($j['alert']) }
    return @{ ok = $true; code = $j['code']; reason = $j['reason']; alert = $j['alert']; starts = @($j['starts']); ends = @($j['ends']); unpaired = @($j['unpaired']); files = @($j['files']); missing = @($j['missing']) }
}
