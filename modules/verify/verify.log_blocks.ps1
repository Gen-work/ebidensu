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
#
# 10/28: and a repair. track.json keeps the pairs of the day it ran; a file
# put right in DATA\GIFT afterwards is never seen again (L's second NG, K-53).
# A pair whose side the log did not store gets the one stored file of the
# same folder, extension and line count that no other pair uses -- the
# pairs / shots outputs carry it, its shot set to null so the DF is taken
# again. Anything less certain is left for the gate.

. (Join-Path $PSScriptRoot '..\..\kernel\LogText.ps1')   # Get-EbiByteLineCount

$Manifest = @{
  id         = 'verify.log_blocks'
  group      = 'verify'
  summary    = 'Check a pasted receive log: START / END pairs, stored files include the compared files'
  tier       = 'core'
  effects    = 'read'
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
    shots        = @{ type='list';   default=@(); desc='the DF shots of the pairs, in the same order (track.json dfShots): a repaired pair''s shot comes back null (take it again)' }
    repair       = @{ type='bool';   default=$true; desc='put a compared file the log did not store right, when exactly one stored file fits (same folder, extension, line count, unused)' }
  }
  outputs    = @{
    code     = @{ type='string'; desc='ok | ng | unknown (no lines, or no START / END / stored file read)' }
    reason   = @{ type='string'; desc='one line for the gate' }
    alert    = @{ type='string'; desc='the reason when code is ng, else empty: the gate draws it red' }
    starts   = @{ type='list';   desc='transfer ids with a START line, in order' }
    ends     = @{ type='list';   desc='transfer ids with an END line, in order' }
    unpaired = @{ type='list';   desc='"END <id> without START" / "START <id> without END"' }
    files    = @{ type='list';   desc='file names the log stored, in order' }
    missing  = @{ type='list';   desc='compared file names the log did not store (after the repair)' }
    pairs    = @{ type='list';   desc='the pairs, with repaired sides: feed these to the DF steps' }
    shots    = @{ type='list';   desc='the shots input, null where the pair was repaired' }
    repaired = @{ type='list';   desc='"<old> -> <new>" per repaired side' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'verify.log_blocks'; with = @{ lines = '{{steps.giftLogRows.out.records}}'; pairs = '{{steps.track.out.data.pairs}}'; label = 'GIFTReceive.log (GIFT receive result)' } }
  notes      = 'Reads only the stored files a repair would use (exists, line count); writes nothing. A pair whose side has no name (no leftName / left path) is not checked, nor one whose extension the log stored no file of (a ZIP transfer stores the .zip). Stored files are compared by file name, case-insensitive.'
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

function VerifyLogBlocks-Repair {
    <#
      PURE. -> @{ pairs; shots; repaired }. For each pair whose side name is
      in Missing: the candidates are the stored Files with the same
      extension that no pair uses, in the side's folder, whose Probe line
      count (path -> lines, -1 = no such file) equals the pair's lines.
      Exactly one candidate -> the side takes it and the pair's shot becomes
      null; none or several -> left as it is.
    #>
    param($Pairs, $Shots, $Missing, $Files, [string]$Side, $Probe)
    $out = New-Object System.Collections.ArrayList
    foreach ($p in @($Pairs)) { if ($p -is [System.Collections.IDictionary]) { $c = @{}; foreach ($k in $p.Keys) { $c[$k] = $p[$k] }; [void]$out.Add($c) } else { [void]$out.Add($p) } }
    $shotsOut = New-Object System.Collections.ArrayList
    foreach ($s in @($Shots)) { [void]$shotsOut.Add($s) }
    $repaired = New-Object System.Collections.ArrayList
    for ($i = 0; $i -lt $out.Count; $i++) {
        $p = $out[$i]
        $name = VerifyLogBlocks-PairName -Pair $p -Side $Side
        if ($name -eq '' -or -not (@($Missing) -contains $name)) { continue }
        $path = if ($p.Contains($Side)) { [string]$p[$Side] } else { '' }
        if ($path -eq '' -or -not $p.Contains('lines') -or $null -eq $p['lines']) { continue }
        $dir = [System.IO.Path]::GetDirectoryName($path)
        $ext = [System.IO.Path]::GetExtension($name)
        $used = @($out | ForEach-Object { VerifyLogBlocks-PairName -Pair $_ -Side $Side })
        $cands = @(foreach ($f in @($Files)) {
            if ([System.IO.Path]::GetExtension($f) -ine $ext -or (@($used | Where-Object { $_ -ieq $f }).Count -gt 0)) { continue }
            $cp = [System.IO.Path]::Combine($dir, $f)
            if ($null -ne $Probe -and $Probe.Contains($cp) -and [int]$Probe[$cp] -eq [int]$p['lines']) { $cp }
        })
        if ($cands.Count -ne 1) { continue }
        $new = [System.IO.Path]::GetFileName($cands[0])
        $p[$Side] = $cands[0]; $p[$Side + 'Name'] = $new
        if ($i -lt $shotsOut.Count) { $shotsOut[$i] = $null }
        [void]$repaired.Add($name + ' -> ' + $new)
    }
    return @{ pairs = $out.ToArray(); shots = $shotsOut.ToArray(); repaired = $repaired.ToArray() }
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
        try { [void][regex]::new([string]$In[$k]) } catch { return @{ ok = $false; failure = 'input_invalid'; message = ($k + ' is not a regex: ' + $_.Exception.Message); code = 'unknown'; reason = ''; alert = ''; starts = @(); ends = @(); unpaired = @(); files = @(); missing = @(); pairs = @($In['pairs']); shots = @($In['shots']); repaired = @() } }
    }
    $side = [string]$In['side']
    $j = VerifyLogBlocks-Judge -Lines $In['lines'] -Field ([string]$In['field']) -StartPattern ([string]$In['startPattern']) -EndPattern ([string]$In['endPattern']) -FilePattern ([string]$In['filePattern']) -Pairs $In['pairs'] -Side $side -Label ([string]$In['label'])
    # @($null) would be one null shot: an absent list stays empty
    $pairsIn = if ($null -eq $In['pairs']) { @() } else { @($In['pairs']) }
    $shotsIn = if ($null -eq $In['shots']) { @() } else { @($In['shots']) }
    $fix = @{ pairs = $pairsIn; shots = $shotsIn; repaired = @() }
    $warnings = @()
    if ([bool]$In['repair'] -and @($j['missing']).Count -gt 0) {
        # what a repair could use: each stored file in each missing side's folder
        $probe = @{}
        foreach ($p in $pairsIn) {
            if (-not ($p -is [System.Collections.IDictionary]) -or -not $p.Contains($side) -or [string]$p[$side] -eq '') { continue }
            $dir = [System.IO.Path]::GetDirectoryName([string]$p[$side])
            foreach ($f in @($j['files'])) {
                $cp = [System.IO.Path]::Combine($dir, $f)
                if ($probe.Contains($cp)) { continue }
                $probe[$cp] = $(if ([System.IO.File]::Exists($cp)) { Get-EbiByteLineCount -Bytes ([System.IO.File]::ReadAllBytes($cp)) } else { -1 })
            }
        }
        $fix = VerifyLogBlocks-Repair -Pairs $pairsIn -Shots $shotsIn -Missing $j['missing'] -Files $j['files'] -Side $side -Probe $probe
        if (@($fix['repaired']).Count -gt 0) {
            $j = VerifyLogBlocks-Judge -Lines $In['lines'] -Field ([string]$In['field']) -StartPattern ([string]$In['startPattern']) -EndPattern ([string]$In['endPattern']) -FilePattern ([string]$In['filePattern']) -Pairs $fix['pairs'] -Side $side -Label ([string]$In['label'])
            $msg = ([string]$In['label'] + ': the DF compared ' + (@($fix['repaired']) -join ', ') + ' -- the file the log stored (same folder, same lines); its DF is taken again. track.json still has the old pair (K-53)')
            $j['reason'] = $msg + $(if ([string]$j['reason'] -ne '') { ' / ' + [string]$j['reason'] } else { '' })
            $warnings = @(@{ code = 'pair_repaired'; message = $msg; data = @{ repaired = @($fix['repaired']) } })
        }
    }
    if ($j['alert'] -ne '') { $Ctx.Log.Warn($j['alert']) }
    $r = @{ ok = $true; code = $j['code']; reason = $j['reason']; alert = $j['alert']; starts = @($j['starts']); ends = @($j['ends']); unpaired = @($j['unpaired']); files = @($j['files']); missing = @($j['missing']); pairs = @($fix['pairs']); shots = @($fix['shots']); repaired = @($fix['repaired']) }
    if ($warnings.Count -gt 0) { $r['warnings'] = $warnings }
    return $r
}
