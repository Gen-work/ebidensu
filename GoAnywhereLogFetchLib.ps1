# ============================================================
#  GoAnywhereLogFetchLib.ps1
#
#  PURE helpers behind the standalone GoAnywhereLogFetch.ps1 driver
#  (one-off: fetch the GoAnywhere receive-job logs for a hand-written
#  list of correls and copy each into its receive system's LOG folder).
#  No COM, no SendKeys, no file I/O except Read-GoAnywhereLogFetchConfig
#  and Find-GoAnywhereLogOwner's log reads. Dot-source only (no param()).
#
#  Identity rules (operator decisions, 2026-10-20):
#    - A job is aimed at by its IF number ONLY: the project name may say
#      VER1 or DEV1, and the sender prefix varies (LLPS / LALC / ...), so
#      neither is matched. 'IF0001' must not hit 'IF00010' -- the digits
#      are bounded.
#    - The receive side is the job's Folder (/<user>/Receive), set per
#      receive system in the config (goFolder).
#    - The correl's LAST character is free: JIDSL25M and JIDSL25S are the
#      same correl. Log ownership is decided from the log CONTENT
#      ('/gfix/recv/<first 7 chars><any one char>'), never from which row
#      was opened, because one IF number can feed several receive jobs.
# ============================================================

# GoAnywhere completed-jobs LIST page text (Ctrl+A / Ctrl+C) -> rows.
# Column-order tolerant: the page has been seen both with and without a
# User column and with a trailing 'resolved' column, so the folder is the
# first field starting with '/' and the times are the datetime-shaped
# fields, not fixed positions.
function ConvertFrom-GoAnywhereJobList {
    param([string]$Text)
    $rows = [System.Collections.Generic.List[object]]::new()
    if ([string]::IsNullOrEmpty($Text)) { return $rows.ToArray() }
    foreach ($line in [regex]::Split($Text, '\r?\n')) {
        $parts = @($line.Trim() -split "`t" | ForEach-Object { $_.Trim() })
        if ($parts.Count -lt 4) { continue }
        if ($parts[0] -notmatch '^\d{6,}$') { continue }
        $folder = ''
        $times = [System.Collections.Generic.List[string]]::new()
        for ($i = 2; $i -lt $parts.Count; $i++) {
            $p = $parts[$i]
            if ($folder -eq '' -and $p.StartsWith('/')) { $folder = $p; continue }
            if ($p -match '^\d{4}-\d{2}-\d{2} \d{1,2}:\d{2}:\d{2}$') { $times.Add($p) }
        }
        $rows.Add([pscustomobject]@{
            JobNo       = $parts[0]
            ProjectName = $parts[1]
            Folder      = $folder
            StartTime   = $(if ($times.Count -gt 0) { $times[0] } else { '' })
            EndTime     = $(if ($times.Count -gt 1) { $times[1] } else { '' })
        })
    }
    return $rows.ToArray()
}

# '1' / '0001' / 'IF0001' -> '0001'. '' when there are no digits.
function ConvertTo-GoAnywhereIfNo {
    param([string]$Value)
    $m = [regex]::Match([string]$Value, '\d+')
    if (-not $m.Success) { return '' }
    return $m.Value.PadLeft(4, '0')
}

# '0001,0002' / '1 2' / array -> @('0001','0002'): one target may be aimed
# at by several IF numbers (e.g. received under one, resent under another).
function ConvertTo-GoAnywhereIfNos {
    param($Value)
    $out = [System.Collections.Generic.List[string]]::new()
    foreach ($v in @($Value)) {
        foreach ($part in ([string]$v -split '[,;/\s]+')) {
            $n = ConvertTo-GoAnywhereIfNo $part
            if ($n -ne '' -and -not $out.Contains($n)) { $out.Add($n) }
        }
    }
    return $out.ToArray()
}

# True when the project name carries exactly this IF number
# (IF0001 matches LLPS_DEV1_IF0001_001_SSH, not ..._IF00010_...).
function Test-GoAnywhereIfMatch {
    param([string]$ProjectName, [string]$IfNo)
    $n = ConvertTo-GoAnywhereIfNo $IfNo
    if ($n -eq '') { return $false }
    $digits = $n.TrimStart('0')
    if ($digits -eq '') { $digits = '0' }
    return ([string]$ProjectName -match ('(?i)IF0*' + [regex]::Escape($digits) + '(?!\d)'))
}

# Receive-side rows for one target (IF number(s) + receive folder).
function Get-GoAnywhereCandidateRows {
    param([object[]]$Rows, [string]$IfNo, [string]$GoFolder)
    $out = [System.Collections.Generic.List[object]]::new()
    $ifNos = @(ConvertTo-GoAnywhereIfNos $IfNo)
    foreach ($r in @($Rows)) {
        $hit = $false
        foreach ($n in $ifNos) { if (Test-GoAnywhereIfMatch $r.ProjectName $n) { $hit = $true; break } }
        if (-not $hit) { continue }
        $f = [string]$r.Folder
        if ($f -notmatch '(?i)Receive') { continue }
        if (-not [string]::IsNullOrWhiteSpace($GoFolder) -and
            -not $f.StartsWith($GoFolder, [StringComparison]::OrdinalIgnoreCase)) { continue }
        $out.Add($r)
    }
    return $out.ToArray()
}

# Merge row lists from several captures; JobNo is the identity, first wins.
function Merge-GoAnywhereRows {
    param([object[]]$Existing, [object[]]$New)
    $seen = @{}
    $out = [System.Collections.Generic.List[object]]::new()
    foreach ($r in @($Existing) + @($New)) {
        if ($null -eq $r) { continue }
        if ($seen.ContainsKey($r.JobNo)) { continue }
        $seen[$r.JobNo] = $true
        $out.Add($r)
    }
    return $out.ToArray()
}

# The plan: per target its candidate rows (newest first), the distinct job
# numbers to download (first-seen order) and the targets with no candidate.
# Targets: objects with Correl, IfNo, Recv. Systems: hashtable
# Recv -> @{ goFolder; dest }.
function Get-GoAnywhereFetchPlan {
    param([object[]]$Targets, [object[]]$Rows, [hashtable]$Systems)
    $items = [System.Collections.Generic.List[object]]::new()
    $jobs = [System.Collections.Generic.List[string]]::new()
    $seen = @{}
    foreach ($t in @($Targets)) {
        $sys = $Systems[[string]$t.Recv]
        $goFolder = ''
        if ($null -ne $sys) { $goFolder = [string]$sys.goFolder }
        # an unknown receive system has no folder to aim at -> no candidates
        # (an empty goFolder would otherwise accept every Receive folder)
        $cand = @()
        if ($null -ne $sys -and -not [string]::IsNullOrWhiteSpace($goFolder)) {
            $cand = @(Get-GoAnywhereCandidateRows -Rows $Rows -IfNo $t.IfNo -GoFolder $goFolder |
                Sort-Object -Property @{ Expression = { $_.StartTime } ; Descending = $true }, @{ Expression = { $_.JobNo } ; Descending = $true })
        }
        foreach ($c in $cand) {
            if (-not $seen.ContainsKey($c.JobNo)) { $seen[$c.JobNo] = $true; $jobs.Add($c.JobNo) }
        }
        $items.Add([pscustomobject]@{
            Correl     = [string]$t.Correl
            IfNo       = (@(ConvertTo-GoAnywhereIfNos $t.IfNo) -join ',')
            Recv       = [string]$t.Recv
            KnownRecv  = ($null -ne $sys)
            Candidates = $cand
        })
    }
    return [pscustomobject]@{
        Items      = $items.ToArray()
        JobNumbers = $jobs.ToArray()
        Missing    = @($items | Where-Object { $_.Candidates.Count -eq 0 })
    }
}

# The receive command path names the correl; its last character is free
# (M / S / ...), an appended transfer stamp (.yyMMdd.HHmmssff) is allowed.
function Get-GoAnywhereCorrelPattern {
    param([string]$Correl)
    $c = ([string]$Correl).Trim()
    $base = $c
    if ($c.Length -ge 2) { $base = $c.Substring(0, $c.Length - 1) }
    return ('(?i)/gfix/recv/' + [regex]::Escape($base) + '[A-Za-z0-9](?:\.\d{6}\.\d{8})?(?=[\s''"]|$)')
}

function Test-GoAnywhereLogNamesCorrel {
    param([string]$Text, [string]$Correl)
    if ([string]::IsNullOrEmpty($Text)) { return $false }
    return [regex]::IsMatch($Text, (Get-GoAnywhereCorrelPattern $Correl), [System.Text.RegularExpressions.RegexOptions]::Multiline)
}

# Decide which of a correl's candidate logs it owns.
#   Logs: objects with JobNo, StartTime, Text (candidates only, any order).
#   ok        -> content names the correl; Chosen = newest such log,
#                Others = older ones that also name it (reruns).
#   nocontent -> logs present but none names the correl; Chosen = $null.
#   nolog     -> nothing downloaded for this correl.
function Select-GoAnywhereLogForCorrel {
    param([string]$Correl, [object[]]$Logs)
    $logs = @($Logs | Where-Object { $null -ne $_ })
    if ($logs.Count -eq 0) { return [pscustomobject]@{ Status = 'nolog'; Chosen = $null; Others = @() } }
    $hits = @($logs | Where-Object { Test-GoAnywhereLogNamesCorrel $_.Text $Correl } |
        Sort-Object -Property @{ Expression = { $_.StartTime } ; Descending = $true }, @{ Expression = { $_.JobNo } ; Descending = $true })
    if ($hits.Count -eq 0) { return [pscustomobject]@{ Status = 'nocontent'; Chosen = $null; Others = @() } }
    return [pscustomobject]@{ Status = 'ok'; Chosen = $hits[0]; Others = @($hits | Select-Object -Skip 1) }
}

# Destination file name from a template: {file} (the downloaded file's own
# name, as is -- the default), {orig} (the same without extension),
# {correl} {jobNo} {start} (yyyyMMdd_HHmmss).
function Format-GoAnywhereLogName {
    param([string]$Format, [string]$Correl, [string]$JobNo, [string]$StartTime, [string]$OrigName)
    if ([string]::IsNullOrWhiteSpace($Format)) { $Format = '{file}' }
    $start = ([string]$StartTime -replace '[^\d]', '')
    if ($start.Length -ge 14) { $start = $start.Substring(0, 8) + '_' + $start.Substring(8, 6) }
    $orig = [IO.Path]::GetFileNameWithoutExtension([string]$OrigName)
    $name = $Format.Replace('{file}', [IO.Path]::GetFileName([string]$OrigName)).Replace('{correl}', $Correl).Replace('{jobNo}', $JobNo).Replace('{start}', $start).Replace('{orig}', $orig)
    foreach ($ch in [IO.Path]::GetInvalidFileNameChars()) { $name = $name.Replace([string]$ch, '_') }
    return $name
}

# Config JSON (UTF-8, may hold Japanese paths) -> normalized hashtable.
#   { "destRoot": "...\\LOG", "nameFormat": "{file}",
#     "systems": { "IDS": { "goFolder": "/idsusr/", "dest": "IDS" }, ... },
#     "jobs": [ { "correl": "...", "if": "0001" | "0001,0002", "recv": "IDS" } ] }
function Read-GoAnywhereLogFetchConfig {
    param([string]$Path)
    $raw = [IO.File]::ReadAllText($Path, [Text.Encoding]::UTF8)
    $j = $raw | ConvertFrom-Json
    $systems = @{}
    foreach ($p in @($j.systems.PSObject.Properties)) {
        $systems[$p.Name] = @{ goFolder = [string]$p.Value.goFolder; dest = [string]$p.Value.dest }
    }
    $targets = [System.Collections.Generic.List[object]]::new()
    foreach ($x in @($j.jobs)) {
        $targets.Add([pscustomobject]@{ Correl = ([string]$x.correl).Trim(); IfNo = (@(ConvertTo-GoAnywhereIfNos $x.if) -join ','); Recv = ([string]$x.recv).Trim() })
    }
    $fmt = ''
    if ($j.PSObject.Properties.Name -contains 'nameFormat') { $fmt = [string]$j.nameFormat }
    return @{
        DestRoot   = [string]$j.destRoot
        NameFormat = $fmt
        Systems    = $systems
        Targets    = $targets.ToArray()
    }
}
