#Requires -Version 5.1
# ============================================================
#  kernel/Mask.ps1
#
#  mask-lite (P2-08): the rule half of Plan.md section 7, enough for a CI
#  gate. `ebi mask check` scans profiles/**/fixtures/ (and any other tracked
#  text the caller lists) for the high-risk shapes -- employee ids, mail
#  addresses on a company domain, UNC paths, C:\Users\<id>, intranet URLs,
#  words from a dictionary -- and fails on the first hit. The interactive
#  decisions and the stable bijective replacement stay in P5 (kernel/Mask
#  grows there); this file only finds.
#
#  Dot-source only (no param(), ASCII source). Pure except Find-EbiMask
#  Hits' file reads.
#
#  Rules are data: Get-EbiMaskRules returns @{ id; pattern; note }; the
#  dictionary is profiles/mask-dictionary.json when present (a JSON list
#  of literal words: company names, project code names, people), and the
#  same file lists `allow` patterns for known-safe hits.
# ============================================================

. (Join-Path $PSScriptRoot 'Json.ps1')

function Get-EbiMaskRules {
    return @(
        @{ id = 'employee_id'; pattern = '\b[A-Z]{2}\d{6}\b';                                              note = 'employee id (two letters + six digits)' }
        @{ id = 'email';       pattern = '\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Za-z]{2,}\b';                note = 'mail address' }
        @{ id = 'unc_path';    pattern = '\\\\[A-Za-z0-9._-]+\\[A-Za-z0-9$._-]+';                             note = 'UNC path \\host\share' }
        @{ id = 'user_home';   pattern = '(?i)[A-Z]:\\Users\\[^\\\s"<>|]+';                                    note = 'C:\Users\<id>' }
        @{ id = 'intranet_url'; pattern = '(?i)\bhttps?://(?:[a-z0-9-]+\.)*(?:local|lan|corp|internal|intra|intranet)\b[^\s"<>]*'; note = 'intranet URL' }
        @{ id = 'private_ip';  pattern = '\b(?:10\.\d{1,3}|192\.168|172\.(?:1[6-9]|2\d|3[01]))\.\d{1,3}\.\d{1,3}\b'; note = 'private IPv4 address' }
    )
}

function Read-EbiMaskDictionary {
    # profiles/mask-dictionary.json -> @{ words = string[]; allow = string[] }
    param([string]$Path)
    $d = @{ words = @(); allow = @() }
    if ([string]::IsNullOrWhiteSpace($Path) -or -not (Test-Path -LiteralPath $Path -PathType Leaf)) { return $d }
    $r = Read-EbiJson -Path $Path
    if (-not $r['ok'] -or -not ($r['value'] -is [System.Collections.IDictionary])) { return $d }
    if ($r['value'].Contains('words') -and $null -ne $r['value']['words']) { $d['words'] = @(@($r['value']['words']) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' }) }
    if ($r['value'].Contains('allow') -and $null -ne $r['value']['allow']) { $d['allow'] = @(@($r['value']['allow']) | ForEach-Object { [string]$_ } | Where-Object { $_ -ne '' }) }
    return $d
}

function Find-EbiMaskHitsInText {
    <#
      PURE. Text -> @(@{ rule; line; match; text }) for every rule / word
      hit that no allow pattern covers. Line numbers are 1-based.
    #>
    param([string]$Text, $Rules = $null, $Words = @(), $Allow = @())
    if ($null -eq $Rules) { $Rules = Get-EbiMaskRules }
    $hits = New-Object System.Collections.ArrayList
    if ([string]::IsNullOrEmpty($Text)) { return $hits.ToArray() }
    $lines = [regex]::Split($Text, "\r?\n")
    $allowRx = @(foreach ($a in @($Allow)) { if ([string]$a -ne '') { [string]$a } })
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ([string]::IsNullOrWhiteSpace($line)) { continue }
        foreach ($rule in @($Rules)) {
            foreach ($m in [regex]::Matches($line, [string]$rule['pattern'])) {
                $val = $m.Value
                $allowed = $false
                foreach ($a in $allowRx) { if ($val -match $a) { $allowed = $true; break } }
                if ($allowed) { continue }
                [void]$hits.Add(@{ rule = [string]$rule['id']; line = ($i + 1); match = $val; text = $line.Trim() })
            }
        }
        foreach ($w in @($Words)) {
            if ([string]$w -eq '') { continue }
            if ($line.IndexOf([string]$w, [System.StringComparison]::OrdinalIgnoreCase) -ge 0) { [void]$hits.Add(@{ rule = 'dictionary'; line = ($i + 1); match = [string]$w; text = $line.Trim() }) }
        }
    }
    return $hits.ToArray()
}

function Get-EbiMaskScanFiles {
    # The files the gate scans: every text file under profiles/**/fixtures,
    # plus profiles/**/*.json, workflows/*.json (a workflow may embed a
    # hostname), and docs/ebi-dance/**/*.md.
    param([string]$RepoRoot)
    $out = New-Object System.Collections.ArrayList
    $prof = Join-Path $RepoRoot 'profiles'
    if (Test-Path -LiteralPath $prof) {
        foreach ($f in @(Get-ChildItem -LiteralPath $prof -Recurse -File | Where-Object { $_.Extension -in @('.txt', '.json', '.md', '.csv') -and $_.Name -ne 'mask-dictionary.json' })) { [void]$out.Add($f.FullName) }
    }
    $wf = Join-Path $RepoRoot 'workflows'
    if (Test-Path -LiteralPath $wf) { foreach ($f in @(Get-ChildItem -LiteralPath $wf -Filter '*.json' -File)) { [void]$out.Add($f.FullName) } }
    return $out.ToArray()
}

function Invoke-EbiMaskCheck {
    <#
      Scan the files -> @{ ok; hits = @(@{ file; rule; line; match; text });
      files; message }. ok is $false on any hit.
    #>
    param([string]$RepoRoot, $Files = $null, [string]$DictionaryPath = '')
    if ($null -eq $Files) { $Files = @(Get-EbiMaskScanFiles -RepoRoot $RepoRoot) }
    if ($DictionaryPath -eq '') { $DictionaryPath = Join-Path (Join-Path $RepoRoot 'profiles') 'mask-dictionary.json' }
    $dict = Read-EbiMaskDictionary -Path $DictionaryPath
    $all = New-Object System.Collections.ArrayList
    foreach ($f in @($Files)) {
        $text = ''
        try { $text = [System.IO.File]::ReadAllText([string]$f, (New-Object System.Text.UTF8Encoding($false))) } catch { continue }
        foreach ($h in @(Find-EbiMaskHitsInText -Text $text -Words $dict['words'] -Allow $dict['allow'])) {
            $h['file'] = [string]$f
            [void]$all.Add($h)
        }
    }
    $msg = if ($all.Count -eq 0) { ('' + @($Files).Count + ' file(s) clean') } else { ('' + $all.Count + ' sensitive item(s) in ' + @(@($all.ToArray()) | ForEach-Object { $_['file'] } | Sort-Object -Unique).Count + ' file(s)') }
    return @{ ok = ($all.Count -eq 0); hits = $all.ToArray(); files = @($Files).Count; message = $msg }
}

function Format-EbiMaskReport {
    param($Result, [string]$RepoRoot = '')
    $L = New-Object System.Collections.ArrayList
    foreach ($h in @($Result['hits'])) {
        $f = [string]$h['file']
        if ($RepoRoot -ne '' -and $f.StartsWith($RepoRoot)) { $f = $f.Substring($RepoRoot.Length).TrimStart('\', '/') }
        [void]$L.Add(('  [MASK] {0}:{1}  {2}  "{3}"' -f $f, $h['line'], $h['rule'], $h['match']))
    }
    [void]$L.Add('  mask check: ' + $Result['message'] + $(if ($Result['ok']) { ' -- OK' } else { ' -- FAIL (mask them, or list a known-safe shape under "allow" in profiles/mask-dictionary.json)' }))
    return $L.ToArray()
}
