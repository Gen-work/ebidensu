# DocsCheck.ps1 -- pure helpers for the ebi-dance documentation checks.
# Dot-sourced by Tests\Test-Docs.ps1. No param() block, ASCII source only:
# every Chinese/Japanese pattern this needs is passed in from
# Tests\docs-checks.json, and the section sign is built from [char] -- never
# written as a literal (CLAUDE.md encoding rule).
#
# Everything here is a pure string function -- no file I/O, no COM -- so the
# checks are unit-tested against inline fixtures before they are ever pointed
# at the real docs tree.

function Get-DocsSectionMark {
    # U+00A7 SECTION SIGN. Built here so callers stay ASCII too.
    return [string][char]0x00A7
}

function Get-DocsCardStats {
    # Count BACKLOG cards per phase from its text. P0-00 is a historical
    # status note, not an executable card, so it is excluded -- that is the
    # counting convention BACKLOG's own status section documents.
    param(
        [string]$Text,
        [string]$BlockMarker
    )
    $total  = @{}
    $blocks = @{}
    $rx = [regex]'(?m)^### \[[ x]\] (P\d)-(\S+)(.*)$'
    foreach ($m in $rx.Matches($Text)) {
        $phase = $m.Groups[1].Value
        $num   = $m.Groups[2].Value
        $rest  = $m.Groups[3].Value
        if ($phase -eq 'P0' -and $num -eq '00') { continue }
        if (-not $total.ContainsKey($phase))  { $total[$phase]  = 0 }
        if (-not $blocks.ContainsKey($phase)) { $blocks[$phase] = 0 }
        $total[$phase] = $total[$phase] + 1
        if ($BlockMarker -and $rest.Contains($BlockMarker)) {
            $blocks[$phase] = $blocks[$phase] + 1
        }
    }
    $early = @('P0', 'P1', 'P2', 'P3')
    $late  = @('P4', 'P5')
    $p0p3 = 0; foreach ($k in $early) { if ($total.ContainsKey($k))  { $p0p3       += $total[$k]  } }
    $p4p5 = 0; foreach ($k in $late)  { if ($total.ContainsKey($k))  { $p4p5       += $total[$k]  } }
    $eb   = 0; foreach ($k in $early) { if ($blocks.ContainsKey($k)) { $eb         += $blocks[$k] } }
    $lb   = 0; foreach ($k in $late)  { if ($blocks.ContainsKey($k)) { $lb         += $blocks[$k] } }
    return @{
        Total      = $total
        Blocks     = $blocks
        P0P3       = $p0p3
        P4P5       = $p4p5
        All        = $p0p3 + $p4p5
        P0P3Blocks = $eb
        AllBlocks  = $eb + $lb
        Rest       = ($p0p3 + $p4p5) - ($eb + $lb)
    }
}

function Get-DocsExpectedCount {
    # Resolve one 'kind' token from a countChecks entry against the stats.
    # 'total:P0' / 'blocks:P2' address one phase; the rest are derived totals.
    param($Stats, [string]$Kind)
    if ($Kind -like 'total:*') {
        $p = $Kind.Substring(6)
        if ($Stats.Total.ContainsKey($p)) { return $Stats.Total[$p] } else { return 0 }
    }
    if ($Kind -like 'blocks:*') {
        $p = $Kind.Substring(7)
        if ($Stats.Blocks.ContainsKey($p)) { return $Stats.Blocks[$p] } else { return 0 }
    }
    switch ($Kind) {
        'p0p3'       { return $Stats.P0P3 }
        'p0p3plus1'  { return $Stats.P0P3 + 1 }
        'p4p5'       { return $Stats.P4P5 }
        'all'        { return $Stats.All }
        'p0p3Blocks' { return $Stats.P0P3Blocks }
        'allBlocks'  { return $Stats.AllBlocks }
        'rest'       { return $Stats.Rest }
    }
    return $null
}

function Find-DocsTextHit {
    # Line numbers where a literal needle appears. Used for the dead-name
    # blacklist: entries there must never legitimately appear anywhere, so a
    # single hit is a failure. Spellings that DO appear in "this was the old
    # bug" illustrations are deliberately kept out of that list.
    param([string]$Text, [string]$Needle)
    $hits  = New-Object System.Collections.Generic.List[int]
    $lines = $Text -split "`n"
    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i].Contains($Needle)) { [void]$hits.Add($i + 1) }
    }
    return $hits.ToArray()
}

function Get-DocsSectionId {
    # Section numbers a markdown file defines: '## 3. x' / '### 6.1 y' -> 3, 6.1
    param([string]$Text)
    $ids = New-Object System.Collections.Generic.List[string]
    $rx  = [regex]'(?m)^#{2,6}\s+(\d+(?:\.\d+)*)\.?\s'
    foreach ($m in $rx.Matches($Text)) { [void]$ids.Add($m.Groups[1].Value) }
    return $ids.ToArray()
}

function Get-DocsSectionRef {
    # Section references on ONE line, each resolved to the nearest PRECEDING
    # `<something>.md` token on that same line. A reference with no filename
    # before it on its own line is skipped, not guessed: card fields wrap
    # across lines ("...`A.md` s1,\n  s2;`B.md` s3"), and guessing there
    # produced only false positives. Precision over coverage -- a checker
    # nobody trusts gets switched off.
    param([string]$Line)
    $mark  = Get-DocsSectionMark
    $out   = New-Object System.Collections.Generic.List[object]
    $files = [regex]::Matches($Line, '`([A-Za-z0-9_/-]+\.md)`')
    foreach ($m in [regex]::Matches($Line, ($mark + '(\d+(?:\.\d+)*)'))) {
        $target = $null
        foreach ($f in $files) {
            if (($f.Index + $f.Length) -le $m.Index) {
                $target = Split-Path -Leaf $f.Groups[1].Value
            }
        }
        if ($target) {
            [void]$out.Add([pscustomobject]@{ Ref = $m.Groups[1].Value; File = $target })
        }
    }
    return $out.ToArray()
}

function Get-DocsJsonBlock {
    # Fenced ```json / ```jsonc blocks, with the file line number of each line.
    param([string]$Text)
    $out   = New-Object System.Collections.Generic.List[object]
    $lines = $Text -split "`n"
    $open  = $false
    $body  = $null
    for ($i = 0; $i -lt $lines.Count; $i++) {
        $line = $lines[$i]
        if ($line.StartsWith('```')) {
            if ($open) {
                [void]$out.Add([pscustomobject]@{ Lines = $body.ToArray() })
                $open = $false
                $body = $null
            } elseif ($line -match '^```jsonc?\s*$') {
                $open = $true
                $body = New-Object System.Collections.Generic.List[object]
            }
            continue
        }
        if ($open) {
            [void]$body.Add([pscustomobject]@{ Line = $i + 1; Text = $line })
        }
    }
    return $out.ToArray()
}

function Test-DocsExampleBlock {
    # Two rules for workflow examples:
    #  1. every step call carries an "id" (P0-R3: it is the ledger key)
    #  2. every {{steps.X}} resolves inside the same block -- but only for
    #     blocks that are a whole workflow ("each"/"setup" present). Short
    #     fragments legitimately reference a step defined elsewhere.
    param($Block)
    $issues   = New-Object System.Collections.Generic.List[object]
    $texts    = @()
    foreach ($l in $Block.Lines) { $texts += $l.Text }
    $bodyText = $texts -join "`n"
    $ids = @{}
    foreach ($m in [regex]::Matches($bodyText, '"id"\s*:\s*"([^"]+)"')) {
        $ids[$m.Groups[1].Value] = $true
    }
    $prev = ''
    foreach ($l in $Block.Lines) {
        if ($l.Text.Contains('"use"') -and -not $l.Text.Contains('"id"') -and -not $prev.Contains('"id"')) {
            [void]$issues.Add([pscustomobject]@{ Line = $l.Line; Kind = 'no-id'; Name = '' })
        }
        $prev = $l.Text
    }
    if ($bodyText.Contains('"each"') -or $bodyText.Contains('"setup"')) {
        foreach ($l in $Block.Lines) {
            foreach ($m in [regex]::Matches($l.Text, '\{\{steps\.([A-Za-z0-9_]+)\.')) {
                $name = $m.Groups[1].Value
                if (-not $ids.ContainsKey($name)) {
                    [void]$issues.Add([pscustomobject]@{ Line = $l.Line; Kind = 'undef-step'; Name = $name })
                }
            }
        }
    }
    return $issues.ToArray()
}
