#Requires -Version 5.1
# ============================================================
#  kernel/Worklist.ps1
#
#  The in-memory worklist and the ONE row-selection implementation
#  (P1-03). Dot-source only (no param() block, ASCII source, no class).
#
#  The runner's source.select and the table.select step both call
#  Select-EbiWorklistRows -- the card says "one function, not two copies",
#  and this file is where that function lives. Pure: no file access.
#
#  The worklist resource (STEP-CONTRACT.md 3.2, 3.4 point 6; kind
#  'worklist') is a hashtable:
#      @{ path = <csv path or ''>; columns = string[]; rows = object[] }
#  where every row is a hashtable column -> string. table.load (P1-24)
#  builds it; the runner iterates rows; table.set / flow.checkpoint write
#  rows and save. Nothing here assumes any column name.
#
#  pendingWhen (WORKFLOW-SCHEMA.md 3.1) is FIVE fixed forms, not an
#  expression language:
#      "empty"        blank or 0 (or the profile's pending code)
#      "!= ok"        logical value is not ok
#      "== ng"        logical value is ng
#      "bit !<name>"  bitmask column: that bit (by name from the column's
#                     bits map, or by number) is not set
#      "always"       every row
#  Comparison happens on LOGICAL values: a verdict column may declare
#  values = @{ ok='1'; ng='2'; unknown=''; pending='0' } (PROFILE-SCHEMA.md
#  6.5, P0-R11) and the stored cell is translated back before comparing,
#  so the workflow JSON only ever says ok / ng.
# ============================================================

function Get-EbiPendingWhenForms {
    return @('empty', '!= ok', '== ng', 'bit !<name>', 'always')
}

function ConvertFrom-EbiPendingWhen {
    <#
      PURE. Parse a pendingWhen text into @{ ok; kind; arg; message }.
      kind: empty | notOk | isNg | bit | always. Whitespace around the
      operator is tolerated ('!=ok' and '!= ok' are the same), nothing
      else is: '!= done' is an error, not a new form.
    #>
    param([string]$Text)
    $t = if ($null -eq $Text) { '' } else { $Text.Trim() }
    if ($t -eq '') { return @{ ok = $false; kind = ''; arg = ''; message = 'pendingWhen is empty; allowed: ' + ((Get-EbiPendingWhenForms) -join ', ') } }
    if ($t -eq 'empty')  { return @{ ok = $true; kind = 'empty';  arg = ''; message = '' } }
    if ($t -eq 'always') { return @{ ok = $true; kind = 'always'; arg = ''; message = '' } }
    if ($t -match '^!=\s*ok$') { return @{ ok = $true; kind = 'notOk'; arg = ''; message = '' } }
    if ($t -match '^==\s*ng$') { return @{ ok = $true; kind = 'isNg';  arg = ''; message = '' } }
    if ($t -match '^bit\s+!\s*([A-Za-z0-9_]+)$') { return @{ ok = $true; kind = 'bit'; arg = $Matches[1]; message = '' } }
    return @{ ok = $false; kind = ''; arg = ''; message = ('pendingWhen "' + $t + '" is not one of the five allowed forms: ' + ((Get-EbiPendingWhenForms) -join ', ')) }
}

function Get-EbiWorklistColumnSpec {
    # The profile's worklist.columns entry for a column name, or $null.
    param($Profile, [string]$Field)
    if ($null -eq $Profile -or -not ($Profile -is [System.Collections.IDictionary])) { return $null }
    if (-not $Profile.Contains('worklist') -or -not ($Profile['worklist'] -is [System.Collections.IDictionary])) { return $null }
    $wl = $Profile['worklist']
    if (-not $wl.Contains('columns') -or $null -eq $wl['columns']) { return $null }
    foreach ($c in $wl['columns']) {
        if ($c -is [System.Collections.IDictionary] -and $c.Contains('name') -and [string]$c['name'] -eq $Field) { return $c }
    }
    return $null
}

function Get-EbiWorklistKeyColumns {
    # profile.worklist.key.columns as string[]; empty when not declared.
    param($Profile)
    $out = New-Object System.Collections.ArrayList
    if ($null -eq $Profile -or -not ($Profile -is [System.Collections.IDictionary])) { return $out.ToArray() }
    if (-not $Profile.Contains('worklist') -or -not ($Profile['worklist'] -is [System.Collections.IDictionary])) { return $out.ToArray() }
    $wl = $Profile['worklist']
    if (-not $wl.Contains('key') -or -not ($wl['key'] -is [System.Collections.IDictionary])) { return $out.ToArray() }
    if (-not $wl['key'].Contains('columns') -or $null -eq $wl['key']['columns']) { return $out.ToArray() }
    $cols = $wl['key']['columns']
    if ($cols -is [string]) { [void]$out.Add($cols); return $out.ToArray() }
    foreach ($c in $cols) { if ($null -ne $c) { [void]$out.Add([string]$c) } }
    return $out.ToArray()
}

function Get-EbiWorklistGroupColumn {
    # profile.vocabulary.columns.group, or ''.
    param($Profile)
    if ($null -eq $Profile -or -not ($Profile -is [System.Collections.IDictionary])) { return '' }
    if (-not $Profile.Contains('vocabulary') -or -not ($Profile['vocabulary'] -is [System.Collections.IDictionary])) { return '' }
    $v = $Profile['vocabulary']
    if (-not $v.Contains('columns') -or -not ($v['columns'] -is [System.Collections.IDictionary])) { return '' }
    if (-not $v['columns'].Contains('group') -or $null -eq $v['columns']['group']) { return '' }
    return [string]$v['columns']['group']
}

function Get-EbiWorklistValuesMap {
    # logical -> stored map of a verdict column, or @{} (identity).
    param($Spec)
    $map = @{}
    if ($null -eq $Spec -or -not ($Spec -is [System.Collections.IDictionary])) { return $map }
    if (-not $Spec.Contains('values') -or -not ($Spec['values'] -is [System.Collections.IDictionary])) { return $map }
    foreach ($k in $Spec['values'].Keys) { $map[[string]$k] = [string]$Spec['values'][$k] }
    return $map
}

function ConvertTo-EbiStoredValue {
    # logical ('ok' / 'ng' / 'unknown' / 'pending' / '') -> what the cell
    # stores. Without a values map the logical value IS the stored value.
    param([string]$Logical, $Spec)
    $map = Get-EbiWorklistValuesMap -Spec $Spec
    $l = if ($null -eq $Logical) { '' } else { $Logical }
    if ($map.Contains($l)) { return $map[$l] }
    return $l
}

function ConvertFrom-EbiStoredValue {
    <#
      stored cell -> logical value. A blank cell is always '' (never
      translated to a name, even when the map says unknown = ''): blank
      means "nothing written yet" everywhere in this project. Otherwise
      the first logical whose stored code equals the cell wins, checked
      in the fixed order ok, ng, unknown, pending; a cell no code matches
      is returned as-is.
    #>
    param($Value, $Spec)
    $raw = if ($null -eq $Value) { '' } else { ([string]$Value).Trim() }
    if ($raw -eq '') { return '' }
    $map = Get-EbiWorklistValuesMap -Spec $Spec
    foreach ($logical in @('ok', 'ng', 'unknown', 'pending')) {
        if ($map.Contains($logical) -and [string]$map[$logical] -ceq $raw) { return $logical }
    }
    return $raw
}

function Get-EbiWorklistBitValue {
    # 'before' -> 1 via the column's bits map; '3' -> 3. 0 when unknown.
    param([string]$Name, $Spec)
    if ($null -eq $Name) { return 0 }
    if ($Name -match '^\d+$') { return [int]$Name }
    if ($null -ne $Spec -and ($Spec -is [System.Collections.IDictionary]) -and $Spec.Contains('bits') -and ($Spec['bits'] -is [System.Collections.IDictionary])) {
        $bits = $Spec['bits']
        if ($bits.Contains($Name)) { return [int]$bits[$Name] }
    }
    return 0
}

function Test-EbiRowPending {
    <#
      PURE. Does this row's Field count as pending under a parsed
      pendingWhen (ConvertFrom-EbiPendingWhen)? Spec is the column's
      profile entry (values / bits), may be $null.
        empty   ''  or 0 or the logical 'pending'
        notOk   logical -ne 'ok'   (so ng and unknown are pending, 3.2)
        isNg    logical -eq 'ng'
        bit     (int cell) -band bit -ne bit; a non-numeric cell is 0
        always  $true
    #>
    param($Row, [string]$Field, [hashtable]$Parsed, $Spec)
    $raw = $null
    if ($null -ne $Row -and ($Row -is [System.Collections.IDictionary]) -and $Row.Contains($Field)) { $raw = $Row[$Field] }
    switch ([string]$Parsed['kind']) {
        'always' { return $true }
        'bit' {
            $n = 0
            $s = if ($null -eq $raw) { '' } else { ([string]$raw).Trim() }
            if ($s -match '^\d+$') { $n = [int]$s }
            $bit = Get-EbiWorklistBitValue -Name ([string]$Parsed['arg']) -Spec $Spec
            if ($bit -le 0) { return $true }   # an unknown bit name is never "done"
            return (($n -band $bit) -ne $bit)
        }
    }
    $logical = ConvertFrom-EbiStoredValue -Value $raw -Spec $Spec
    switch ([string]$Parsed['kind']) {
        'empty' { return ($logical -eq '' -or $logical -eq '0' -or $logical -eq 'pending') }
        'notOk' { return ($logical -cne 'ok') }
        'isNg'  { return ($logical -ceq 'ng') }
    }
    return $false
}

function Select-EbiWorklistRows {
    <#
      PURE. The one row filter (runner source.select AND table.select).
      Returns @{ ok; rows; message; total; selected }. Rows keep their
      original order and identity (the same hashtable objects, so a write
      through a selected row reaches the worklist).
        Field / PendingWhen   what to test; PendingWhen must parse
        Spec                  the column's profile entry (values / bits)
        Only                  optional key display strings (--only): a row
                              is kept only when its key is listed
        KeyColumns            needed with Only
        Limit                 0 = all
    #>
    param($Rows, [string]$Field, [string]$PendingWhen, $Spec, $Only = $null, $KeyColumns = @(), [int]$Limit = 0)
    $out = New-Object System.Collections.ArrayList
    $parsed = ConvertFrom-EbiPendingWhen -Text $PendingWhen
    if (-not $parsed['ok']) { return @{ ok = $false; rows = $out.ToArray(); message = $parsed['message']; total = 0; selected = 0 } }
    if ($parsed['kind'] -ne 'always' -and [string]::IsNullOrWhiteSpace($Field)) {
        return @{ ok = $false; rows = $out.ToArray(); message = 'select.field is required unless pendingWhen is "always"'; total = 0; selected = 0 }
    }
    $onlySet = $null
    if ($null -ne $Only) {
        $onlySet = New-Object System.Collections.ArrayList
        foreach ($k in @($Only)) { if ($null -ne $k -and [string]$k -ne '') { [void]$onlySet.Add(([string]$k).Trim()) } }
    }
    $total = 0
    foreach ($row in @($Rows)) { if ($null -ne $row) { $total++ } }
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        if (-not (Test-EbiRowPending -Row $row -Field $Field -Parsed $parsed -Spec $Spec)) { continue }
        if ($null -ne $onlySet -and $onlySet.Count -gt 0) {
            $key = Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns
            if (-not ($onlySet -contains $key)) { continue }
        }
        [void]$out.Add($row)
        if ($Limit -gt 0 -and $out.Count -ge $Limit) { break }
    }
    return @{ ok = $true; rows = $out.ToArray(); message = ''; total = $total; selected = $out.Count }
}

function Sort-EbiWorklistRows {
    <#
      PURE, STABLE. Order rows for iteration: by GroupBy column first
      (so a group's items are contiguous and once:group / once:groupEnd
      fire once per group), then by OrderBy ('key' = the key display,
      else a column name). Ordinal string comparison, blanks last.
      Returns object[] of the same row objects.
    #>
    param($Rows, [string]$GroupBy = '', [string]$OrderBy = '', $KeyColumns = @())
    $list = New-Object System.Collections.ArrayList
    $i = 0
    foreach ($row in @($Rows)) {
        if ($null -eq $row) { continue }
        $g = ''
        if ($GroupBy -ne '' -and $row.Contains($GroupBy) -and $null -ne $row[$GroupBy]) { $g = [string]$row[$GroupBy] }
        $o = ''
        if ($OrderBy -eq 'key') { $o = Get-EbiKeyDisplay -Item $row -KeyColumns $KeyColumns }
        elseif ($OrderBy -ne '' -and $row.Contains($OrderBy) -and $null -ne $row[$OrderBy]) { $o = [string]$row[$OrderBy] }
        [void]$list.Add(@{ row = $row; g = $g; o = $o; i = $i })
        $i++
    }
    if ($GroupBy -eq '' -and $OrderBy -eq '') {
        $plain = New-Object System.Collections.ArrayList
        foreach ($e in $list) { [void]$plain.Add($e['row']) }
        return $plain.ToArray()
    }
    $sorted = @($list.ToArray() | Sort-Object -Property @{ Expression = { $_['g'] -eq '' } }, @{ Expression = { $_['g'] } }, @{ Expression = { $_['o'] -eq '' } }, @{ Expression = { $_['o'] } }, @{ Expression = { $_['i'] } })
    $out = New-Object System.Collections.ArrayList
    foreach ($e in $sorted) { [void]$out.Add($e['row']) }
    return $out.ToArray()
}

function Get-EbiWorklistGroupOf {
    # The group value of a row under GroupBy ('' when no GroupBy).
    param($Row, [string]$GroupBy)
    if ([string]::IsNullOrEmpty($GroupBy) -or $null -eq $Row) { return '' }
    if ($Row.Contains($GroupBy) -and $null -ne $Row[$GroupBy]) { return [string]$Row[$GroupBy] }
    return ''
}
