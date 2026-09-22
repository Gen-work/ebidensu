# ============================================================
#  kernel/Context.ps1
#
#  {{...}} template evaluation for workflow JSON (P1-01). Pure: no I/O,
#  no COM, no state. Dot-source only (no param()), ASCII source, no class.
#  The runner wires it in at P1-03; until then kernel/Runner.ps1 keeps
#  refusing templates up front.
#
#  Spec: docs/ebi-dance/spec/WORKFLOW-SCHEMA.md
#    4.1 scopes   vars / profile / page / run / item / steps.<id>.out.<f>
#    4.2 rules    lookup + string concatenation only; a value that is
#                 exactly one {{...}} keeps its type; \{\{ is a literal
#                 {{; nested templates are an error
#    4.3          a subtree pulled out of profile/page is evaluated ONCE
#                 (its own {{...}} resolve; the result is not re-parsed)
#    5.1          a when-skipped step has null outputs + skipped=true --
#                 the runner puts that shape into the scope, so here a
#                 null field is a value, not an error
#    7.4          a step id may contain "." (flow.call inlining), so the
#                 steps path is split at the literal ".out." boundary
#  and docs/ebi-dance/spec/PROFILE-SCHEMA.md 6.6 for item.key / keySafe
#  (kernel/Key.ps1).
#
#  Failures are return values (BACKLOG rule R6), never exceptions:
#    @{ ok = $false; error = <kind>; segment = <which piece>; path = <ref>;
#       message = <one English line> }
#  error kinds:
#    unknown_scope     prefix is not one of the six scopes
#    missing_segment   a path segment does not exist in its container
#    step_not_found    steps.<id> names no step in scope (not run yet, or typo)
#    bad_steps_path    steps.<...> without a ".out." boundary
#    no_page           {{page.X}} used but the workflow binds no page
#    nested_template   {{ inside {{ }}
#    not_scalar        a hashtable/array was interpolated into a string
#
#  Hashtables are read by index ($h['k']) throughout: dot access on a
#  missing key throws under Set-StrictMode, which the tests turn on.
# ============================================================

. (Join-Path $PSScriptRoot 'Key.ps1')

function New-EbiTemplateScope {
    # Everything a template can see, in one hashtable. Every part is
    # optional; a missing part just makes references into it fail with a
    # clear segment.
    #   Vars / Profile / Run / Item : hashtables
    #   PageName                    : the workflow's top-level "page"
    #   Steps                       : id -> outputs hashtable (the runner's
    #                                 record for that step; a skipped step
    #                                 has null fields + skipped=$true)
    #   KeyColumns                  : worklist.json key.columns
    #   GroupColumn                 : vocabulary.json columns.group
    param(
        [hashtable]$Vars = @{},
        [hashtable]$Profile = @{},
        [string]$PageName = '',
        [hashtable]$Run = @{},
        [hashtable]$Item = $null,
        [hashtable]$Steps = @{},
        $KeyColumns = @(),
        [string]$GroupColumn = ''
    )
    return @{
        vars        = $Vars
        profile     = $Profile
        page        = $PageName
        run         = $Run
        item        = $Item
        steps       = $Steps
        keyColumns  = @($KeyColumns)
        groupColumn = $GroupColumn
    }
}

function New-EbiPathFailure {
    param([string]$Kind, [string]$Segment, [string]$Path, [string]$Message)
    return @{ ok = $false; error = $Kind; segment = $Segment; path = $Path; message = $Message }
}

function Resolve-EbiPathSegments {
    # Walk $Root along $Segments. Hashtable keys by name, arrays by
    # non-negative integer index. Reports the first segment that does
    # not resolve.
    param($Root, $Segments, [string]$Path, [string]$Prefix)
    $cur = $Root
    $walked = $Prefix
    foreach ($seg in @($Segments)) {
        $s = [string]$seg
        if ($cur -is [System.Collections.IDictionary]) {
            if (-not $cur.Contains($s)) {
                return (New-EbiPathFailure 'missing_segment' $s $Path ('"' + $s + '" does not exist under "' + $walked + '"'))
            }
            $cur = $cur[$s]
        } elseif (($cur -is [System.Collections.IList]) -and -not ($cur -is [string])) {
            $idx = 0
            if (-not [int]::TryParse($s, [ref]$idx) -or $idx -lt 0 -or $idx -ge $cur.Count) {
                return (New-EbiPathFailure 'missing_segment' $s $Path ('"' + $s + '" is not a valid index into the array at "' + $walked + '"'))
            }
            $cur = $cur[$idx]
        } else {
            return (New-EbiPathFailure 'missing_segment' $s $Path ('"' + $walked + '" is a scalar (or null); it has no "' + $s + '"'))
        }
        $walked = $walked + '.' + $s
    }
    return @{ ok = $true; value = $cur }
}

function Resolve-EbiPath {
    <#
      One reference path (the text between {{ and }}) -> @{ ok; value }
      or a failure record. Also returns 'subtree' = $true when the value
      came out of profile/page, which is what 4.3's one-pass rule keys on.
    #>
    param([hashtable]$Scope, [string]$Path)

    $p = ([string]$Path).Trim()
    if ($p.Length -eq 0) { return (New-EbiPathFailure 'unknown_scope' '' $Path 'empty reference') }

    $dot = $p.IndexOf('.')
    $prefix = if ($dot -lt 0) { $p } else { $p.Substring(0, $dot) }
    $rest   = if ($dot -lt 0) { '' } else { $p.Substring($dot + 1) }
    $restSegs = @()
    if ($rest -ne '') { $restSegs = @($rest -split '\.') }

    switch ($prefix) {
        'vars' {
            $r = Resolve-EbiPathSegments -Root $Scope['vars'] -Segments $restSegs -Path $p -Prefix 'vars'
            return $r
        }
        'run' {
            $r = Resolve-EbiPathSegments -Root $Scope['run'] -Segments $restSegs -Path $p -Prefix 'run'
            return $r
        }
        'profile' {
            $r = Resolve-EbiPathSegments -Root $Scope['profile'] -Segments $restSegs -Path $p -Prefix 'profile'
            if ($r['ok']) { $r['subtree'] = $true }
            return $r
        }
        'page' {
            $pageName = [string]$Scope['page']
            if ([string]::IsNullOrEmpty($pageName)) {
                return (New-EbiPathFailure 'no_page' 'page' $p 'the workflow declares no top-level "page", so {{page.X}} is not available (WORKFLOW-SCHEMA 4.1)')
            }
            if ($restSegs.Count -eq 0) {
                return (New-EbiPathFailure 'missing_segment' 'page' $p '{{page}} needs a field; {{page.id}} is the page name')
            }
            $first = [string]$restSegs[0]
            $tail  = @()
            if ($restSegs.Count -gt 1) { $tail = @($restSegs[1..($restSegs.Count - 1)]) }
            $profile = $Scope['profile']
            if ($first -eq 'id') {
                if ($tail.Count -gt 0) { return (New-EbiPathFailure 'missing_segment' ([string]$tail[0]) $p '{{page.id}} is a string; it has no fields') }
                return @{ ok = $true; value = $pageName }
            }
            if ($first -eq 'grammar' -or $first -eq 'rules') {
                # 4.1: page.grammar / page.rules alias grammar.json / rules.json entries keyed by the page name
                if ($null -eq $profile -or -not $profile.Contains($first)) {
                    return (New-EbiPathFailure 'missing_segment' $first $p ('profile has no "' + $first + '" (grammar.json / rules.json not loaded)'))
                }
                $table = $profile[$first]
                if (-not ($table -is [System.Collections.IDictionary]) -or -not $table.Contains($pageName)) {
                    return (New-EbiPathFailure 'missing_segment' $pageName $p ('profile.' + $first + ' has no entry for page "' + $pageName + '"'))
                }
                $r = Resolve-EbiPathSegments -Root $table[$pageName] -Segments $tail -Path $p -Prefix ('profile.' + $first + '.' + $pageName)
                if ($r['ok']) { $r['subtree'] = $true }
                return $r
            }
            if ($null -eq $profile -or -not $profile.Contains('pages') -or -not ($profile['pages'] -is [System.Collections.IDictionary])) {
                return (New-EbiPathFailure 'missing_segment' 'pages' $p 'profile has no "pages" (pages.json not loaded)')
            }
            if (-not $profile['pages'].Contains($pageName)) {
                return (New-EbiPathFailure 'missing_segment' $pageName $p ('pages.json has no page named "' + $pageName + '"'))
            }
            $r = Resolve-EbiPathSegments -Root $profile['pages'][$pageName] -Segments $restSegs -Path $p -Prefix ('profile.pages.' + $pageName)
            if ($r['ok']) { $r['subtree'] = $true }
            return $r
        }
        'item' {
            $item = $Scope['item']
            if ($null -eq $item) {
                return (New-EbiPathFailure 'missing_segment' 'item' $p '{{item.X}} is only available inside "each" (no current item)')
            }
            if ($restSegs.Count -eq 0) {
                return (New-EbiPathFailure 'missing_segment' 'item' $p '{{item}} needs a column name, or key / keySafe / group')
            }
            $first = [string]$restSegs[0]
            if ($restSegs.Count -gt 1) {
                return (New-EbiPathFailure 'missing_segment' ([string]$restSegs[1]) $p ('item columns are scalars; "' + $first + '" has no "' + [string]$restSegs[1] + '"'))
            }
            # Derived keys win over same-named real columns (4.1).
            switch ($first) {
                'key'     { return @{ ok = $true; value = (Get-EbiKeyDisplay -Item $item -KeyColumns $Scope['keyColumns']) } }
                'keySafe' { return @{ ok = $true; value = (ConvertTo-EbiKeySafe -Item $item -KeyColumns $Scope['keyColumns']) } }
                'group'   {
                    $gcol = [string]$Scope['groupColumn']
                    if ($gcol -eq '') { return (New-EbiPathFailure 'missing_segment' 'group' $p 'no group column declared (vocabulary.json columns.group)') }
                    if (-not $item.Contains($gcol)) { return (New-EbiPathFailure 'missing_segment' $gcol $p ('the item has no column "' + $gcol + '" (declared as the group column)')) }
                    return @{ ok = $true; value = $item[$gcol] }
                }
            }
            if (-not $item.Contains($first)) {
                return (New-EbiPathFailure 'missing_segment' $first $p ('the item has no column "' + $first + '"'))
            }
            return @{ ok = $true; value = $item[$first] }
        }
        'steps' {
            # 7.4: the id may itself contain dots, so cut at the literal ".out." boundary
            $marker = '.out.'
            $idx = $rest.IndexOf($marker)
            if ($idx -le 0) {
                if ($rest.EndsWith('.out')) {
                    return (New-EbiPathFailure 'bad_steps_path' 'out' $p 'steps.<id>.out needs a field after ".out."')
                }
                return (New-EbiPathFailure 'bad_steps_path' $rest $p 'a steps reference is steps.<id>.out.<field>; no ".out." boundary found')
            }
            $stepId = $rest.Substring(0, $idx)
            $fieldPath = $rest.Substring($idx + $marker.Length)
            $steps = $Scope['steps']
            if ($null -eq $steps -or -not $steps.Contains($stepId)) {
                return (New-EbiPathFailure 'step_not_found' $stepId $p ('step "' + $stepId + '" has no outputs in scope (not run yet, in another section, or a typo)'))
            }
            $fieldSegs = @()
            if ($fieldPath -ne '') { $fieldSegs = @($fieldPath -split '\.') }
            if ($fieldSegs.Count -eq 0) {
                return (New-EbiPathFailure 'bad_steps_path' 'out' $p 'steps.<id>.out needs a field after ".out."')
            }
            return (Resolve-EbiPathSegments -Root $steps[$stepId] -Segments $fieldSegs -Path $p -Prefix ('steps.' + $stepId + '.out'))
        }
        default {
            return (New-EbiPathFailure 'unknown_scope' $prefix $p ('"' + $prefix + '" is not a template scope (vars / profile / page / run / item / steps)'))
        }
    }
}

# Escape handling: \{\{ is a literal "{{". It is parked as U+0001 while the
# real tokens are found, then restored -- so an escaped brace can never be
# mistaken for the start of a token.
function Get-EbiTemplateEscapeMark { return [string][char]1 }

function Get-EbiTemplateTokens {
    # Every {{...}} in a string, in order: @( @{ start; length; path } ).
    # A "{{" inside a token's body is the nested-template error.
    # Returns @{ ok; tokens; text } where text has escapes parked.
    param([string]$Value)
    $text = $Value.Replace('\{\{', (Get-EbiTemplateEscapeMark))
    $tokens = New-Object System.Collections.ArrayList
    $pos = 0
    while ($true) {
        $open = $text.IndexOf('{{', $pos)
        if ($open -lt 0) { break }
        $close = $text.IndexOf('}}', $open + 2)
        if ($close -lt 0) {
            return @{ ok = $false; error = 'nested_template'; segment = $text.Substring($open); path = $Value; message = 'unterminated {{ (no closing }})' }
        }
        $body = $text.Substring($open + 2, $close - $open - 2)
        if ($body.IndexOf('{{') -ge 0 -or $body.IndexOf('{') -ge 0 -or $body.IndexOf('}') -ge 0) {
            return @{ ok = $false; error = 'nested_template'; segment = $body.Trim(); path = $Value; message = 'nested templates are not allowed (WORKFLOW-SCHEMA 4.2); bind the page instead' }
        }
        [void]$tokens.Add(@{ start = $open; length = ($close + 2 - $open); path = $body.Trim() })
        $pos = $close + 2
    }
    return @{ ok = $true; tokens = $tokens.ToArray(); text = $text }
}

function Test-EbiTemplateString {
    # Does this string carry any template (escapes excluded)?
    param($Value)
    if (-not ($Value -is [string])) { return $false }
    $t = Get-EbiTemplateTokens -Value $Value
    if (-not $t['ok']) { return $true }
    return (@($t['tokens']).Count -gt 0)
}

function Get-EbiTemplateReferences {
    # Every reference path used anywhere inside a value (strings, nested
    # hashtables, arrays). For ebi lint (P1-08): static "does this
    # resolve" checks without a runtime scope. Unparseable strings
    # contribute nothing here; Expand-EbiTemplate reports them.
    param($Value)
    $refs = New-Object System.Collections.ArrayList
    if ($null -eq $Value) { return $refs.ToArray() }
    if ($Value -is [string]) {
        $t = Get-EbiTemplateTokens -Value $Value
        if ($t['ok']) { foreach ($tok in @($t['tokens'])) { [void]$refs.Add([string]$tok['path']) } }
        return $refs.ToArray()
    }
    if ($Value -is [System.Collections.IDictionary]) {
        foreach ($k in $Value.Keys) { foreach ($r in @(Get-EbiTemplateReferences -Value $Value[$k])) { [void]$refs.Add($r) } }
        return $refs.ToArray()
    }
    if ($Value -is [System.Collections.IList]) {
        foreach ($i in $Value) { foreach ($r in @(Get-EbiTemplateReferences -Value $i)) { [void]$refs.Add($r) } }
        return $refs.ToArray()
    }
    return $refs.ToArray()
}

function ConvertTo-EbiTemplateText {
    # A resolved value inside a larger string: scalars only.
    param($Value, [string]$Path)
    if ($null -eq $Value) { return @{ ok = $true; text = '' } }
    if ($Value -is [string]) { return @{ ok = $true; text = $Value } }
    if ($Value -is [bool]) { return @{ ok = $true; text = $(if ($Value) { 'true' } else { 'false' }) } }
    if ($Value -is [System.Collections.IDictionary] -or ($Value -is [System.Collections.IList])) {
        return (New-EbiPathFailure 'not_scalar' $Path $Path ('{{' + $Path + '}} is an object/array; it can only be used as a whole value, not inside a string'))
    }
    return @{ ok = $true; text = [string]$Value }
}

function Expand-EbiTemplate {
    <#
      Evaluate every {{...}} inside $Value against $Scope.
        string     -> the value of the single token (type kept), or the
                      string with all tokens substituted
        hashtable  -> new hashtable, values expanded
        array      -> new array, elements expanded
        other      -> unchanged
      Returns @{ ok = $true; value = ... } or a failure record whose
      'path' is the offending reference and 'segment' the piece that
      failed.

      -Depth is internal: a subtree pulled from profile/page is expanded
      once with Depth 1, and at Depth 1 a resolved subtree is NOT expanded
      again (4.3: one pass, no second level).
    #>
    param($Value, [hashtable]$Scope, [int]$Depth = 0)

    if ($null -eq $Value) { return @{ ok = $true; value = $null } }

    if ($Value -is [string]) {
        $t = Get-EbiTemplateTokens -Value $Value
        if (-not $t['ok']) { return $t }
        $tokens = @($t['tokens'])
        $text = [string]$t['text']
        if ($tokens.Count -eq 0) {
            return @{ ok = $true; value = $text.Replace((Get-EbiTemplateEscapeMark), '{{') }
        }

        # Whole value is exactly one token -> keep the resolved type.
        if ($tokens.Count -eq 1 -and $text.Trim() -eq $text.Substring($tokens[0]['start'], $tokens[0]['length'])) {
            $r = Resolve-EbiPath -Scope $Scope -Path $tokens[0]['path']
            if (-not $r['ok']) { return $r }
            $v = $r['value']
            # 4.3: anything pulled out of profile/page -- a subtree OR a single
            # string such as worklist.file = "mapping_{{run.operator}}.csv"
            # (P0-R11) -- gets exactly one pass of its own; at Depth 1 nothing
            # is expanded again, so a template inside that result stays literal.
            $fromProfile = ($r.Contains('subtree') -and $r['subtree'])
            if ($fromProfile -and $Depth -eq 0) {
                return (Expand-EbiTemplate -Value $v -Scope $Scope -Depth 1)
            }
            return @{ ok = $true; value = $v }
        }

        # Interpolation: scalars only.
        $sb = New-Object System.Text.StringBuilder
        $cursor = 0
        foreach ($tok in $tokens) {
            [void]$sb.Append($text.Substring($cursor, $tok['start'] - $cursor))
            $r = Resolve-EbiPath -Scope $Scope -Path $tok['path']
            if (-not $r['ok']) { return $r }
            $asText = ConvertTo-EbiTemplateText -Value $r['value'] -Path $tok['path']
            if (-not $asText['ok']) { return $asText }
            [void]$sb.Append([string]$asText['text'])
            $cursor = $tok['start'] + $tok['length']
        }
        [void]$sb.Append($text.Substring($cursor))
        return @{ ok = $true; value = $sb.ToString().Replace((Get-EbiTemplateEscapeMark), '{{') }
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $out = @{}
        foreach ($k in $Value.Keys) {
            $r = Expand-EbiTemplate -Value $Value[$k] -Scope $Scope -Depth $Depth
            if (-not $r['ok']) { return $r }
            $out[$k] = $r['value']
        }
        return @{ ok = $true; value = $out }
    }

    if ($Value -is [System.Collections.IList]) {
        $list = New-Object System.Collections.ArrayList
        foreach ($i in $Value) {
            $r = Expand-EbiTemplate -Value $i -Scope $Scope -Depth $Depth
            if (-not $r['ok']) { return $r }
            [void]$list.Add($r['value'])
        }
        # No comma guard here: inside a hashtable literal nothing is unrolled,
        # and ,$x would store an array wrapped in another array.
        return @{ ok = $true; value = $list.ToArray() }
    }

    return @{ ok = $true; value = $Value }
}
