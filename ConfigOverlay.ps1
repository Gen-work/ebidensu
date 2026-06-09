# ============================================================
# ConfigOverlay.ps1  --  shared dot-source lib (no param(); ASCII; no BOM)
#
# Per-work-folder JSON config overlay support.
#
# VerifyConfig.psd1 holds the project DEFAULTS (and the structural schema:
# Scripts / PhaseOrder / Aliases). Each work folder may carry a JSON file
# (default name verify_config.json) that OVERRIDES those defaults for that
# folder only. The JSON is deep-merged over the .psd1 hashtable: JSON wins,
# CLI args still win over JSON.
#
# This file is pure (no Excel, no Edge, no file IO except the generator's
# string work) so it is unit-tested by Tests\Test-ConfigOverlay.ps1.
#
# Public functions:
#   ConvertTo-ConfigHashtable   PSCustomObject/array tree -> hashtable/array tree
#   ConvertFrom-ConfigJson      JSON text -> config hashtable
#   Merge-ConfigHashtable       deep-merge overlay onto base (base wins structure)
#   New-ConfigOverlaySnapshot   curated, operator-facing subset of a config
#   Remove-ConfigEmptyArray     drop empty arrays (PS 5.1 serializes them as "")
#   ConvertFrom-JsonUnicodeEscape  turn \uXXXX (>=0x80) back into real chars
#   Get-ConfigOverlayJson       config hashtable -> pretty, readable JSON text
# ============================================================

function ConvertTo-ConfigHashtable {
    # Recursively normalise a ConvertFrom-Json result (PSCustomObject / arrays /
    # scalars) into hashtables + arrays. Mark.Boxes consumers index by key and
    # call .ContainsKey(), so JSON objects MUST become hashtables, not
    # PSCustomObjects.
    param([object]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Management.Automation.PSCustomObject]) {
        $h = @{}
        foreach ($prop in $Value.PSObject.Properties) {
            $h[$prop.Name] = ConvertTo-ConfigHashtable $prop.Value
        }
        return $h
    }

    if ($Value -is [System.Collections.IDictionary]) {
        $h = @{}
        foreach ($k in @($Value.Keys)) {
            $h[$k] = ConvertTo-ConfigHashtable $Value[$k]
        }
        return $h
    }

    if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
        $list = @()
        foreach ($item in $Value) { $list += ,(ConvertTo-ConfigHashtable $item) }
        return ,$list
    }

    return $Value
}

function ConvertFrom-ConfigJson {
    # JSON text -> config hashtable (empty hashtable when blank / not an object).
    param([string]$Json)

    if ([string]::IsNullOrWhiteSpace($Json)) { return @{} }
    $obj = $Json | ConvertFrom-Json
    $h = ConvertTo-ConfigHashtable $obj
    if ($h -is [hashtable]) { return $h }
    return @{}
}

function Merge-ConfigHashtable {
    # Deep-merge $Overlay onto $Base (mutating + returning $Base).
    # Nested hashtables merge key-by-key; every other value (scalars, arrays)
    # is replaced wholesale by the overlay value.
    param([hashtable]$Base, [hashtable]$Overlay)

    if ($null -eq $Base)    { $Base = @{} }
    if ($null -eq $Overlay) { return $Base }

    foreach ($k in @($Overlay.Keys)) {
        $ov = $Overlay[$k]
        if ($Base.ContainsKey($k) -and ($Base[$k] -is [hashtable]) -and ($ov -is [hashtable])) {
            $Base[$k] = Merge-ConfigHashtable $Base[$k] $ov
        } else {
            $Base[$k] = $ov
        }
    }
    return $Base
}

function Remove-ConfigEmptyArray {
    # Windows PowerShell 5.1 serialises an empty array @() as "" via
    # ConvertTo-Json, which would re-read as a string and break list consumers
    # (e.g. Mark.Boxes folders). Drop empty-array entries from a hashtable tree
    # so the generated JSON never carries that trap; the .psd1 default is kept
    # by the deep-merge when a key is simply absent.
    param([object]$Value)

    if ($null -eq $Value) { return $null }

    if ($Value -is [System.Collections.IDictionary]) {
        $out = @{}
        foreach ($k in @($Value.Keys)) {
            $cleaned = Remove-ConfigEmptyArray $Value[$k]
            if (($null -ne $cleaned) -and
                ($cleaned -is [System.Collections.IEnumerable]) -and
                ($cleaned -isnot [string]) -and
                ($cleaned -isnot [System.Collections.IDictionary])) {
                if (@($cleaned).Count -eq 0) { continue }
            }
            $out[$k] = $cleaned
        }
        return $out
    }

    if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
        $list = @()
        foreach ($item in $Value) { $list += ,(Remove-ConfigEmptyArray $item) }
        return ,$list
    }

    return $Value
}

function New-ConfigOverlaySnapshot {
    # Build the curated, operator-facing subset of an effective config that
    # InitConfig writes to <WorkDir>\verify_config.json. Structural keys
    # (Scripts / PhaseOrder / Aliases) are intentionally excluded; any key may
    # still be added to the JSON by hand.
    param([hashtable]$Config)

    $snap = @{}
    $snap['_README'] = @(
        'This file overrides VerifyConfig.psd1 for THIS work folder only.',
        'JSON values win over the .psd1 defaults; CLI args still win over JSON.',
        'Any VerifyConfig.psd1 key may be added here - this is just a starter set.',
        'Regenerate with:  VerifyTool.ps1 -Phase InitConfig -Force   (keeps a .bak)',
        'Save as UTF-8. Mail / CheckSheet Japanese text is fine as plain UTF-8.'
    )

    $copyKeys = @(
        'DefaultOwner', 'Window', 'Timing', 'Review', 'Replace', 'Mark',
        'GfixLog', 'Df', 'Align', 'Clone', 'Reviewer', 'Mail', 'CheckSheet',
        'DeliverFiles', 'ExpectedTime', 'Paths'
    )
    foreach ($k in $copyKeys) {
        if ($Config.ContainsKey($k)) { $snap[$k] = $Config[$k] }
    }

    return (Remove-ConfigEmptyArray $snap)
}

function ConvertFrom-JsonUnicodeEscape {
    # Windows PowerShell 5.1 ConvertTo-Json escapes every non-ASCII char as
    # \uXXXX, which makes Japanese mail templates unreadable in the file. Turn
    # escapes for code points >= 0x80 back into real characters; leave ASCII /
    # control escapes (< 0x80) untouched so the JSON stays valid.
    param([string]$Json)

    if ([string]::IsNullOrEmpty($Json)) { return $Json }

    $evaluator = {
        param($m)
        $code = [Convert]::ToInt32($m.Groups[1].Value, 16)
        if ($code -lt 0x80) { return $m.Value }
        return ([string][char]$code)
    }
    return [regex]::Replace($Json, '\\u([0-9a-fA-F]{4})', $evaluator)
}

function Get-ConfigOverlayJson {
    # Config hashtable -> pretty JSON text with readable (non-escaped) Japanese.
    param([object]$Data)

    $json = $Data | ConvertTo-Json -Depth 12
    return (ConvertFrom-JsonUnicodeEscape $json)
}
