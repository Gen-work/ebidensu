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
#   Remove-ConfigMetadataKeys   drop _README/_comment docs before runtime merge
#   Merge-ConfigHashtable       deep-merge overlay onto base (base wins structure)
#   New-ConfigOverlaySnapshot   operator-facing editable snapshot of a config
#   Get-ConfigOverlayGroups      grouped InitConfig editor/readme sections
#   Remove-ConfigEmptyArray     drop empty arrays (PS 5.1 serializes them as "")
#   ConvertFrom-JsonUnicodeEscape  turn \uXXXX (>=0x80) back into real chars
#   Get-ConfigOverlayReadmeText separate field guide for InitConfig
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


function Remove-ConfigMetadataKeys {
    # Strip human-readable metadata keys before merging an overlay into runtime
    # config. This lets verify_config.json stay valid JSON while carrying a small
    # _README pointer without leaking documentation into $Config.
    param([object]$Value)

    if ($Value -is [System.Collections.IDictionary]) {
        $out = @{}
        foreach ($k in @($Value.Keys)) {
            if ([string]$k -match '^_(README|COMMENT|COMMENTS|SCHEMA|HELP)$') { continue }
            $out[$k] = Remove-ConfigMetadataKeys $Value[$k]
        }
        return $out
    }

    if (($Value -is [System.Collections.IEnumerable]) -and ($Value -isnot [string])) {
        $list = @()
        foreach ($item in $Value) { $list += ,(Remove-ConfigMetadataKeys $item) }
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
    if ($h -is [hashtable]) { return (Remove-ConfigMetadataKeys $h) }
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
    # Build the operator-facing snapshot that InitConfig writes to
    # <WorkDir>\verify_config.json. Runtime-only bootstrap values are excluded,
    # but all editable workflow config (including PhaseOrder) is emitted so a
    # work folder can be refreshed when VerifyConfig.psd1 gains new keys.
    param([hashtable]$Config)

    $snap = @{}
    $snap['_README'] = @(
        'Clean JSON only: see verify_config.README.txt for field explanations.',
        'Precedence: CLI args > this JSON > VerifyConfig.psd1 > session fallback.',
        'Run .\VerifyTool.ps1 -Phase InitConfig -Interactive -- pick w to walk a group',
        'field-by-field (no path typing needed), or view/edit/delete/save by JSON path.'
    )

    $skipKeys = @{
        DefaultWorkDir = $true  # WorkDir must be known before this overlay can be loaded.
        Scripts        = $true  # Repository script wiring is shared tool structure.
        Aliases        = $true  # Command aliases are shared tool structure.
    }

    foreach ($k in @($Config.Keys | Sort-Object)) {
        if ($skipKeys.ContainsKey([string]$k)) { continue }
        $snap[$k] = $Config[$k]
    }

    return (Remove-ConfigEmptyArray $snap)
}

function Get-ConfigOverlayGroups {
    # Group definitions used by the InitConfig editor/readme. A key can appear
    # in more than one group when operators commonly think about it in multiple
    # ways (for example Mail.EvidenceFolder is both mail text and a path).
    return @(
        @{ Key = 'intro'; Label = 'Introduction / README'; Paths = @('_README') },
        @{ Key = 'phase'; Label = 'Phase order / labels / progress fields'; Paths = @('PhaseOrder') },
        @{ Key = 'snap';  Label = 'Snap size / waits / capture geometry'; Paths = @('Window','Timing','Hm','Mq','Df','Mark') },
        @{ Key = 'excel'; Label = 'Excel workbook / replace / review / check sheet'; Paths = @('Workbook','ExcelSnap','Review','Replace','CheckSheet','SendVsGift') },
        @{ Key = 'wbs';   Label = 'WBS / mapping / compare helpers'; Paths = @('DefaultOwner','ExpectedTime','Align') },
        @{ Key = 'path';  Label = 'Paths / folders / external tools'; Paths = @('Paths','Clone','Align.J4BaseDir','Df.ExePath','Df.GiftDataDir','Df.GfixDataDir','Review.EvidenceDir','Mail.EvidenceFolder','Mail.CheckSheetFolder','CheckSheet.Path','DeliverFiles') },
        @{ Key = 'mail';  Label = 'Mail / reviewer / delivery'; Paths = @('Reviewer','Mail','DeliverFiles','CheckSheet') },
        @{ Key = 'all';   Label = 'All editable JSON variables'; Paths = @('*') }
    )
}


function Get-ConfigOverlayReadmeText {
    param([string]$OverlayName = 'verify_config.json')

    $lines = @(
        '# verify_config.json field guide',
        '',
        ('This text is generated next to {0}. Keep the JSON itself clean and valid.' -f $OverlayName),
        '',
        'Precedence',
        '- CLI arguments win first.',
        ('- Work-folder {0} overrides VerifyConfig.psd1 defaults.' -f $OverlayName),
        '- verify_session.json is only a last-used fallback/cache for values that still support it.',
        '',
        'Important rules',
        '- Do not put WorkDir in this JSON. The tool must know WorkDir before it can load this file.',
        '- Standard JSON has no // or /* */ comments. Use this README for comments.',
        '- Save as UTF-8. Japanese strings are OK.',
        '- Run .\VerifyTool.ps1 -Phase InitConfig -Interactive to view groups, edit values, delete keys, and confirm save.',
        '- In the editor, pick w to WALK a group: it prompts field-by-field (Enter=keep,',
        '  value=set, -del=delete, q=stop) so you never have to type a JSON path yourself.',
        '  v/e/d still take a manual path when you already know exactly what to touch.',
        '',
        'Groups in the InitConfig editor',
        '- intro: _README introduction lines shown at the top of the JSON.',
        '- phase: PhaseOrder labels, order, fields, and bit values.',
        '- snap: Window/Timing/Hm/Mq/Df/Mark capture size and geometry settings.',
        '- excel: Workbook/ExcelSnap/Review/Replace/CheckSheet/SendVsGift Excel-related settings.',
        '- wbs: DefaultOwner/ExpectedTime/Align settings related to mapping/WBS/precheck.',
        '- path: Paths, Clone, evidence folders, tool paths, and delivery destinations.',
        '- mail: Reviewer/Mail/DeliverFiles/CheckSheet hand-off settings.',
        '',
        'Common fields',
        '- DefaultOwner: owner suffix for mapping_<Owner>.csv and operator name used by phases.',
        '- Workbook.ExcelPrefix: fixed project prefix before _<Excel_NAME>. Example: J4 review title (REQ-000xxxxx_GIFT project).',
        '- Window.Width / Height / CropPx / NoResize: browser screenshot window and crop behavior for HM/MQ/Jenkins snapshots.',
        '- Review.EvidenceDir: evidence workbook folder. Relative paths are based on WorkDir.',
        '- Review.CursorCell: initial cell for visual review.',
        '- Clone.SourceDir: source folder used by Clone. If Align.J4BaseDir is blank, Align can reuse this path.',
        '- Align.J4BaseDir: J4 baseline folder for Align. Set only when different from Clone.SourceDir.',
        '- Align.HostSystemTypes: FROM_sys / TO_sys values treated as Host; empty means auto/legacy fallback.',
        '- CheckSheet.Path: shared review check sheet workbook path.',
        '- Reviewer.* and Mail.*: Outlook draft recipient, subject and body templates.',
        '- Df.*: df.exe path, capture mode, region, crop and data file lookup.',
        '- Mark.Boxes: red rectangle definitions per snap folder.',
        '- DeliverFiles.*: J4 delivery destinations + local BackupJ4 folder override.',
        '',
        'Workbook prefix note',
        '- New mappings no longer generate Excel_Prefix.',
        '- Existing mapping rows with legacy Excel_Prefix still override Workbook.ExcelPrefix for compatibility or rare per-workbook exceptions.'
    )
    return ($lines -join "`r`n")
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
