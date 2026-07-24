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
#                               (canonical J4EvidenceDir/Address; stamps _SCHEMA)
#   Get-ConfigSchemaPaths       dotted leaf-path list of a config tree (the
#                               _SCHEMA stamp repair mode compares against)
#   Update-ConfigOverlayData    repair mode: add ONLY fields the tool gained
#                               since the overlay's _SCHEMA stamp was written --
#                               a sparse operator file stays sparse
#   Get-ConfigJ4EvidenceDir     canonical J4 evidence folder (legacy
#                               DeliverFiles.J4EvidenceDir / Mail.EvidenceFolder
#                               still win when set)
#   Get-ConfigReviewerAddress   canonical reviewer mail address (legacy
#                               Reviewer.Address still wins when set)
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

function Copy-ConfigSectionWithoutKeys {
    # Shallow-copy a config section minus the named keys. Used by the snapshot
    # builder so dropping a deprecated duplicate field never mutates the live
    # runtime $Config the section hashtable is shared with.
    param([hashtable]$Section, [string[]]$RemoveKeys)

    $out = @{}
    foreach ($k in @($Section.Keys)) {
        if (@($RemoveKeys) -contains [string]$k) { continue }
        $out[$k] = $Section[$k]
    }
    return $out
}

function Get-ConfigOverlayPathValue {
    # Hashtable-only dotted-path read (schema paths never index into arrays).
    param([hashtable]$Data, [string]$Path)

    $cur = $Data
    foreach ($part in @($Path -split '\.')) {
        if (-not ($cur -is [hashtable]) -or -not $cur.ContainsKey($part)) { return $null }
        $cur = $cur[$part]
    }
    return $cur
}

function Test-ConfigOverlayPathPresent {
    # $true when every segment of the dotted path exists (hashtable walk).
    param([hashtable]$Data, [string]$Path)

    $cur = $Data
    foreach ($part in @($Path -split '\.')) {
        if (-not ($cur -is [hashtable]) -or -not $cur.ContainsKey($part)) { return $false }
        $cur = $cur[$part]
    }
    return $true
}

function Set-ConfigOverlayPathValue {
    # Hashtable-only dotted-path write; creates missing parent hashtables.
    # Returns $false (and writes nothing) when an existing ancestor is not a
    # hashtable -- the operator holds a non-object there and it must be kept.
    param([hashtable]$Data, [string]$Path, [object]$Value)

    $parts = @($Path -split '\.')
    $cur = $Data
    for ($i = 0; $i -lt ($parts.Count - 1); $i++) {
        $part = $parts[$i]
        if (-not $cur.ContainsKey($part)) { $cur[$part] = @{} }
        elseif (-not ($cur[$part] -is [hashtable])) { return $false }
        $cur = $cur[$part]
    }
    $cur[$parts[$parts.Count - 1]] = $Value
    return $true
}

function Save-ConfigOverlayValue {
    # Persist ONE dotted-path value into a work folder's sparse
    # verify_config.json, creating the file when absent. Everything else in
    # the file -- operator values, metadata keys (_README/_SCHEMA/...) -- is
    # preserved byte-for-byte at the data level (re-serialised, not re-typed).
    # Used by VerifyTool's first-run prompts so a project-scoped answer (e.g.
    # CheckSheet.Path) lands in the WORK FOLDER config instead of the global
    # verify_session.json, where it would leak into unrelated work folders.
    # Returns $true on success; $false (file untouched) when the existing
    # file cannot be parsed, an ancestor of Path is a non-object, or the
    # write itself fails (e.g. file open/locked in an editor).
    param([string]$OverlayPath, [string]$Path, [object]$Value)

    if ([string]::IsNullOrWhiteSpace($OverlayPath) -or [string]::IsNullOrWhiteSpace($Path)) { return $false }

    $data = @{}
    if (Test-Path -LiteralPath $OverlayPath) {
        try {
            $raw = Get-Content -LiteralPath $OverlayPath -Raw -Encoding UTF8
            if (-not [string]::IsNullOrWhiteSpace($raw)) {
                # Deliberately NOT ConvertFrom-ConfigJson: that strips the
                # _README/_SCHEMA metadata keys, which must survive a rewrite.
                $parsed = ConvertTo-ConfigHashtable ($raw | ConvertFrom-Json)
                if ($parsed -is [hashtable]) { $data = $parsed } else { return $false }
            }
        } catch { return $false }
    }

    if (-not (Set-ConfigOverlayPathValue $data $Path $Value)) { return $false }

    try {
        $json = Get-ConfigOverlayJson $data
        [System.IO.File]::WriteAllText($OverlayPath, $json, (New-Object System.Text.UTF8Encoding($false)))
        return $true
    } catch { return $false }
}

function Get-ConfigFirstNonBlank {
    # First non-blank string value among the dotted paths, else ''.
    param([hashtable]$Config, [string[]]$Paths)

    if ($null -eq $Config) { return '' }
    foreach ($p in @($Paths)) {
        $v = Get-ConfigOverlayPathValue $Config $p
        if ($null -ne $v -and -not [string]::IsNullOrWhiteSpace([string]$v)) { return ([string]$v).Trim() }
    }
    return ''
}

function Get-ConfigJ4EvidenceDir {
    # The ONE J4 evidence folder consumers should read (DeliverFiles /
    # BackupJ4 / DeliverMail body path / Review j4 option). Canonical field:
    # top-level J4EvidenceDir. The legacy duplicates DeliverFiles.J4EvidenceDir
    # and Mail.EvidenceFolder still WIN when non-empty so configs written
    # before the merge keep behaving unchanged; InitConfig no longer emits them.
    param([hashtable]$Config)

    return (Get-ConfigFirstNonBlank $Config @('DeliverFiles.J4EvidenceDir', 'Mail.EvidenceFolder', 'J4EvidenceDir'))
}

function Get-ConfigReviewerAddress {
    # The reviewer mail address (Outlook To). Canonical field: top-level
    # Address. Legacy Reviewer.Address still wins when set (old configs keep
    # working); InitConfig no longer emits it.
    param([hashtable]$Config)

    return (Get-ConfigFirstNonBlank $Config @('Reviewer.Address', 'Address'))
}

function Get-ConfigSchemaPaths {
    # Dotted leaf-path inventory of a config tree: hashtables recurse; arrays
    # and scalars are one atomic leaf (matching repair's add semantics);
    # metadata keys (_README/_SCHEMA/...) are excluded. This is what the
    # _SCHEMA stamp stores so repair can tell "field the tool gained since the
    # overlay was written" apart from "field the operator deliberately removed".
    param([hashtable]$Data)

    function Get-SchemaPathsWorker([hashtable]$Cur, [string]$Prefix, $OutList) {
        foreach ($k in @($Cur.Keys | Sort-Object)) {
            if ([string]$k -match '^_') { continue }
            $path = if ([string]::IsNullOrEmpty($Prefix)) { [string]$k } else { ('{0}.{1}' -f $Prefix, $k) }
            if (($Cur[$k] -is [hashtable]) -and ($Cur[$k].Count -gt 0)) {
                Get-SchemaPathsWorker $Cur[$k] $path $OutList
            } else {
                $OutList.Add($path)
            }
        }
    }

    # Plain array return (never ,@(...)): callers wrap the call in @(...),
    # and that combination NESTS in PS 5.1 (repo convention, v2.8.1).
    $list = New-Object System.Collections.Generic.List[string]
    if ($null -ne $Data) { Get-SchemaPathsWorker $Data '' $list }
    return @($list.ToArray() | Sort-Object)
}

function New-ConfigOverlaySnapshot {
    # Build the operator-facing snapshot that InitConfig writes to
    # <WorkDir>\verify_config.json. Runtime-only bootstrap values are excluded,
    # but all editable workflow config (including PhaseOrder) is emitted so a
    # work folder can be refreshed when VerifyConfig.psd1 gains new keys.
    # The duplicated fields are merged into their canonical single field
    # (J4EvidenceDir / Address; any legacy value migrates in), and a _SCHEMA
    # field inventory is stamped so a later repair run adds only fields the
    # tool gained after this write.
    param([hashtable]$Config)

    $snap = @{}
    $snap['_README'] = @(
        'Clean JSON only: see verify_config.README.txt for field explanations.',
        'Precedence: CLI args > this JSON > VerifyConfig.psd1 > session fallback.',
        'Re-running -Phase InitConfig on this existing file REPAIRS it: your settings',
        'are kept as-is and only config fields the tool gained since the last write',
        'are appended (f=Force regenerates the full snapshot instead). Repair relies',
        'on the _SCHEMA field inventory below -- leave _SCHEMA alone (ignored at runtime).',
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

    # Merge the duplicated fields into their canonical top-level single field.
    # A value set in a legacy duplicate migrates into the canonical field; the
    # duplicates themselves are dropped from the snapshot (section hashtables
    # are copied first -- $snap shares them with the live $Config).
    $snap['J4EvidenceDir'] = Get-ConfigJ4EvidenceDir $Config
    $snap['Address']       = Get-ConfigReviewerAddress $Config
    if ($snap['Mail'] -is [hashtable])         { $snap['Mail']         = Copy-ConfigSectionWithoutKeys $snap['Mail'] @('EvidenceFolder') }
    if ($snap['DeliverFiles'] -is [hashtable]) { $snap['DeliverFiles'] = Copy-ConfigSectionWithoutKeys $snap['DeliverFiles'] @('J4EvidenceDir') }
    if ($snap['Reviewer'] -is [hashtable])     { $snap['Reviewer']     = Copy-ConfigSectionWithoutKeys $snap['Reviewer'] @('Address') }

    $clean = Remove-ConfigEmptyArray $snap
    $clean['_SCHEMA'] = @(Get-ConfigSchemaPaths $clean)
    return $clean
}

function Update-ConfigOverlayData {
    # REPAIR/UPDATE an existing overlay in place of a full regenerate.
    # The overlay's _SCHEMA stamp (written by New-ConfigOverlaySnapshot) lists
    # every field path that existed when the file was last written. Repair
    # adds ONLY the $Defaults fields that are NOT in that stamp -- fields the
    # tool gained since -- so a sparse, hand-trimmed operator file stays
    # sparse and a deliberately deleted field is never re-added. Existing
    # values are never overwritten. A file WITHOUT a _SCHEMA stamp (written
    # by an older InitConfig, or by hand) is only STAMPED on its first repair
    # -- nothing is added, so the whole snapshot is never dumped into it;
    # use f=Force / the interactive editor to pull in the full field set.
    # Arrays and scalars stay atomic, matching Merge-ConfigHashtable.
    # Returns @{ Data = <hashtable>; Added = <string[] dotted paths>; Stamped = <bool> }.
    param([hashtable]$Existing, [hashtable]$Defaults)

    $data = @{}
    if ($null -ne $Existing) { $data = $Existing }
    $defaultPaths = @(Get-ConfigSchemaPaths $Defaults)

    $known = $null
    if ($data.ContainsKey('_SCHEMA')) {
        $known = @{}
        foreach ($p in @($data['_SCHEMA'])) {
            if (-not [string]::IsNullOrWhiteSpace([string]$p)) { $known[[string]$p] = $true }
        }
    }

    if ($null -eq $known) {
        # Legacy file without a stamp: record the current field inventory and
        # change nothing else. From the next run on, repair can tell genuinely
        # new fields apart from fields this file simply never carried.
        $data['_SCHEMA'] = $defaultPaths
        return @{ Data = $data; Added = @(); Stamped = $true }
    }

    $added = New-Object System.Collections.Generic.List[string]
    foreach ($path in $defaultPaths) {
        if ($known.ContainsKey($path)) { continue }
        if (Test-ConfigOverlayPathPresent $data $path) { continue }   # operator already added it by hand
        $value = Get-ConfigOverlayPathValue $Defaults $path
        if (Set-ConfigOverlayPathValue $data $path $value) { $added.Add($path) }
    }

    # Refresh the stamp as the UNION of what was known and the current
    # defaults, so a field the operator deleted stays "known" and is never
    # re-added by a later repair.
    $stamp = @{}
    foreach ($p in @($known.Keys)) { $stamp[[string]$p] = $true }
    foreach ($p in $defaultPaths)  { $stamp[[string]$p] = $true }
    $data['_SCHEMA'] = @($stamp.Keys | Sort-Object)

    return @{ Data = $data; Added = $added.ToArray(); Stamped = $false }
}

function Get-ConfigOverlayGroups {
    # Group definitions used by the InitConfig editor/readme. A key can appear
    # in more than one group when operators commonly think about it in multiple
    # ways (for example Mail.EvidenceFolder is both mail text and a path).
    return @(
        @{ Key = 'intro'; Label = 'Introduction / README'; Paths = @('_README') },
        @{ Key = 'phase'; Label = 'Phase order / labels / progress fields'; Paths = @('PhaseOrder') },
        @{ Key = 'snap';  Label = 'Snap size / waits / capture geometry / NG detection'; Paths = @('Window','Timing','Hm','Mq','Df','Mark','SnapVerify') },
        @{ Key = 'excel'; Label = 'Excel workbook / replace / review / check sheet'; Paths = @('Workbook','ExcelSnap','Review','Replace','CheckSheet','SendVsGift','GfixLog','ProcessTime') },
        @{ Key = 'wbs';   Label = 'WBS / mapping / compare helpers'; Paths = @('DefaultOwner','ExpectedTime','Align') },
        @{ Key = 'path';  Label = 'Paths / folders / external tools'; Paths = @('Paths','Clone','Align.J4BaseDir','Df.ExePath','Df.GiftDataDir','Df.GfixDataDir','Review.EvidenceDir','J4EvidenceDir','Mail.CheckSheetFolder','CheckSheet.Path','DeliverFiles') },
        @{ Key = 'mail';  Label = 'Mail / reviewer / delivery'; Paths = @('Address','J4EvidenceDir','Reviewer','Mail','DeliverFiles','CheckSheet') },
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
        ('- Re-running -Phase InitConfig on an existing {0} REPAIRS it by default:' -f $OverlayName),
        '  your settings are kept as-is and ONLY config fields the tool gained since the',
        '  file was last written are appended -- a sparse file stays sparse. Repair uses',
        '  the hidden _SCHEMA field inventory in the JSON; leave _SCHEMA alone (it is',
        '  ignored at runtime). A file without _SCHEMA is stamped on its first repair',
        '  (nothing added). Use the f=Force option to regenerate the full snapshot',
        '  (keeps a .bak).',
        '- Run .\VerifyTool.ps1 -Phase InitConfig -Interactive to view groups, edit values, delete keys, and confirm save.',
        '- In the editor, pick w to WALK a group: it prompts field-by-field (Enter=keep,',
        '  value=set, -d=delete, q=stop), then offers the next group / save when the',
        '  group is done -- you never have to type a JSON path yourself.',
        '  v/e/d still take a manual path when you already know exactly what to touch.',
        '- If saving fails (e.g. the JSON is open in an editor), nothing is lost: close',
        '  the file, then r=retry; Enter keeps your edits in the editor menu.',
        '',
        'Groups in the InitConfig editor',
        '- intro: _README introduction lines shown at the top of the JSON.',
        '- phase: PhaseOrder labels, order, fields, and bit values.',
        '- snap: Window/Timing/Hm/Mq/Df/Mark/SnapVerify capture size, geometry, and NG-detection settings.',
        '- excel: Workbook/ExcelSnap/Review/Replace/CheckSheet/SendVsGift/GfixLog/ProcessTime Excel-related settings.',
        '- wbs: DefaultOwner/ExpectedTime/Align settings related to mapping/WBS/precheck.',
        '- path: Paths, Clone, evidence folders, tool paths, and delivery destinations.',
        '- mail: Reviewer/Mail/DeliverFiles/CheckSheet hand-off settings.',
        '',
        'Common fields',
        '- J4EvidenceDir: the ONE J4 evidence folder, used by DeliverFiles, BackupJ4,',
        '  the DeliverMail body path, and the review phases'' j4 option. (Legacy',
        '  DeliverFiles.J4EvidenceDir / Mail.EvidenceFolder from an old file still win',
        '  when set, but are no longer generated.)',
        '- Address: reviewer mail address (Outlook To). (Legacy Reviewer.Address from',
        '  an old file still wins when set, but is no longer generated.)',
        '- DefaultOwner: owner suffix for mapping_<Owner>.csv and operator name used by phases.',
        '- Workbook.ExcelPrefix: fixed project prefix before _<Excel_NAME>. Example: J4 review title (REQ-000xxxxx_GIFT project).',
        '- Window.Width / Height / CropPx / NoResize: browser screenshot window and crop behavior for HM/MQ/Jenkins snapshots.',
        '- Window.CropLeft / CropTop / CropRight / CropBottom: per-side crop override (px, -1 = inherit CropPx',
        '  for that side); applies globally to every HM/MQ/Jenkins snapshot.',
        '- Window.CropByFolder: per-snap-folder crop override, keyed by folder name (GIFT_HM, GFIX_HM, GIFT_MQ,',
        '  GIFT_Jenkins, GFIX_Jenkins, GIFT_noGfixfile); each entry may set any of Left/Top/Right/Bottom (px),',
        '  falling back to the resolved global Crop<Side>/CropPx value for any side left out.',
        '- Review.EvidenceDir: evidence workbook folder. Relative paths are based on WorkDir.',
        '- Review.CursorCell: initial cell for visual review.',
        '- Clone.SourceDir: source folder used by Clone. If Align.J4BaseDir is blank, Align can reuse this path.',
        '- Align.J4BaseDir: J4 baseline folder for Align. Set only when different from Clone.SourceDir.',
        '- Align.HostSystemTypes: FROM_sys / TO_sys values treated as Host; empty means auto/legacy fallback.',
        '- CheckSheet.Path: shared review check sheet workbook path.',
        '- Reviewer.* and Mail.*: Outlook draft recipient, subject and body templates.',
        '- Df.*: df.exe path, capture mode, region, crop and data file lookup.',
        '- Mark.Boxes: red rectangle definitions per snap folder. A box may add',
        '  Template (+ optional Tolerance/PadX/PadY) to try image-recognition',
        '  placement first (Mark.TemplateDir / mark_templates/), falling back to',
        '  OffsetX/OffsetY/Width/Height when no template match is found. Adding',
        '  StampImage alongside Template inserts that image at the match instead',
        '  of a rectangle, with NO fixed-offset fallback (no match = no stamp) --',
        '  wired for GIFT_noGfixfile (NoGfixHit.png -> already_exists.png).',
        '- Mark.NoteStamps: a SEPARATE stamp mechanism -- image inserted next to',
        '  a verifyNote annotation using its saved pixel rect (snap-time',
        '  loc.json/note.json), not a live Template match. Keyed by Folder (only',
        '  GIFT_noGfixfile today); Image/Column/RowOffset. See mark_templates/README.txt.',
        '- GfixLog.AutoHighlightWidth / HighlightPadCols: size the GFIX log',
        '  Command-row yellow highlight to the row''s actual text width instead',
        '  of the fixed HighlightColStart..HighlightColEnd range (still the cap).',
        '- Replace.GfixLogFontName / GfixLogFontSize: font + size forced on every',
        '  pasted GFIX log line (default MS Gothic / 11); blank name or size 0',
        '  leaves the workbook default. The same font/size drive the auto-width',
        '  highlight measurement in MarkGfix.',
        '- SnapVerify.*: instant NG detection for the HM/MQ/Jenkins snap phases.',
        '  Enabled is the master switch ($false = pure screenshot, no detection);',
        '  TimeCheck (off by default) adds a run-time-window prompt/compare;',
        '  NoGfixNoteColumn is the past-data annotation column (F4). Localize.*',
        '  (off by default) enables M5 pixel localisation for the Mark phase --',
        '  calibrate the Hm*/Mq* Row1Top/RowHeight/ColLeft/ColWidth fields with',
        '  Calibrate-HmGeometry.ps1 before turning Localize.Enabled on.',
        '- DeliverFiles.*: J4 delivery destinations + local BackupJ4 folder override.',
        '- ProcessTime.*: HM processing start/end/duration extraction (GIFT+GFIX) into',
        '  one workbook per output tag. AnchorCol matches Replace.ColAnchor;',
        '  OutputDirectory defaults to <WorkDir> (legacy OutputPath is honored as a',
        '  directory hint when OutputDirectory is blank); OutputSheetName defaults to a',
        '  [char] label. OutputMode ''Split'' (default) classifies each result row by the',
        '  first OutputTags entry (default JDL/JRV, e.g. add "JDS") found in its mapping',
        '  Excel_NAME and writes ''<label>(<Tag>).xlsx''; when no tag matches,',
        '  AutoDeriveTag (default true) derives the tag from the Excel_NAME''s own',
        '  ''?XXX????'' shape (chars 2-4: CJODWDEJ -> JOD) so an unlisted project',
        '  family still gets its own workbook; only a non-conforming name goes to',
        '  UnclassifiedTag (default Other) instead of aborting the run. OutputMode',
        '  ''Single'' ignores tags and writes one ''<label>.xlsx''. OutputDirectoryByTag',
        '  routes a tag to its own destination directory (e.g. a real J4 folder per',
        '  project) instead of every tag sharing OutputDirectory. OcrLanguage is the',
        '  secondary Windows OCR language pooled alongside en-US when no archived',
        '  snap\GIFT_HM|GFIX_HM\<correl>.txt is available. OcrPreprocess (default',
        '  true) upscales + grayscales + contrast-stretches every picture before',
        '  OCR (OcrPreprocessScale/OcrPreprocessContrast tune it) so thin digit',
        '  strokes (9 vs 3) read reliably; set false to OCR raw pixels.',
        '  OcrPreprocessBinarize/OcrPreprocessThreshold are RESERVED (v2.16.0):',
        '  carried through but not yet wired into the image pipeline. EmitCheckColumns',
        '  (default true) appends the on-sheet audit columns after the A..H data',
        '  columns -- I duration re-derivation (=E-D), J T/F compare of the written',
        '  vs re-derived duration, K record-count check (ProcessTimeCheck.ps1); set',
        '  false to write A..H data only. Stage (Ocr/Write/Both, default',
        '  Both) picks which of the extract-and-cache / write-the-output-workbooks',
        '  stages -Phase ProcessTime runs by default; CLI -Stage / the ''stage'' menu',
        '  option override it per run. ProcessTime_Inserted is a bitmask (1=OCR''d,',
        '  2=written, 3=both done), migrated from a pre-v2.15.0 plain 0/1 value.',
        '  ProcessTime.OldSnapVerify.* triages the finite backlog of OLD snaps',
        '  that have only a low-res PNG (no immune Ctrl+A .txt) and so fell back',
        '  to OCR (ja 9->3 misread): Enabled (master gate), EmitHyperlink (D1 --',
        '  correl-id cell links to the snap image for one-click human review),',
        '  EmitVerifyColumn (the 検証 column: txt / OCR-OK / 要確認 / 画像なし),',
        '  SnapDirPattern ({0}=GIFT/GFIX), RenderFont, and PixelDiff.Enabled --',
        '  the per-digit 3/9 image check, OFF until the Phase-0 separability',
        '  gate passes -- with PixelDiff.Threshold. CrossEngine.Enabled is a',
        '  reserved en-US-vs-ja digit cross-check (off).',
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
