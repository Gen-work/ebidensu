#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'ConfigOverlay.ps1')

Reset-Tests 'ConfigOverlay'

# --- ConvertFrom-ConfigJson : JSON objects become hashtables, arrays of
#     objects become arrays of hashtables (Mark.Boxes consumers need .ContainsKey).
$json = '{"Mark":{"Boxes":{"GIFT_HM":[{"OffsetX":1.5,"OffsetY":2.0,"Width":3,"Height":4}]}}}'
$h = ConvertFrom-ConfigJson $json
Assert-True ($h -is [hashtable]) 'parsed root is hashtable'
$box = $h['Mark']['Boxes']['GIFT_HM'][0]
Assert-True ($box -is [hashtable]) 'box object is a hashtable'
Assert-True ($box.ContainsKey('OffsetX')) 'box exposes ContainsKey(OffsetX)'
Assert-Equal '1.5' $box['OffsetX'] 'box OffsetX value preserved'

Assert-Equal 0 (ConvertFrom-ConfigJson '').Count 'blank JSON -> empty hashtable'
Assert-Equal 0 (ConvertFrom-ConfigJson '[1,2,3]').Count 'top-level array -> empty hashtable'

$meta = ConvertFrom-ConfigJson '{"_README":["help"],"Window":{"_comment":"doc","Width":123}}'
Assert-True (-not $meta.ContainsKey('_README')) 'metadata _README stripped from runtime config'
Assert-True (-not $meta['Window'].ContainsKey('_comment')) 'nested metadata _comment stripped from runtime config'
Assert-Equal '123' $meta['Window']['Width'] 'real config key survives metadata stripping'

# --- Merge-ConfigHashtable : nested hashtables merge, arrays/scalars replace.
$base = @{ A = 1; B = @{ X = 1; Y = 2 }; L = @(1, 2) }
$ov   = @{ B = @{ Y = 9; Z = 3 }; L = @(9); C = 5 }
$m = Merge-ConfigHashtable $base $ov
Assert-Equal 1 $m['A']        'A untouched'
Assert-Equal 1 $m['B']['X']   'B.X untouched (deep merge)'
Assert-Equal 9 $m['B']['Y']   'B.Y overridden'
Assert-Equal 3 $m['B']['Z']   'B.Z added'
Assert-Equal 5 $m['C']        'C added'
Assert-Equal 1 (@($m['L']).Count) 'array replaced wholesale, not concatenated'
Assert-Equal 9 (@($m['L'])[0])    'replacement array value'

# --- Remove-ConfigEmptyArray : empty arrays dropped, the rest kept.
$in  = @{ Keep = @(1, 2); Empty = @(); Nested = @{ E = @(); K = 'v' } }
$out = Remove-ConfigEmptyArray $in
Assert-True ($out.ContainsKey('Keep'))            'non-empty array kept'
Assert-True (-not $out.ContainsKey('Empty'))      'top-level empty array dropped'
Assert-True ($out['Nested'].ContainsKey('K'))     'nested scalar kept'
Assert-True (-not $out['Nested'].ContainsKey('E')) 'nested empty array dropped'

# --- ConvertFrom-JsonUnicodeEscape : >=0x80 decoded, <0x80 escapes kept.
#     Input is the ASCII \uXXXX escape form (what PS 5.1 ConvertTo-Json emits).
$esc      = '"\u3055\u3093"'
$expected = '"' + ([string][char]0x3055) + ([string][char]0x3093) + '"'
Assert-Equal $expected (ConvertFrom-JsonUnicodeEscape $esc) 'decodes escape >= 0x80'
Assert-Equal 'a\u0022b' (ConvertFrom-JsonUnicodeEscape 'a\u0022b') 'keeps ASCII escape (quote, < 0x80)'

# --- New-ConfigOverlaySnapshot : curated subset, no structural keys, no empties.
$cfg = @{
    DefaultOwner = '0602'
    Workbook = @{ ExcelPrefix = 'ProjectPrefix' }
    Mark   = @{ Boxes = @{ GIFT_HM = @( @{ OffsetX = 1.0; OffsetY = 2.0; Width = 3.0; Height = 4.0 } ); excel = @() } }
    Align  = @{ HostSystemTypes = @(); J4BaseDir = '' }
    SendVsGift = @{ Ocr = $false }
    PhaseOrder = @( @{ Key = 'InitConfig'; Label = 'config' } )
    DefaultWorkDir = 'C:\work'
    Scripts = @{ Foo = 'bar' }
    Aliases = @{ Config = 'InitConfig' }
}
$snap = New-ConfigOverlaySnapshot $cfg
Assert-True ($snap.ContainsKey('_README'))        'snapshot carries _README guidance'
Assert-True ($snap.ContainsKey('DefaultOwner'))   'snapshot carries DefaultOwner'
Assert-True ($snap.ContainsKey('Workbook'))       'snapshot carries Workbook prefix config'
Assert-True ($snap.ContainsKey('Mark'))           'snapshot carries Mark'
Assert-True ($snap.ContainsKey('SendVsGift'))     'snapshot carries SendVsGift config'
Assert-True ($snap.ContainsKey('PhaseOrder'))     'snapshot carries editable phase config'
Assert-True (-not $snap.ContainsKey('DefaultWorkDir')) 'snapshot excludes bootstrap DefaultWorkDir'
Assert-True (-not $snap.ContainsKey('Scripts'))   'snapshot excludes structural Scripts'
Assert-True (-not $snap.ContainsKey('Aliases'))   'snapshot excludes structural Aliases'
Assert-True ($snap['Mark']['Boxes'].ContainsKey('GIFT_HM'))      'non-empty box folder kept'
Assert-True (-not $snap['Mark']['Boxes'].ContainsKey('excel'))   'empty box folder dropped'
Assert-True (-not $snap['Align'].ContainsKey('HostSystemTypes')) 'empty HostSystemTypes dropped'
Assert-True ($snap.ContainsKey('_SCHEMA'))        'snapshot stamps a _SCHEMA field inventory'
Assert-True (@($snap['_SCHEMA']) -contains 'Workbook.ExcelPrefix') '_SCHEMA lists snapshot leaf paths'
Assert-True (-not (@($snap['_SCHEMA']) -contains '_README')) '_SCHEMA excludes metadata keys'

# --- New-ConfigOverlaySnapshot : duplicated fields merge into the canonical
#     top-level J4EvidenceDir / Address; legacy values migrate in; the live
#     config hashtable is never mutated by the scrub.
$cfgDup = @{
    Mail         = @{ EvidenceFolder = '\\srv\j4'; CheckSheetFile = 'x.xlsx' }
    DeliverFiles = @{ J4EvidenceDir = ''; Backup = $false }
    Reviewer     = @{ DisplayName = 'D'; Address = 'a@b'; ShortName = 'S' }
}
$snapDup = New-ConfigOverlaySnapshot $cfgDup
Assert-Equal '\\srv\j4' $snapDup['J4EvidenceDir'] 'legacy Mail.EvidenceFolder migrates into canonical J4EvidenceDir'
Assert-Equal 'a@b'      $snapDup['Address']       'legacy Reviewer.Address migrates into canonical Address'
Assert-True (-not $snapDup['Mail'].ContainsKey('EvidenceFolder'))         'snapshot drops Mail.EvidenceFolder'
Assert-True (-not $snapDup['DeliverFiles'].ContainsKey('J4EvidenceDir'))  'snapshot drops DeliverFiles.J4EvidenceDir'
Assert-True (-not $snapDup['Reviewer'].ContainsKey('Address'))            'snapshot drops Reviewer.Address'
Assert-True ($cfgDup['Mail'].ContainsKey('EvidenceFolder'))               'live config not mutated by the scrub'
Assert-True ($cfgDup['Reviewer'].ContainsKey('Address'))                  'live Reviewer section not mutated'

# --- canonical resolution helpers : legacy wins when set, then canonical.
$resCfg = @{ J4EvidenceDir = '\\srv\top'; Mail = @{ EvidenceFolder = '' }; DeliverFiles = @{ J4EvidenceDir = '' } }
Assert-Equal '\\srv\top' (Get-ConfigJ4EvidenceDir $resCfg) 'top-level J4EvidenceDir used when legacy fields blank'
$resCfg2 = @{ J4EvidenceDir = '\\srv\top'; Mail = @{ EvidenceFolder = '\\srv\legacy' } }
Assert-Equal '\\srv\legacy' (Get-ConfigJ4EvidenceDir $resCfg2) 'legacy Mail.EvidenceFolder wins when set'
Assert-Equal '' (Get-ConfigJ4EvidenceDir @{}) 'no J4 dir configured -> empty string'
$addrCfg = @{ Address = 'top@x'; Reviewer = @{ Address = '' } }
Assert-Equal 'top@x' (Get-ConfigReviewerAddress $addrCfg) 'top-level Address used when legacy blank'
$addrCfg2 = @{ Address = 'top@x'; Reviewer = @{ Address = 'legacy@x' } }
Assert-Equal 'legacy@x' (Get-ConfigReviewerAddress $addrCfg2) 'legacy Reviewer.Address wins when set'

# --- Generator round-trip : readable Japanese + boxes survive a write/read cycle.
$jpJson = Get-ConfigOverlayJson $snap
$rt = ConvertFrom-ConfigJson $jpJson
Assert-Equal '0602' $rt['DefaultOwner'] 'round-trip DefaultOwner'
Assert-Equal 'ProjectPrefix' $rt['Workbook']['ExcelPrefix'] 'round-trip Workbook.ExcelPrefix'
$rtBox = $rt['Mark']['Boxes']['GIFT_HM'][0]
Assert-True ($rtBox.ContainsKey('OffsetX')) 'round-trip box stays a hashtable'

$jp = ([string][char]0x3055) + ([string][char]0x3093)
$jpData = @{ Mail = @{ Greeting = $jp } }
$jpOut = Get-ConfigOverlayJson $jpData
Assert-True ($jpOut.Contains($jp)) 'generated JSON keeps Japanese readable (not escaped)'
$jpRt = ConvertFrom-ConfigJson $jpOut
Assert-Equal $jp $jpRt['Mail']['Greeting'] 'Japanese value round-trips'

# --- Get-ConfigSchemaPaths : dotted leaf inventory; hashtables recurse,
#     arrays and scalars stay atomic, metadata keys are excluded.
$schemaData = @{
    _README      = @('doc')
    DefaultOwner = '0602'
    Window       = @{ Width = 1050; Height = 761 }
    Mail         = @{ BodyLines = @('a', 'b') }
    Empty        = @{}
}
$schemaPaths = @(Get-ConfigSchemaPaths $schemaData)
Assert-True ($schemaPaths -contains 'DefaultOwner')   'schema lists scalar leaf'
Assert-True ($schemaPaths -contains 'Window.Width')   'schema recurses into hashtables'
Assert-True ($schemaPaths -contains 'Mail.BodyLines') 'schema keeps arrays atomic'
Assert-True ($schemaPaths -contains 'Empty')          'empty hashtable is one leaf'
Assert-True (-not ($schemaPaths -contains '_README')) 'metadata keys excluded from schema'
Assert-True (-not ($schemaPaths -contains 'Window'))  'non-empty hashtable is not itself a leaf'

# --- Update-ConfigOverlayData : a stamp-less file is only STAMPED -- values
#     untouched, NOTHING added (the whole snapshot is never dumped into a
#     sparse operator file).
$existing = @{
    DefaultOwner = '9999'                      # operator-changed scalar: must survive
    Window       = @{ Width = 800 }            # deliberately sparse: must stay sparse
}
$defaults = @{
    _README      = @('readme line')
    DefaultOwner = '0602'
    Window       = @{ Width = 1050; Height = 761 }
    Align        = @{ HostSystemTypes = @(); J4BaseDir = '' }
    Replace      = @{ GfixLogFontName = 'MS Gothic'; GfixLogFontSize = 11 }
}
$rep = Update-ConfigOverlayData -Existing $existing -Defaults $defaults
Assert-True  ([bool]$rep.Stamped)                        'stamp-less overlay reports Stamped'
Assert-Equal 0 (@($rep.Added).Count)                     'stamp-less overlay gains no fields'
Assert-Equal '9999' $rep.Data['DefaultOwner']            'operator scalar untouched'
Assert-Equal 800    $rep.Data['Window']['Width']         'nested operator value untouched'
Assert-True (-not $rep.Data['Window'].ContainsKey('Height')) 'sparse file stays sparse (no Height)'
Assert-True (-not $rep.Data.ContainsKey('Replace'))      'whole snapshot NOT dumped into sparse file'
Assert-True (-not $rep.Data.ContainsKey('_README'))      'no readme injected into operator file'
Assert-True ($rep.Data.ContainsKey('_SCHEMA'))           '_SCHEMA stamp written'
Assert-True (@($rep.Data['_SCHEMA']) -contains 'Replace.GfixLogFontName') 'stamp lists current default fields'

# --- Update-ConfigOverlayData : with a stamp, ONLY fields newer than the
#     stamp are appended; a field the operator deleted (still in the stamp)
#     is never re-added; existing values survive.
$existing2 = @{
    DefaultOwner = '9999'
    Window       = @{ Width = 800 }            # Height deleted by the operator
    Align        = @{ HostSystemTypes = @('HOST') }
    _SCHEMA      = @('DefaultOwner', 'Window.Width', 'Window.Height', 'Align.HostSystemTypes')
}
$defaults2 = @{
    DefaultOwner = '0602'
    Window       = @{ Width = 1050; Height = 761 }
    Align        = @{ HostSystemTypes = @(); J4BaseDir = '' }
    Replace      = @{ GfixLogFontName = 'MS Gothic' }
}
$rep2 = Update-ConfigOverlayData -Existing $existing2 -Defaults $defaults2
Assert-True (-not [bool]$rep2.Stamped)                   'stamped overlay is repaired, not just stamped'
Assert-Equal '9999' $rep2.Data['DefaultOwner']           'repair keeps operator scalar'
Assert-Equal 800    $rep2.Data['Window']['Width']        'repair keeps nested operator value'
Assert-True (-not $rep2.Data['Window'].ContainsKey('Height')) 'deliberately deleted field NOT re-added'
Assert-Equal 'HOST' (@($rep2.Data['Align']['HostSystemTypes'])[0]) 'operator array survives wholesale'
Assert-Equal ''     $rep2.Data['Align']['J4BaseDir']     'new sibling field appended into existing section'
Assert-Equal 'MS Gothic' $rep2.Data['Replace']['GfixLogFontName'] 'new section field appended'
$addedPaths = @($rep2.Added)
Assert-True ($addedPaths -contains 'Align.J4BaseDir')             'added list reports dotted nested path'
Assert-True ($addedPaths -contains 'Replace.GfixLogFontName')     'added list reports new section leaf'
Assert-True (-not ($addedPaths -contains 'Window.Height'))        'deleted field not reported as added'
Assert-True (@($rep2.Data['_SCHEMA']) -contains 'Replace.GfixLogFontName') 'stamp refreshed with new field'
Assert-True (@($rep2.Data['_SCHEMA']) -contains 'Window.Height')  'stamp keeps previously known (deleted) field'

# A second repair right after adds nothing.
$rep3 = Update-ConfigOverlayData -Existing $rep2.Data -Defaults $defaults2
Assert-Equal 0 (@($rep3.Added).Count) 'second repair in a row adds nothing'

# --- Get-ConfigOverlayGroups : editor exposes the requested group tags.
$groups = @(Get-ConfigOverlayGroups)
$groupKeys = @($groups | ForEach-Object { $_.Key })
foreach ($needed in @('intro','phase','snap','excel','wbs','path','mail','all')) {
    Assert-True ($groupKeys -contains $needed) ("group exists: {0}" -f $needed)
}

# --- Schema-drift guard : every InitConfig snapshot field must be reachable
#     from a NAMED editor group, not only the catch-all "all". Get-ConfigOverlayGroups
#     is a hand-maintained index of VerifyConfig.psd1's sections; the
#     JSON/repair layer (New-ConfigOverlaySnapshot / Update-ConfigOverlayData)
#     picks up a new .psd1 top-level section automatically, but the grouped
#     field walker and the README stay silent about it until someone also
#     registers it here. This is what should have caught SnapVerify (added
#     v2.9.4) sitting reachable only via "all" for many releases -- re-run
#     this test after any VerifyConfig.psd1 structural change; a failure
#     names the section(s) that still need a Get-ConfigOverlayGroups entry.
$repoRoot   = Split-Path $here -Parent
$realConfig = Import-PowerShellDataFile -LiteralPath (Join-Path $repoRoot 'VerifyConfig.psd1')
$realSnapshot = New-ConfigOverlaySnapshot $realConfig

$namedGroupTopKeys = @{}
foreach ($g in @($groups | Where-Object { $_.Key -ne 'all' })) {
    foreach ($p in @($g.Paths)) {
        $namedGroupTopKeys[(([string]$p) -split '\.')[0]] = $true
    }
}
$unreachable = @()
foreach ($k in @($realSnapshot.Keys | Sort-Object)) {
    if ($k -eq '_README' -or $k -eq '_SCHEMA') { continue }
    if (-not $namedGroupTopKeys.ContainsKey([string]$k)) { $unreachable += [string]$k }
}
Assert-Equal '' ($unreachable -join ',') 'every real VerifyConfig.psd1 snapshot field is reachable from a named editor group'

# --- Repair safety against the REAL VerifyConfig.psd1 shape (Mark.Boxes
#     arrays-of-hashtables, PhaseOrder array, SnapVerify.Localize nested
#     hashtable, etc.), not just hand-built fixtures -- confirms running
#     -Phase InitConfig repair on a file that predates a whole section never
#     throws, never drops an operator value, and re-adds exactly that section.
$reducedSnapshot = @{}
foreach ($k in @($realSnapshot.Keys)) {
    if ($k -eq 'SnapVerify') { continue }
    $reducedSnapshot[$k] = $realSnapshot[$k]
}
$existingBeforeSnapVerify = @{
    DefaultOwner = 'op-set-value'
    _SCHEMA      = @(Get-ConfigSchemaPaths $reducedSnapshot)
}
$realRepair = Update-ConfigOverlayData -Existing $existingBeforeSnapVerify -Defaults $realSnapshot
Assert-True (-not [bool]$realRepair.Stamped) 'real-defaults repair treats a pre-stamped file as already stamped, not stamp-less'
Assert-Equal 'op-set-value' $realRepair.Data['DefaultOwner'] 'real-defaults repair keeps the operator value untouched'
Assert-True ($realRepair.Data.ContainsKey('SnapVerify')) 'real-defaults repair re-adds a whole section the file predates'
Assert-True (@($realRepair.Added) -contains 'SnapVerify.Enabled') 'added-fields list names the real SnapVerify.Enabled leaf path'

$secondPass = Update-ConfigOverlayData -Existing $realRepair.Data -Defaults $realSnapshot
Assert-Equal 0 (@($secondPass.Added).Count) 'repairing twice in a row against real defaults is idempotent'
Assert-Equal 'op-set-value' $secondPass.Data['DefaultOwner'] 'second repair pass still keeps the operator value'

exit (Complete-Tests)
