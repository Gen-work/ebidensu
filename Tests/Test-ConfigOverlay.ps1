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

# --- Update-ConfigOverlayData : repair mode keeps operator values, adds only
#     missing default keys (deep), arrays/scalars stay atomic.
$existing = @{
    DefaultOwner = '9999'                      # operator-changed scalar: must survive
    Window       = @{ Width = 800 }            # sparse nested: keep Width, gain missing keys
    Align        = @{ HostSystemTypes = @('HOST') }  # operator array: must survive wholesale
}
$defaults = @{
    _README      = @('readme line')
    DefaultOwner = '0602'
    Window       = @{ Width = 1050; Height = 761 }
    Align        = @{ HostSystemTypes = @(); J4BaseDir = '' }
    Replace      = @{ GfixLogFontName = 'MS Gothic'; GfixLogFontSize = 11 }
}
$rep = Update-ConfigOverlayData -Existing $existing -Defaults $defaults
Assert-Equal '9999' $rep.Data['DefaultOwner']            'repair keeps operator scalar'
Assert-Equal 800    $rep.Data['Window']['Width']         'repair keeps nested operator value'
Assert-Equal 761    $rep.Data['Window']['Height']        'repair adds missing nested default'
Assert-Equal 'HOST' (@($rep.Data['Align']['HostSystemTypes'])[0]) 'repair keeps operator array wholesale'
Assert-Equal ''     $rep.Data['Align']['J4BaseDir']      'repair adds missing sibling of kept array'
Assert-Equal 'MS Gothic' $rep.Data['Replace']['GfixLogFontName'] 'repair adds whole missing section'
Assert-True ($rep.Data.ContainsKey('_README'))           'repair adds missing _README'
$addedPaths = @($rep.Added)
Assert-True ($addedPaths -contains 'Window.Height')      'added list reports dotted nested path'
Assert-True ($addedPaths -contains 'Replace')            'added list reports new top-level section'
Assert-True (-not ($addedPaths -contains 'DefaultOwner')) 'existing key not reported as added'

# A complete overlay (every default key already present) adds nothing.
$repNone = Update-ConfigOverlayData -Existing $defaults -Defaults $defaults
Assert-Equal 0 (@($repNone.Added).Count) 'repair of a complete overlay adds nothing'

# --- Get-ConfigOverlayGroups : editor exposes the requested group tags.
$groups = @(Get-ConfigOverlayGroups)
$groupKeys = @($groups | ForEach-Object { $_.Key })
foreach ($needed in @('intro','phase','snap','excel','wbs','path','mail','all')) {
    Assert-True ($groupKeys -contains $needed) ("group exists: {0}" -f $needed)
}

exit (Complete-Tests)
