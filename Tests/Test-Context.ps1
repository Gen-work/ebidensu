#Requires -Version 5.1
# Test-Context.ps1 -- kernel/Context.ps1 ({{}} template evaluation, P1-01)
# and kernel/Key.ps1 (key display / keySafe normalization, P1-27 seed).
# Pure: fixtures only, no files.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Context.ps1')

Reset-Tests 'Context'

# ---------------------------------------------------------------- Key.ps1
$fw = [string]([char]0xFF21) + [char]0xFF22 + [char]0xFF11 + [char]0x3000 + [char]0x65E5   # 'AB1' full-width + ideographic space + a kanji
$hw = ConvertTo-EbiHalfWidth -Value $fw
Assert-Equal ('AB1 ' + [string][char]0x65E5) $hw 'half-width: full-width ASCII and U+3000 fold, kanji untouched'
Assert-True  (Test-EbiContainsFullWidth -Value $fw)      'full-width detected'
Assert-True  (-not (Test-EbiContainsFullWidth -Value 'ABC')) 'plain ASCII is not full-width'
Assert-Equal '' (ConvertTo-EbiHalfWidth -Value '')       'half-width of empty is empty'

Assert-Equal 'a_b_c_d_e_f_g_h_i_j' (ConvertTo-EbiKeySafeSegment -Value 'a\b/c:d*e?f"g<h>i|j') 'keySafe: every Windows-illegal char becomes _'
Assert-Equal 'x_y' (ConvertTo-EbiKeySafeSegment -Value ('x' + [char]9 + 'y')) 'keySafe: control chars become _'
Assert-Equal 'ABC123' (ConvertTo-EbiKeySafeSegment -Value ([string][char]0xFF21 + [char]0xFF22 + [char]0xFF23 + '123')) 'keySafe: full-width folded first'

$row = @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB/A'; Excel_NAME = 'W1' }
Assert-Equal 'ABC123' (Get-EbiKeyDisplay -Item $row -KeyColumns @('Correl_ID_S')) 'key display: single column is the raw value'
Assert-Equal 'ABC123 / JOB/A' (Get-EbiKeyDisplay -Item $row -KeyColumns @('Correl_ID_S', 'JOB_NAME')) 'key display: composite joined with " / "'
Assert-Equal 'ABC123_JOB_A' (ConvertTo-EbiKeySafe -Item $row -KeyColumns @('Correl_ID_S', 'JOB_NAME')) 'keySafe: composite joined with _, illegal chars replaced'
Assert-Equal 'A_B_C' (ConvertTo-EbiKeySafe -Item @{ a = 'A_B'; b = 'C' } -KeyColumns @('a', 'b')) 'keySafe: the documented collision shape (A_B + C)'
Assert-Equal 'A_B_C' (ConvertTo-EbiKeySafe -Item @{ a = 'A'; b = 'B_C' } -KeyColumns @('a', 'b')) 'keySafe: ... equals (A + B_C) -- table.load must detect this'
Assert-Equal 'ABC123_' (ConvertTo-EbiKeySafe -Item $row -KeyColumns @('Correl_ID_S', 'Missing')) 'keySafe: a missing key column reads as empty, never throws'

# ---------------------------------------------------------------- scope fixture
$profile = @{
    window = @{ width = 1050; height = 761 }
    pages  = @{
        transferStatus = @{ role = 'list'; url = 'https://h/ts'; timeoutSec = 12; crop = @{ left = 6; top = 6 }
                            fingerprint = @{ ok = @('a', 'b') } }
        fileList       = @{ role = 'list'; url = 'https://h/fl' }
    }
    grammar = @{ transferStatus = @{ parser = 'delimited'; delimiter = "`t" } }
    rules   = @{ transferStatus = @{ rules = @( @{ field = 'recvTime'; op = 'within'; value = '{{run.timeWindow}}' } ); default = 'ok' } }
    deep    = @{ inner = '{{profile.deeper}}' }
    deeper  = @{ z = '{{vars.side}}' }
    worklist = @{ file = 'mapping_{{run.operator}}.csv' }
}
$run   = @{ runId = 'r1'; operator = 'misaki'; timeWindow = @{ from = '2026-09-22T09:00:00'; to = '2026-09-22T10:00:00' } }
$item  = @{ Correl_ID_S = 'ABC123'; JOB_NAME = 'JOB_A'; key = 'REALKEY'; Excel_NAME = 'W1' }
$steps = @{
    shot            = @{ path = 'capture/x.png'; width = 1280; height = 800; rect = @{ W = 10 } }
    'refocus.shot'  = @{ path = 'capture/y.png' }
    gate            = @{ ok = $true; skipped = $true; code = $null; action = $null }
}
$scope = New-EbiTemplateScope -Vars @{ side = 'before'; n = 3; flag = $true } -Profile $profile -PageName 'transferStatus' `
    -Run $run -Item $item -Steps $steps -KeyColumns @('Correl_ID_S', 'JOB_NAME') -GroupColumn 'JOB_NAME'

function X { param($v) return (Expand-EbiTemplate -Value $v -Scope $scope) }

# ---------------------------------------------------------------- scopes
$r = X 'capture/{{vars.side}}_{{page.id}}/{{item.keySafe}}.png'
Assert-True  $r['ok'] 'interpolation ok'
Assert-Equal 'capture/before_transferStatus/ABC123_JOB_A.png' $r['value'] 'vars + page.id + item.keySafe interpolate into one string'
Assert-Equal 'https://h/ts' (X '{{page.url}}')['value']                   'page.X reads profile.pages[<bound page>]'
Assert-Equal 6 (X '{{page.crop.left}}')['value']                          'page.X.Y walks deeper'
Assert-Equal 'delimited' (X '{{page.grammar}}')['value']['parser']        'page.grammar aliases profile.grammar[<page>]'
Assert-Equal 'ok' (X '{{page.rules}}')['value']['default']                'page.rules aliases profile.rules[<page>]'
Assert-Equal 'https://h/fl' (X '{{profile.pages.fileList.url}}')['value'] 'an unbound page is still reachable by full profile path'
Assert-Equal 'r1' (X '{{run.runId}}')['value']                            'run scope'
Assert-Equal 'ABC123' (X '{{item.Correl_ID_S}}')['value']                 'item column'
Assert-Equal 'ABC123 / JOB_A' (X '{{item.key}}')['value']                 'item.key is the derived display form...'
Assert-True  ((X '{{item.key}}')['value'] -ne 'REALKEY')                  '...and wins over a real column literally named key'
Assert-Equal 'JOB_A' (X '{{item.group}}')['value']                        'item.group reads the declared group column'
Assert-Equal 'capture/x.png' (X '{{steps.shot.out.path}}')['value']       'steps.<id>.out.<field>'
Assert-Equal 10 (X '{{steps.shot.out.rect.W}}')['value']                  'steps field path walks deeper'
Assert-Equal 'capture/y.png' (X '{{steps.refocus.shot.out.path}}')['value'] 'a step id containing "." is cut at the .out. boundary (7.4)'

# ---------------------------------------------------------------- type preservation
$r = X '{{profile.window.width}}'
Assert-True  ($r['value'] -is [int] -and $r['value'] -eq 1050)            'a whole-value template keeps an int'
Assert-True  ((X '{{vars.flag}}')['value'] -is [bool])                     'a whole-value template keeps a bool'
Assert-True  ((X '{{profile.window}}')['value'] -is [hashtable])          'a whole-value template keeps a hashtable'
Assert-True  ((X '{{page.fingerprint.ok}}')['value'] -is [array])         'a whole-value template keeps an array'
Assert-True  ((X ' {{vars.n}} ')['value'] -is [int])                      'surrounding whitespace still counts as whole-value'
Assert-Equal 'n=3' (X 'n={{vars.n}}')['value']                            'interpolated int becomes text'
Assert-Equal 'f=true' (X 'f={{vars.flag}}')['value']                      'interpolated bool becomes lower-case text'

# ---------------------------------------------------------------- skipped step (5.1)
$r = X '{{steps.gate.out.code}}'
Assert-True  ($r['ok'] -and $null -eq $r['value'])                        'a when-skipped step field resolves to null, not an error'
Assert-Equal 'True' ([string](X '{{steps.gate.out.skipped}}')['value'])   'skipped=true is readable'
Assert-Equal 'code=' (X 'code={{steps.gate.out.code}}')['value']          'null interpolates as empty text'

# ---------------------------------------------------------------- 4.3 one-pass subtree evaluation
$r = X '{{page.rules}}'
Assert-Equal '2026-09-22T09:00:00' $r['value']['rules'][0]['value']['from'] 'templates inside a profile subtree are evaluated once on the way out'
$r = X '{{profile.deep}}'
Assert-Equal '{{vars.side}}' $r['value']['inner']['z']                    'but the result is not parsed a second time (no second level)'
Assert-Equal 'mapping_misaki.csv' (X '{{profile.worklist.file}}')['value'] 'a profile string with a template is evaluated (P0-R11 file template)'

# ---------------------------------------------------------------- escape + nesting
Assert-Equal '{{literal}} and before' (X '\{\{literal}} and {{vars.side}}')['value'] '\{\{ is a literal {{ next to a real token'
Assert-Equal '{{x}}' (X '\{\{x}}')['value']                               'an escaped-only string has no tokens'
$r = X '{{profile.pages.{{vars.page}}.url}}'
Assert-True  (-not $r['ok'] -and $r['error'] -eq 'nested_template')       'nested templates are refused'
$r = X '{{vars.side'
Assert-True  (-not $r['ok'] -and $r['error'] -eq 'nested_template')       'an unterminated {{ is refused'

# ---------------------------------------------------------------- failures name the segment
$r = X '{{foo.bar}}'
Assert-Equal 'unknown_scope' $r['error']                                  'unknown prefix -> unknown_scope'
Assert-Equal 'foo' $r['segment']                                          '...naming the prefix'
$r = X '{{vars.nope}}'
Assert-Equal 'missing_segment' $r['error']                                'missing key -> missing_segment'
Assert-Equal 'nope' $r['segment']                                         '...naming the missing piece'
$r = X '{{profile.window.width.px}}'
Assert-Equal 'px' $r['segment']                                           'walking into a scalar names the extra segment'
$r = X '{{steps.ghost.out.path}}'
Assert-Equal 'step_not_found' $r['error']                                 'unknown step -> step_not_found'
Assert-Equal 'ghost' $r['segment']                                        '...naming the step id'
$r = X '{{steps.shot.out.nope}}'
Assert-Equal 'missing_segment' $r['error']                                'unknown output field -> missing_segment'
$r = X '{{steps.shot.path}}'
Assert-Equal 'bad_steps_path' $r['error']                                 'steps reference without .out. -> bad_steps_path'
$r = X 'x{{profile.window}}'
Assert-Equal 'not_scalar' $r['error']                                     'an object inside a string -> not_scalar'
Assert-Equal 'profile.window' $r['path']                                  '...naming the reference'
$r = X '{{item.nocol}}'
Assert-Equal 'nocol' $r['segment']                                        'unknown item column names the column'
$noPage = New-EbiTemplateScope -Vars @{} -Profile $profile -Run $run -Item $item -Steps $steps
$r = Expand-EbiTemplate -Value '{{page.url}}' -Scope $noPage
Assert-Equal 'no_page' $r['error']                                        'page.X without a bound page -> no_page'
$noItem = New-EbiTemplateScope -Vars @{} -Profile $profile -PageName 'transferStatus' -Run $run -Steps $steps
$r = Expand-EbiTemplate -Value '{{item.key}}' -Scope $noItem
Assert-Equal 'item' $r['segment']                                         'item.X outside each names item'

# ---------------------------------------------------------------- containers
$r = X @{ path = '{{steps.shot.out.path}}'; left = '{{page.crop.left}}'; data = @{ row = @{ index = '{{vars.n}}' } } }
Assert-True  $r['ok'] 'hashtable with expands'
Assert-Equal 'capture/x.png' $r['value']['path']                          'hashtable: string value expanded'
Assert-True  ($r['value']['left'] -is [int])                              'hashtable: whole-value keeps type'
Assert-Equal 3 $r['value']['data']['row']['index']                        'hashtable: nested map expanded'
$r = X @('{{vars.side}}')
Assert-True  ($r['value'] -is [array] -and $r['value'].Count -eq 1)      'a one-element array stays a one-element array'
Assert-Equal 'before' $r['value'][0]                                      'array element expanded'
$r = X @{ a = 'x{{vars.nope}}' }
Assert-True  (-not $r['ok'] -and $r['segment'] -eq 'nope')                'a failure deep inside a container surfaces with its segment'
Assert-Equal 7 (X 7)['value']                                             'non-string scalars pass through'

# ---------------------------------------------------------------- lint helpers
Assert-True  (Test-EbiTemplateString -Value 'a {{vars.x}}')                'Test-EbiTemplateString: token found'
Assert-True  (-not (Test-EbiTemplateString -Value '\{\{no}}'))            'Test-EbiTemplateString: escape is not a token'
Assert-True  (-not (Test-EbiTemplateString -Value 5))                     'Test-EbiTemplateString: non-string is false'
$refs = @(Get-EbiTemplateReferences -Value @{ a = '{{vars.x}}/{{item.key}}'; b = @('{{steps.s.out.p}}') })
Assert-Equal 3 $refs.Count                                                'references collected across a container'
Assert-True  ($refs -contains 'steps.s.out.p')                            '...including inside arrays'

exit (Complete-Tests)
