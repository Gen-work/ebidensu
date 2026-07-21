#Requires -Version 5.1
$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'MappingInput.ps1')
Reset-Tests 'MappingInput'

$parsed = @(ConvertFrom-MappingIdText @('x AJODW001 y BJODJ9Z8', 'AJODJ001 duplicate'))
Assert-Equal 'AJODJ001,BJODJ9Z8' ($parsed -join ',') 'extracts JOD W/J series, normalizes W to J, and dedupes'
$invalid = @(ConvertFrom-MappingIdText @('JIDSJ001 XJODJ1234'))
Assert-Equal 0 $invalid.Count 'rejects non-JOD and overlong tokens'

$root = Join-Path ([IO.Path]::GetTempPath()) ('mapping_input_' + [guid]::NewGuid().ToString('N'))
try {
    New-Item -ItemType Directory -Path (Join-Path $root 'snap\GFIX_HM') -Force | Out-Null
    New-Item -ItemType Directory -Path (Join-Path $root 'snap\GIFT_HM') -Force | Out-Null
    Set-Content -LiteralPath (Join-Path $root 'snap\GFIX_HM\ID.txt') -Value 'CJODW123' -Encoding UTF8
    Set-Content -LiteralPath (Join-Path $root 'snap\GIFT_HM\ID.png') -Value 'fake' -Encoding ASCII
    $ocrCalled = $false
    $found = Get-MappingIdInput -WorkDir $root -OcrImage { param($p) $script:ocrCalled=$true; 'DJODJ999' }
    Assert-Equal 'CJODJ123' $found.Jobs[0] 'ID.txt is used and Excel W name is normalized'
    Assert-Equal 'text' $found.Kind 'ID.txt has priority over PNG OCR'
    Assert-True (-not $ocrCalled) 'OCR is not called when usable ID.txt exists'

    Remove-Item -LiteralPath (Join-Path $root 'snap\GFIX_HM\ID.txt') -Force
    $found2 = Get-MappingIdInput -WorkDir $root -OcrImage { param($p) 'DJODW999' }
    Assert-Equal 'DJODJ999' $found2.Jobs[0] 'GIFT ID.png is OCR fallback'
    Assert-Equal 'ocr' $found2.Kind 'PNG source reports OCR kind'
} finally { Remove-Item -LiteralPath $root -Recurse -Force -ErrorAction SilentlyContinue }

$sorted = @(Sort-MappingJobsByTemplateOrder -Jobs @('CJODJ003','CJODJ001','CJODJ002','OTHER') -TemplateJobs @('CJODW002','CJODJ003'))
Assert-Equal 'CJODJ002,CJODJ003,CJODJ001,OTHER' ($sorted -join ',') 'template jobs first in col-H order, unmatched jobs remain stable'
exit (Complete-Tests)
