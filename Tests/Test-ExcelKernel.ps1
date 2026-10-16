#Requires -Version 5.1
# Test-ExcelKernel.ps1 -- the pure parts of kernel/Excel.ps1. The COM
# entry points (Export-EbiShapePng, ...) need Excel and are checked on the
# office PC; here only the error line they report.

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here     = Split-Path $MyInvocation.MyCommand.Path
$repoRoot = Split-Path $here -Parent
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Join-Path $repoRoot 'kernel') 'Excel.ps1')

Reset-Tests 'ExcelKernel'

# ---------------------------------------------------------------- Format-EbiExcelError
$e1 = New-Object System.InvalidOperationException 'Paste method of Chart class failed'
Assert-Equal 'paste into chart failed: Paste method of Chart class failed [InvalidOperationException 0x80131509]' (Format-EbiExcelError -Stage 'paste into chart' -ErrorRecord $e1) 'error line: stage, message, type and HRESULT'

# 10/16 office run: an exception whose Message is empty
$e2 = New-Object System.Runtime.InteropServices.COMException ' ', -2146827284
$line2 = Format-EbiExcelError -Stage 'export' -ErrorRecord $e2
Assert-Equal 'export failed: (no message) [COMException 0x800A03EC]' $line2 'error line: an empty message still names the stage, the type and the HRESULT'

# the MethodInvocationException a COM method call is wrapped in is unwrapped
$inner = New-Object System.Runtime.InteropServices.COMException 'Exception from HRESULT: 0x800A03EC', -2146827284
$outer = New-Object System.Management.Automation.MethodInvocationException 'Exception calling "Export"', $inner
Assert-Equal 'export failed: Exception from HRESULT: 0x800A03EC [COMException 0x800A03EC]' (Format-EbiExcelError -Stage 'export' -ErrorRecord $outer) 'error line: the COM error inside a MethodInvocationException is the one reported'

# an ErrorRecord (what a catch block's $_ is)
$rec = New-Object System.Management.Automation.ErrorRecord ((New-Object System.IO.IOException 'disk full'), 'test', ([System.Management.Automation.ErrorCategory]::NotSpecified), $null)
$line4 = Format-EbiExcelError -Stage 'prepare' -ErrorRecord $rec
Assert-True ($line4 -like 'prepare failed: disk full `[IOException*') 'error line: an ErrorRecord is read through .Exception'

Assert-Equal 'copy shape failed: (no message)' (Format-EbiExcelError -Stage 'copy shape' -ErrorRecord $null) 'error line: no error object -> stage only'

exit (Complete-Tests)
