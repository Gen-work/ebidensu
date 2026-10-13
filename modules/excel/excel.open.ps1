# modules/excel/excel.open.ps1
# Open a workbook in the registered Excel instance and register it as a
# 'workbook' resource (P4-01). A workbook somebody else holds open comes
# back read-only from Excel: that is a failure here unless read-only was
# asked for, because every write after it would be lost silently.

. (Join-Path $PSScriptRoot '..\..\kernel\Native.ps1')   # Resolve-EbiWorkPath
. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.open'
  group      = 'excel'
  summary    = 'Open a workbook in the registered Excel instance and register it'
  tier       = 'core'
  effects    = 'ui'
  needs      = @('excel')
  provides   = @('workbook')
  releases   = @()
  idempotent = $true
  inputs     = @{
    app      = @{ type='session'; sessionKind='excelApp'; required=$true }
    path     = @{ type='path'; required=$true; desc='the .xlsx (relative: under the work dir)' }
    readOnly = @{ type='bool'; default=$false; desc='open read-only (reading a shared plan someone may have open)' }
  }
  outputs    = @{
    path     = @{ type='path' }
    name     = @{ type='string' }
    sheets   = @{ type='list'; desc='sheet names in tab order' }
    readOnly = @{ type='bool' }
  }
  failures   = @(
    @{ id = 'file_not_found'; transient = $true  }
    @{ id = 'read_only';      transient = $true  }
    @{ id = 'open_failed';    transient = $true  }
  )
  example    = @{ use = 'excel.open'; with = @{ app = 'xl'; path = '{{vars.evidenceDir}}/{{item.Excel_NAME}}.xlsx'; as = 'wb' } }
  notes      = 'read_only is transient: close the workbook in the other Excel (or ask whoever has it open) and retry; the failure names who holds it when Excel''s ~$ lock file next to the workbook says. Release with excel.close (once:groupEnd for a per-deliverable workbook, teardown for a run-wide one).'
}

function ExcelOpen-LockOwner {
    <#
      PURE. The bytes of Excel's owner file (~$<name>.xlsx, written next
      to a workbook while someone has it open) -> the user name in it, or
      ''. Layout: byte 0 = length, then the ANSI name; at byte 54 the same
      length and the name again in UTF-16LE (preferred: it survives any
      code page).
    #>
    param([byte[]]$Bytes)
    if ($null -eq $Bytes -or $Bytes.Length -lt 2) { return '' }
    $n = [int]$Bytes[0]
    try {
        if ($Bytes.Length -ge (55 + 2 * $n) -and [int]$Bytes[54] -eq $n -and $n -gt 0) {
            $u = [System.Text.Encoding]::Unicode.GetString($Bytes, 55, 2 * $n).Trim()
            if ($u -ne '') { return $u }
        }
        if ($n -gt 0 -and $Bytes.Length -ge (1 + $n)) {
            return ([System.Text.Encoding]::GetEncoding(932)).GetString($Bytes, 1, $n).Trim()
        }
    } catch { }
    return ''
}

function ExcelOpen-HeldBy {
    # The ~$ owner file of a workbook path -> name or '' (never throws).
    param([string]$Path)
    try {
        $lock = Join-Path (Split-Path -Path $Path -Parent) ('~$' + (Split-Path -Path $Path -Leaf))
        if (-not (Test-Path -LiteralPath $lock -PathType Leaf)) { return '' }
        return (ExcelOpen-LockOwner -Bytes ([System.IO.File]::ReadAllBytes($lock)))
    } catch { return '' }
}

function Invoke-Step {
    param($In, $Ctx)
    $p = Resolve-EbiWorkPath -PathValue ([string]$In['path']) -WorkDir ([string]$Ctx['WorkDir'])
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would open {0}{1}' -f $p, $(if ([bool]$In['readOnly']) { ' read-only' } else { '' }))); return @{ ok = $true; resource = $null; path = $p; name = (Split-Path -Path $p -Leaf); sheets = @(); readOnly = [bool]$In['readOnly'] } }
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { return @{ ok = $false; failure = 'file_not_found'; message = $p; path = $p; name = ''; sheets = @(); readOnly = [bool]$In['readOnly'] } }
    $wb = $null
    try { $wb = $In['app'].Workbooks.Open($p, 0, [bool]$In['readOnly']) }
    catch { return @{ ok = $false; failure = 'open_failed'; message = $_.Exception.Message; path = $p; name = ''; sheets = @(); readOnly = [bool]$In['readOnly'] } }
    $ro = $false
    try { $ro = [bool]$wb.ReadOnly } catch { }
    if ($ro -and -not [bool]$In['readOnly']) {
        try { $wb.Close($false) } catch { }
        Invoke-EbiComRelease $wb
        $who = ExcelOpen-HeldBy -Path $p
        $msg = if ($who -ne '') { ('{0} opened read-only -- {1} has it open (Excel''s ~$ lock file). Ask them to close it, then r' -f (Split-Path -Path $p -Leaf), $who) } else { ('{0} opened read-only -- is it open in another Excel window (no ~$ lock file says who)? Close it there, then r' -f (Split-Path -Path $p -Leaf)) }
        return @{ ok = $false; failure = 'read_only'; message = $msg; path = $p; name = ''; sheets = @(); readOnly = $true; heldBy = $who }
    }
    return @{ ok = $true; resource = $wb; path = $p; name = [string]$wb.Name; sheets = @(Get-EbiSheetNames -Workbook $wb); readOnly = $ro }
}
