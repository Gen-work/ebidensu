# modules/excel/excel.copy_picture.ps1
# Copy the n-th picture of one sheet to a cell of another sheet (the
# deliverable repeats its first screenshot on a second sheet). Pictures
# are counted top-to-bottom, left-to-right. The clipboard is used (Excel
# has no other cross-sheet picture copy) and cleared afterwards.

. (Join-Path $PSScriptRoot '..\..\kernel\Excel.ps1')

$Manifest = @{
  id         = 'excel.copy_picture'
  group      = 'excel'
  summary    = 'Copy the n-th picture of a sheet to a cell of another sheet'
  tier       = 'core'
  effects    = 'write'
  needs      = @('excel')
  provides   = @()
  releases   = @()
  idempotent = $false
  inputs     = @{
    workbook  = @{ type='session'; sessionKind='workbook'; required=$true }
    fromSheet = @{ type='string'; required=$true }
    index     = @{ type='int';    default=1; desc='1-based, top-to-bottom' }
    toSheet   = @{ type='string'; required=$true }
    cell      = @{ type='string'; default='B3'; desc='top-left cell of the copy' }
    gapRows   = @{ type='int';    default=0; desc='added to nextRow' }
  }
  outputs    = @{
    name    = @{ type='string' }
    nextRow = @{ type='int'; desc='first row whose top is below the copy, plus gapRows' }
  }
  failures   = @(
    @{ id = 'sheet_not_found'; transient = $false }
    @{ id = 'not_found';       transient = $false }
    @{ id = 'copy_failed';     transient = $true  }
  )
  example    = @{ use = 'excel.copy_picture'; with = @{ workbook = 'wb'; fromSheet = 'sheetA'; toSheet = 'sheetB'; cell = 'B3' } }
  notes      = 'Clobbers the clipboard. Not idempotent: a rerun pastes a second copy.'
}

function Invoke-Step {
    param($In, $Ctx)
    if ($Ctx['DryRun']) { $Ctx.Log.Info(('would copy picture #{0} of {1} to {2}!{3}' -f $In['index'], $In['fromSheet'], $In['toSheet'], $In['cell'])); return @{ ok = $true; name = ''; nextRow = 0 } }
    $wb = $In['workbook']
    $src = Get-EbiSheet -Workbook $wb -Sheet $In['fromSheet']
    $dst = Get-EbiSheet -Workbook $wb -Sheet $In['toSheet']
    if ($null -eq $src -or $null -eq $dst) { return @{ ok = $false; failure = 'sheet_not_found'; message = ('sheet "' + $(if ($null -eq $src) { $In['fromSheet'] } else { $In['toSheet'] }) + '" not found'); name = ''; nextRow = 0 } }
    $pics = @(Get-EbiPictureShapes -Worksheet $src)
    $i = [int]$In['index']
    if ($i -lt 1 -or $i -gt $pics.Count) { return @{ ok = $false; failure = 'not_found'; message = ('{0} has {1} picture(s); #{2} asked' -f $In['fromSheet'], $pics.Count, $i); name = ''; nextRow = 0 } }
    try {
        $shape = $pics[$i - 1]['shape']
        $before = [int]$dst.Shapes.Count
        [void]$shape.Copy()
        Start-Sleep -Milliseconds 300
        $target = $dst.Range([string]$In['cell'])
        try { [void]$dst.Activate() } catch { }   # Paste is unreliable into a sheet that is not active
        [void]$dst.Paste($target)
        Start-Sleep -Milliseconds 300
        if ([int]$dst.Shapes.Count -le $before) { return @{ ok = $false; failure = 'copy_failed'; message = 'the paste added no shape'; name = ''; nextRow = 0 } }
        $new = $dst.Shapes.Item([int]$dst.Shapes.Count)
        $new.Left = [double]$target.Left
        $new.Top = [double]$target.Top
        try { $wb.Application.CutCopyMode = $false } catch { }
        $next = Get-EbiRowAtOrBelow -Worksheet $dst -TopPoints ([double]$new.Top + [double]$new.Height) -StartRow ([int]$target.Row)
        return @{ ok = $true; name = [string]$new.Name; nextRow = ($next + [int]$In['gapRows']) }
    } catch { return @{ ok = $false; failure = 'copy_failed'; message = $_.Exception.Message; name = ''; nextRow = 0 } }
}
