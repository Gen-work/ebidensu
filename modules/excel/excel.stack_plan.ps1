# modules/excel/excel.stack_plan.ps1
# Turn capture sets (screen.launch_capture shots: a first capture, and a
# last one when the data runs past one screen) into the picture list
# excel.insert_pictures stacks: first, then an optional separator picture
# (the "cut here" wave between the top and the end of a long listing),
# then last; the marking boxes go on the picture that shows the END of the
# data. Pure: no Excel is touched (kernel/Layout.ps1 New-EbiStackPlan).

. (Join-Path $PSScriptRoot '..\..\kernel\Layout.ps1')

$Manifest = @{
  id         = 'excel.stack_plan'
  group      = 'excel'
  summary    = 'Plan a picture stack from capture sets (first, separator, last; boxes on the end)'
  tier       = 'core'
  effects    = 'pure'
  needs      = @()
  provides   = @()
  releases   = @()
  idempotent = $true
  inputs     = @{
    sets            = @{ type='list';   required=$true; desc='@{ first; last } per capture set (last empty = one screen)' }
    separator       = @{ type='path';   default=''; desc='picture placed between first and last' }
    separatorColumn = @{ type='string'; default=''; desc='column of the separator (default: column)' }
    separatorRows   = @{ type='int';    default=4; desc='rows the separator takes' }
    markRects       = @{ type='list';   default=@(); desc='excel.insert_pictures rect specs for the end picture' }
    gapRows         = @{ type='int';    default=1; desc='blank rows between capture sets' }
    column          = @{ type='string'; default='B' }
  }
  outputs    = @{
    pictures = @{ type='list'; desc='excel.insert_pictures pictures entries' }
    total    = @{ type='int' }
  }
  failures   = @(
    @{ id = 'input_invalid'; transient = $false }
  )
  example    = @{ use = 'excel.stack_plan'; with = @{ sets = '{{steps.df.out.shots}}'; separator = '{{run.toolDir}}/profiles/demo/assets/wave.png'; separatorColumn = 'Z' } }
}

function Invoke-Step {
    param($In, $Ctx)
    $sets = @(@($In['sets']) | Where-Object { $_ -is [System.Collections.IDictionary] -and -not [string]::IsNullOrWhiteSpace([string]$_['first']) })
    $plan = @(New-EbiStackPlan -Sets $sets -Separator ([string]$In['separator']) -SeparatorColumn ([string]$In['separatorColumn']) -SeparatorRows ([int]$In['separatorRows']) -MarkRects @($In['markRects']) -GapRows ([int]$In['gapRows']) -Column ([string]$In['column']))
    return @{ ok = $true; pictures = $plan; total = $plan.Count }
}
