#Requires -Version 5.1
Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$here = Split-Path $MyInvocation.MyCommand.Path
. (Join-Path $here '_TestCommon.ps1')
. (Join-Path (Split-Path $here -Parent) 'OwnerFilter.ps1')

Reset-Tests 'OwnerFilter'

# Arrows are built from [char] so this test stays ASCII-source (same rule as
# OwnerFilter.ps1 itself); a raw arrow literal would mojibake on CP932.
# Build the cells by concatenation to avoid any string-escape ambiguity.
$L = [char]0x2190   # owner<-other
$R = [char]0x2192   # other->owner
$ownerRecvArrow   = 'BBB' + $R + 'AAA'   # other->owner : owner is receiver (owned)
$ownerLeftArrow   = 'AAA' + $L + 'BBB'   # owner<-other : owner is receiver (owned)
$reverseLeftArrow = 'BBB' + $L + 'AAA'   # other<-owner : owner is sender   (NOT owned)

# ---- Test-OwnerMatch ----
Assert-True  (Test-OwnerMatch 'AAA' 'AAA')                    'exact owner matches'
Assert-True  (Test-OwnerMatch $ownerLeftArrow 'AAA')         'owner<-other matches (owner is receiver)'
Assert-True  (Test-OwnerMatch $ownerRecvArrow 'AAA')         'other->owner matches (owner is receiver)'
Assert-True  (-not (Test-OwnerMatch $reverseLeftArrow 'AAA')) 'other<-owner does NOT match (reverse dir)'
Assert-True  (-not (Test-OwnerMatch 'BBB' 'AAA'))            'different owner does not match'
Assert-True  (-not (Test-OwnerMatch '' 'AAA'))               'empty cell does not match'
Assert-True  (-not (Test-OwnerMatch 'AAA' ''))               'empty owner input does not match'
Assert-True  (Test-OwnerMatch '  AAA  ' 'AAA')               'whitespace is trimmed'

# ---- Select-JobsByOwner ----
$map = @{
    'JOB_MINE'    = 'AAA'
    'JOB_RECV'    = $ownerRecvArrow    # other->owner : owned
    'JOB_OTHER'   = 'BBB'              # someone else
    'JOB_REVERSE' = $ownerLeftArrow    # AAA<-BBB : owned (owner receives)
}

# Mix of: owned, owned-via-arrow, other-owner, and a temp job absent from WBS.
$sel = Select-JobsByOwner -Jobs @('JOB_MINE','JOB_RECV','JOB_OTHER','JOB_TEMP') -JobOwnerMap $map -OwnerInput 'AAA'
Assert-Equal 3 $sel.Kept.Count     'kept = 2 owned + 1 temp (absent from WBS)'
Assert-True  ($sel.Kept -contains 'JOB_MINE') 'JOB_MINE kept'
Assert-True  ($sel.Kept -contains 'JOB_RECV') 'JOB_RECV (other->owner) kept'
Assert-True  ($sel.Kept -contains 'JOB_TEMP') 'JOB_TEMP (absent) kept'
Assert-True  (-not ($sel.Kept -contains 'JOB_OTHER')) 'JOB_OTHER excluded from kept'
Assert-Equal 1 $sel.Excluded.Count 'one job excluded by owner filter'
Assert-Equal 'JOB_OTHER' $sel.Excluded[0].Job 'excluded job is JOB_OTHER'
Assert-Equal 'BBB' $sel.Excluded[0].OwnerCell 'excluded carries its WBS owner cell'
Assert-Equal 1 $sel.Unknown.Count  'one job unknown (absent from WBS)'
Assert-Equal 'JOB_TEMP' $sel.Unknown[0] 'unknown job is JOB_TEMP'

# Reverse-direction owner cell is NOT owned -> excluded.
$sel2 = Select-JobsByOwner -Jobs @('JOB_X') -JobOwnerMap @{ 'JOB_X' = $reverseLeftArrow } -OwnerInput 'AAA'
Assert-Equal 0 $sel2.Kept.Count     'reverse-direction job not kept'
Assert-Equal 1 $sel2.Excluded.Count 'reverse-direction job excluded'

# De-dup + whitespace + empty/null entries collapse.
$sel3 = Select-JobsByOwner -Jobs @('JOB_MINE',' JOB_MINE ','', $null) -JobOwnerMap $map -OwnerInput 'AAA'
Assert-Equal 1 $sel3.Kept.Count 'duplicate/blank/null candidates collapsed to one'

# Null map -> everything kept (no judging possible) and marked unknown.
$sel4 = Select-JobsByOwner -Jobs @('JOB_OTHER') -JobOwnerMap $null -OwnerInput 'AAA'
Assert-Equal 1 $sel4.Kept.Count    'null map keeps all candidates'
Assert-Equal 1 $sel4.Unknown.Count 'null map marks candidate unknown'

exit (Complete-Tests)
