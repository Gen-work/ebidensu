# CLAUDE.md — VerifyTool Project Context

Read this file first when opening the project from an IDE or LLM session.

## Project purpose

**VerifyTool** automates GIFT→GFIX migration evidence collection at Honda Japan.
Operator: 厳 (Misaki). Environment: Windows 10/11 + PowerShell 5.1 + Excel 2019.

The tool captures screenshots from HM / MQ / Jenkins, inserts them into evidence Excel
workbooks, draws red rectangles on the relevant cells, and tracks completion state in
a CSV mapping file.

## Repository

Remote: `gen-work/ebidensu`
Branch convention: `claude/<slug>`
Local clone (office): `\\fs-f3170-1\...\厳\Work\0514_JRV-IDS,IGP2\`

## File map

```
VerifyTool.ps1          main entry, menu, phase router, status display
VerifyConfig.psd1       project config (paths, scripts, PhaseOrder, Aliases, Mark.Boxes)
verify_session.json     last settings (WorkDir, Owner, WindowSize, CursorCell, CloneSourceDir)

ExcelHelpers.ps1        dot-source lib: Excel COM, bitmask, shape metadata helpers (no param())
Clone.ps1               Phase Clone
ReplaceEvidence.ps1     Phase ReplaceGift / ReplaceGfix / ReplaceDf
Mark.ps1                Phase MarkGift / MarkGfix / MarkDf
ReviewEvidence.ps1      Phase ReviewGift / ReviewGfix / ReviewDf / ReviewEvidence
Validate.ps1            Phase Validate (read-only diagnostic)

JenkinsSnap.ps1         Phase GiftJenkins / GfixJenkins / GiftJenkinsNoFile
HmSnap.ps1              Phase GiftHmSnap / GfixHmSnap   (kept as-is)
MqSnap.ps1              Phase GiftMqSnap                (kept as-is)
ExcelSnap.ps1           Phase ExcelSnap                 (legacy, kept as-is)
Common.ps1              shared WinAPI/screenshot/SendKeys helpers (dot-sourceable)
Generate-HostOpenMapping.ps1  generates mapping CSV from wipGFIX一覧.xlsx

Calibrate-HmGeometry.ps1  4-click WinForms calibration for HM geometry offsets
Find-Abend.ps1          template-match for HM status cell
Find-ActiveHighlightRow.ps1  detects Edge Ctrl+F active-match (orange) row
Locate-ByImage.ps1      C#-compiled LockBits template matcher
Mark.ps1                draws red rectangles on evidence Excel shapes
Pack-LlmContext.ps1     packs project context to clipboard for LLM ingestion
Apply-LlmPatch.ps1      applies XML / git-unified-diff patches from clipboard
Export-DailyPatch.ps1   extracts today's git diff to clipboard
Parse-GiftMq.ps1        parses GIFT/MQ transfer status page text
Parse-JenkinsList.ps1   parses Jenkins file list page text
Probe-Shapes.ps1        lists all shapes in an evidence workbook (calibration aid)
Read-ClipboardJson.ps1  polls clipboard for JSON from bookmarklet
Read-PageText.ps1       captures visible text from foreground Edge page via clipboard
Resolve-ExpectedTime.ps1  interactive Expected_Time column helper
ReviewEvidence.ps1      manual review driver
Sample-HighlightColor.ps1  samples a single pixel RGB for highlight calibration

CLAUDE.md               this file
README.md               user-facing documentation
CHANGELOG.md            iteration log
```

## Conventions

### Dot-source safety rule

Only `ExcelHelpers.ps1` (which has **no** `param()` block) is ever dot-sourced.
All other scripts have `param()` and are called via `& $path @args`.

The critical pattern before any dot-source:
```powershell
$forceFlag = [bool]$Force.IsPresent   # capture switch BEFORE dot-sourcing
. $cfg.Scripts.ExcelHelpers            # dot-source (no param() = safe)
```

Never dot-source a script that has a `param()` block — it will overwrite the caller's
switch parameters with `$false`.

### Encoding table

| File type | Encoding | BOM |
|-----------|----------|-----|
| .ps1 | UTF-8 | yes (BOM) recommended for PS 5.1 |
| .psd1 | UTF-8 | yes |
| .json | UTF-8 | no |
| .csv (mapping) | UTF-8 BOM | yes (Excel needs BOM) |
| .md | UTF-8 | no |

`Apply-LlmPatch.ps1` preserves the original BOM state when writing patched files.

### Bitmask fields

Three integer CSV columns track multi-mode completion:

| Field | bit 1 (1) | bit 2 (2) | bit 4 (4) | all done |
|-------|-----------|-----------|-----------|---------|
| isReplaced | GIFT replace | GFIX replace | DF replace | 7 |
| isMarked | GIFT mark | GFIX mark | DF mark | 7 |
| isReviewed | GIFT review | GFIX review | DF review | 7 |

Test: `($value -band $bit) -eq $bit`

`PhaseOrder` in `VerifyConfig.psd1` has a `BitValue` key for each bitmask phase.

### Excel COM rules

- Always `$xl.Visible = $true` first, then `$xl.DisplayAlerts = $false`.
- Release COM objects in reverse order: `[Runtime.InteropServices.Marshal]::ReleaseComObject($ws)` etc.
- Use `$xl.Quit()` only when you opened a fresh Excel instance.
- `ExcelHelpers.ps1` functions: `Open-ExcelWorkbook`, `Save-ExcelWorkbook`, `Close-ExcelWorkbook`,
  `Set-MappingBit`, `Get-MappingValue`, `Find-ShapeByAltText`, `Add-VerifyMarkRect`.

### Shape metadata

Red rectangle mark shapes are named `verifyMark_<folder>_<idx>` and have AltText
`verifyMark|<folder>|<idx>|<correl>`.

`Probe-Shapes.ps1` reads these to help calibrate `Mark.Boxes` offsets in `VerifyConfig.psd1`.

### Switch flag pattern

```powershell
param(
    [switch]$Force,
    ...
)
$forceFlag = [bool]$Force.IsPresent
. $cfg.Scripts.ExcelHelpers
# use $forceFlag from here on, NOT $Force
```

## Current state (last bump: 2026-05-28 v2.4)

All core phases implemented and tested:
- Mapping, ExcelSnap (legacy), GiftHmSnap, GiftMqSnap, GiftJenkins, GiftJenkinsNoFile
- GfixHmSnap, GfixJenkins
- GfixLogDownload — downloads GFIX receive logs from GoAnywhere; moves .log to work\log\
- DfSnap — DF evidence screenshot (df.exe compare + fullscreen capture)
- MarkGfixLog — yellow-highlights the Command: line in GFIX log evidence Excel
- Clone, ReplaceGift, ReplaceGfix, ReplaceDf
- MarkGift, MarkGfix, MarkDf
- ReviewGift, ReviewGfix, ReviewDf, ReviewEvidence
- Validate, RepairMapping, ProbeShapes, Crop

`ReplaceGfix` writes `<<TODO: GFIX 受信 log>>` placeholder if log not yet downloaded.
Once GfixLogDownload runs, `Get-GfixLogLines` in `ReplaceEvidence.ps1` greps `work\log\`.

## Known issues / open points

- **JenkinsSnap.ps1 local rewrite risk**: the local copy visible in the 2026-05-28 daily
  diff references helper functions (`Get-EdgeHwnd`, `Activate-Window`, `Resize-Window`,
  `Capture-Window`) that do NOT exist in Common.ps1. If Jenkins phases crash, restore
  JenkinsSnap.ps1 from repo with `git checkout origin/claude/practical-maxwell-XvKaA -- JenkinsSnap.ps1`,
  then re-run `Fix-Encoding.ps1`.

## TODOs

- **GfixLogDownload: auto-set GoAnywhere max rows to 100**
  Currently requires manual setup (default GoAnywhere list shows 20 rows — not enough for
  busy BIZ codes). Future: use SendKeys / UI automation to set the rows-per-page dropdown
  to 100 automatically after `Switch-ToEdge`, before the per-row search loop.

- **DfSnap: DfExePath not yet configured** — set `Df.ExePath` in `VerifyConfig.psd1` to
  the full path of df.exe so the prompt is skipped. Or just type it when prompted each run.

## Cross-environment workflow

```
Office PC  →  Pack-LlmContext.ps1  →  clipboard
clipboard  →  paste to Claude/Cursor
Claude     →  XML patch or git diff
patch      →  Apply-LlmPatch.ps1   →  local files
today diff →  Export-DailyPatch.ps1 → clipboard → git push from home
```

`Apply-LlmPatch.ps1` accepts:
1. XML patch: `<patch><file name="..."><search>...</search><replace>...</replace></file></patch>`
2. git unified diff: standard `--- a/file` / `+++ b/file` / `@@ ... @@` format
3. Markdown fences around either format are stripped automatically.

## Session config (verify_session.json)

Remembered between runs:
- `WorkDir` — last work folder path
- `Owner` — mapping owner suffix (default: 厳)
- `WindowWidth`, `WindowHeight`, `CropPx` — screenshot window size
- `CursorCell` — review cursor cell (default: A3)
- `CloneSourceDir` — external path for Clone phase
- `EvidenceDir` — evidence output folder (default: `<WorkDir>\evidence`)
