# CLAUDE.md — VerifyTool Project Context

Read this file first when opening the project from an IDE or LLM session.

## Project purpose

**VerifyTool** automates GIFT→GFIX migration evidence collection at Honda Japan.
Operator: project user. Environment: Windows 10/11 + PowerShell 5.1 + Excel 2019.

The tool captures screenshots from HM / MQ / Jenkins, inserts them into evidence Excel
workbooks, draws red rectangles on the relevant cells, and tracks completion state in
a CSV mapping file.

## Repository

Remote: `gen-work/ebidensu`
Branch convention: `claude/<slug>`
Local clone: configure per environment (do not commit personal paths).

## File map

```
VerifyTool.ps1          main entry, menu, phase router, status display
VerifyConfig.psd1       project config (paths, scripts, PhaseOrder, Aliases, Mark.Boxes)
verify_session.json     last settings (WorkDir, Owner, WindowSize, CursorCell, CloneSourceDir)
verify_config.json      OPTIONAL per-work-folder JSON overlay (lives in WorkDir): deep-
                        merged over VerifyConfig.psd1 (JSON wins; CLI still wins).
                        Generate via -Phase InitConfig. Customizes owner / window /
                        Mark.Boxes / Mail / Reviewer / Df / ExpectedTime / etc.

ExcelHelpers.ps1        dot-source lib: Excel COM, bitmask, shape metadata helpers (no param())

  -- shared dot-source libraries (no param(); ASCII source; no BOM) --
MappingStore.ps1        single source of truth for mapping_<Owner>.csv: read/filter/
                        atomic-write. Import-Mapping, Export-MappingAtomic,
                        Ensure-MappingColumns, ConvertTo-TargetIdList, Test-TargetRow,
                        Get-PendingRows, Set-MappingBit. ALL scripts use this.
GfixLog.ps1             pure GFIX receive-log matcher (SS_CODE=Substring(4,1); newest
                        wins; whole-file lines). No Excel. Unit-tested.
EvidencePlan.ps1        pure correl-major Replace plan builders (Build-Gift/Gfix/Df
                        EvidencePlan) encoding the review order. No Excel. Unit-tested.
EvidenceExecutor.ps1    walks an EvidencePlan and performs the Excel inserts.
ProjectLabels.ps1       Japanese sheet/label names from [char] (keeps consumers ASCII /
                        codepage-agnostic). Get-AlignSendSheets / Get-AlignRecvSheets.
ProgressLog.ps1         append-only status\progress.jsonl events (UTF-8 no BOM).
ScreenRegion.ps1        pure screen-region clamp math. Unit-tested.
AlignCompare.ps1        pure sheet-compare + migration-type logic. Unit-tested.
ConfigOverlay.ps1       pure per-work-folder JSON overlay: deep-merge + JSON<->hashtable
                        + InitConfig snapshot/generator helpers. Unit-tested.
SendMetadata.ps1        pure SEND-side OCR-line parsing + send-vs-gift compare for
                        SendVsGift Stage 2 (word-box spacing rebuild, 0-byte rules:
                        used-CYLINDERS-0 / begin+end-on-one-image, row-label record
                        extraction, 80% prefix-similarity record compare,
                        Compare-SendGiftEvidence ok/ng/unknown verdict). Unit-tested.
OcrWindows.ps1          Windows built-in OCR (Windows.Media.Ocr WinRT) from PS 5.1;
                        same engine family as Snipping Tool text extraction. Safe to
                        dot-source anywhere (lazy init; Test-WinOcrAvailable).
EvidenceImageExport.ps1 Excel COM: export embedded sheet pictures to PNG via temp
                        ChartObject (skips verifyMark_* shapes; flattens Ctrl+G groups
                        to child pictures; optional Top range filter for one correl
                        section; clipboard clobbered).
SnapVerify.ps1          pure snap-phase NG detection library (no COM, no SendKeys).
                        ConvertFrom-HmPageText / Test-HmAbend (F1),
                        ConvertFrom-MqPageText / Test-MqRecord (F2),
                        ConvertFrom-JenkinsListText / Test-JenkinsFile (F3/F4),
                        Get-SnapPageKind (A3 sentinel), Resolve-SnapRunTime (2.2).
                        Unit-tested via Tests\Test-SnapVerify.ps1.
WorkbookResolver.ps1    dot-source helper: evidence/J4 workbook filename resolution
                        (prefix + Excel_NAME stem) plus reusable full-width
                        ASCII filename fallback (`FullWidthFilenameResolver`).
                        Unit-tested.

Clone.ps1               Phase Clone
Align.ps1               Phase Align/Precheck: compare work evidence vs J4 baseline
ReplaceEvidence.ps1     Phase ReplaceGift / ReplaceGfix / ReplaceDf (plan-driven)
Mark.ps1                Phase MarkGift / MarkGfix / MarkDf
ReviewEvidence.ps1      Phase ReviewGift / ReviewGfix / ReviewDf / ReviewEvidence
SendVsGift.ps1          Phase SendVsGift: GIFT file metadata vs SEND evidence review.
                        Rows grouped per workbook (opened once); cursor jumps to each
                        Correl_ID_S label in column A of the send sheet; Excel is
                        refocused after every console answer. Enter=1, n=2(NG), s, q.
                        Stage 2 -Ocr exports each correl's section pictures, OCRs and
                        auto-marks ok->1 / ng->2 / unknown->prompt (docs/SendVsGift.md).
OcrTool.ps1             standalone Windows-OCR CLI over OcrWindows/SendMetadata/
                        EvidenceImageExport: images, dirs, wildcards or -Workbook
                        picture export; -Json output; -ListLanguages. Reusable by
                        future features (has param(): call via &, never dot-source).
FillCheckSheet.ps1      Phase CheckSheet: append a row per Excel to the shared
                        review check sheet (Check Sheet_J4) via a temp-copy
                        preview, then commit only if the original is unchanged.
DeliverMail.ps1         Phase DeliverMail: one Outlook *draft* per Excel_NAME
                        (CreateItem+Display, never auto-sent); operator clicks
                        Send then Enter -> sets isDelivered. ASCII source;
                        subject/body/reviewer come from config (Mail/Reviewer).
Validate.ps1            Phase Validate (read-only diagnostic)
Watch-MappingProgress.ps1  read-only progress monitor (does NOT lock mapping)
Check-Encoding.ps1      read-only encoding policy checker + label self-test
Tests/                  Run-Tests.ps1 (parse-check all + units) + Test-*.ps1

JenkinsSnap.ps1         Phase GiftJenkins / GfixJenkins / GiftJenkinsNoFile
HmSnap.ps1              Phase GiftHmSnap / GfixHmSnap   (kept as-is)
MqSnap.ps1              Phase GiftMqSnap. MappingStore + ProgressLog + SnapVerify
                        F2 detection (page-text poll, page-kind sentinel, MQ
                        record verdict ok=1/ng=2, batch Expected_Time prompt).
ExcelSnap.ps1           Phase ExcelSnap                 (legacy, kept as-is)
Common.ps1              shared WinAPI/screenshot/SendKeys helpers (dot-sourceable)
Generate-HostOpenMapping.ps1  generates mapping CSV from wipGFIX一覧.xlsx.
                        -Add merges new selectors (JobNames / CorrelIdsM /
                        ExcelNames / WBS range) into an existing mapping,
                        keeping every existing row + its progress.

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

Only files with **no** `param()` block are ever dot-sourced: `ExcelHelpers.ps1`,
`MappingStore.ps1`, `GfixLog.ps1`, `EvidencePlan.ps1`, `EvidenceExecutor.ps1`,
`ProjectLabels.ps1`, `ProgressLog.ps1`, `ScreenRegion.ps1`, `AlignCompare.ps1`,
`ConfigOverlay.ps1`, `Common.ps1`, `WorkbookResolver.ps1`, `SendMetadata.ps1`,
`OcrWindows.ps1`, `EvidenceImageExport.ps1`, `SnapVerify.ps1`. All phase scripts have `param()`
and are called via `& $path @args`.

The critical pattern before any dot-source:
```powershell
$forceFlag = [bool]$Force.IsPresent   # capture switch BEFORE dot-sourcing
. $cfg.Scripts.ExcelHelpers            # dot-source (no param() = safe)
```

Never dot-source a script that has a `param()` block — it will overwrite the caller's
switch parameters with `$false`.

### Full-width filename fallback

`WorkbookResolver.ps1` exposes a reusable `FullWidthFilenameResolver` class and
wrapper functions for filename misses caused by full-width ASCII characters
(e.g. `０` instead of `0`). Use `Resolve-FullWidthFileName` after a normal exact
lookup fails when any file type needs the same tolerance:

```powershell
$path = Resolve-FullWidthFileName -Dir $dir -Name 'report0.txt' -Filter '*.txt' `
    -ItemKind 'file' -FullWidthFallback Prompt
```

Workbook callers should continue to use `Find-WorkbookByExcelName`; it preserves
exact and wildcard matching first, then delegates to the generic resolver for
full-width fallback with `Prompt` / `Accept` / `Reject` policy. Interactive tools
should keep the default `Prompt`; tests and non-interactive batch flows should
pass `Accept` or `Reject` explicitly.

### Encoding table

| File type | Encoding | BOM |
|-----------|----------|-----|
| .ps1 | UTF-8 | **no** (keep source ASCII; build Japanese via `[char]`) |
| .psd1 | UTF-8 | **yes** if it holds raw Japanese (Import-PowerShellDataFile can't use `[char]`); no if pure-ASCII |
| .json / .jsonl | UTF-8 | no |
| .csv (mapping) | UTF-8 | yes (BOM; Excel needs it for Japanese) |
| .md | UTF-8 | no |

New/rewritten `.ps1` must be ASCII-only: any Japanese used at runtime comes from
`ProjectLabels.ps1` ([char] code points). This makes the file work on **any**
Windows codepage. Raw Japanese in a no-BOM `.ps1` mojibakes on a JP-locale host
(this silently broke owner-matching in Generate-HostOpenMapping). Older files
still carry raw Japanese / a BOM; migrate them when touched. `Check-Encoding.ps1`
enforces this policy; `Apply-LlmPatch.ps1` preserves original BOM state.

### Bitmask fields

Three integer CSV columns track multi-mode completion:

| Field | bit 1 (1) | bit 2 (2) | bit 4 (4) | all done |
|-------|-----------|-----------|-----------|---------|
| isReplaced | GIFT replace | GFIX replace | DF replace | 7 |
| isMarked | GIFT mark | GFIX mark | DF mark | 7 |
| isReviewed | GIFT review | GFIX review | DF review | 7 |

Test: `($value -band $bit) -eq $bit` (use `Test-BitDone` / `Set-MappingBit` from
MappingStore). The GFIX-log yellow highlight is now part of MarkGfix (bit 2 of
`isMarked`); the old standalone `isGfixLogMarked` column was removed. Replace marks
a mode's bit only when ALL its required pieces inserted; which correl/step/file
failed is recorded in `status\progress.jsonl`, not in extra columns.

A free-text `ReviewComment` column (per Excel_NAME group) holds review notes
captured via the `-m "comment"` option at the Review prompt; list them with the
`Comments` phase.

`SendVsGift` is a plain value column (NOT a bitmask): `0`/empty = pending,
`1` = OK, `2` = NG (OCR auto-compare or the operator's `n` answer flagged a
disagreement). `2` still counts as pending: it is re-offered on the next
SendVsGift run and listed in the end-of-run NG summary.

`isDelivered` is a plain `0/1` flag (NOT a bitmask): set to `1` per Excel_NAME
when the operator confirms the DeliverMail draft was sent. `DeliverComment` is
its free-text note column (captured with `-m "comment"` at the DeliverMail
prompt). Both are defaulted by MappingStore; `isDelivered` is a `PhaseOrder`
field so it is auto-added on startup and shown in Status. The `CheckSheet` phase
writes only to the external review check sheet workbook — it does not touch the
mapping.

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

## Current state (last bump: 2026-06-17 v2.9.5)

v2.9.5 (SnapVerify M3 -- Jenkins instant NG detection): `JenkinsSnap.ps1` is wired
to F3 for the `GiftRecv` / `GfixRecv` modes. After the Ctrl+F search and screenshot
it polls the page text (A2), classifies the page (`Get-SnapPageKind`, sentinel A3),
archives the Ctrl+A text as `snap\<folder>\<correl>.txt` (A1), then
`ConvertFrom-JenkinsListText` + `Test-JenkinsFile` decides ok->`<field>_snap`=1 /
ng->2 (file missing from the list, or its timestamp outside the Expected_Time
window). The same polled text feeds the existing receive-file download (page read
once per row). NG=`2` stays pending (re-offered next run) and prints an end-of-run
NG summary. Pending uses a local `Test-JenkinsSnapDone` (done == '1'), not
`Get-PendingRows`, so NG rows aren't hidden; the batch run-time prompt
(`Resolve-SnapRunTime`) fills empty `Expected_Time` cells (plan 2.2). NoGfix (F4)
still runs pure screenshot (->1), pending M6. `SnapVerify.Enabled=$false` reverts
all Jenkins modes to pure screenshot. F3's pure helpers + unit tests shipped in M1
(this is wiring only).

v2.9.4 (SnapVerify M2 -- MQ instant NG detection): `MqSnap.ps1` rewritten from
the legacy bare-CSV version to the modern stack -- MappingStore (atomic writes),
ProgressLog events, and the pure `SnapVerify.ps1` detection library. It now runs
F2 (plan 2.4): after the search it polls the page text (A2), classifies the page
(`Get-SnapPageKind` sentinel A3), archives the Ctrl+A text as `<correl>.txt` (A1),
screenshots, then `ConvertFrom-MqPageText` + `Test-MqRecord` decides ok->`GIFT_MQ_snap`=1
/ ng->2 (No Data! / no matching row / RecvDate outside window / non-zero Rtncd|Rsncd).
NG=`2` still counts as pending (re-offered next run) and prints an end-of-run NG
summary. A one-time batch run-time prompt (`Resolve-SnapRunTime`) fills empty
`Expected_Time` cells on the pending rows (plan 2.2). Off-page kinds
(OuterFrame/Empty/Unknown) stop and ask `r=retry / s=skip / q=quit`.
`SnapVerify.Enabled=$false` reverts MqSnap to pure screenshot (legacy behavior).
Two new pure helpers in `SnapVerify.ps1` (`ConvertTo-ExpectedDateTime`,
`Set-EmptyRunTimeCells`) are unit-tested in `Tests/Test-SnapVerify.ps1`.
(v2.9.0-2.9.3 history: M1 detection library; NoGfix image overlap fix; ReviewGift/Gfix/Df
open the mode sheet; ReplaceGfix SS_CODE column -- see CHANGELOG.)

v2.8.1 (field fixes): PS 5.1 array-nesting bugs fixed across the OCR stack.
**Convention: shared lib functions return plain arrays -- never
`return ,@(...)` -- because callers wrap calls in `@(...)` and that
combination NESTS in PS 5.1** (one element = whole inner array; member
enumeration then yields Object[] and casts explode). Ctrl+G section
filtering now uses the top-level shape's Top (children can report
group-relative coordinates); export upscales 3x via the temp chart
(Chart.Shapes/Chart.Pictures stretch + GDI+ min-width fallback) and
prints a `[DIAG]` pixel size; `OcrTool.ps1 -Diag` sweeps every installed
recognizer language per image. OPEN TODO: Windows OCR still returns zero
lines on the correctly-sized evidence PNGs -- see docs/SendVsGift.md
"Troubleshooting: OCR reads nothing" for the next-session checklist.

v2.8: SendVsGift review rework (per-workbook grouping, column-A correl cursor,
Excel refocus after every answer, `n`=NG answer) + OCR auto-compare with the
operator's verdict rules (0-byte CYLINDERS/begin-end rules, max-row + first/last
record compare with 80% prefix similarity; ok->1, ng->2, unknown->prompt;
`SendVsGift.AutoMark`). Ctrl+G grouped pictures are now flattened on export.
New standalone `OcrTool.ps1` makes the OCR stack reusable. `o`/`-Ocr` toggle in
VerifyTool; standalone SendVsGift launch resolves WorkDir/Owner from session.


v2.7: per-work-folder JSON config overlay (`verify_config.json`, deep-merged over
VerifyConfig.psd1; generate with `-Phase InitConfig`) makes every case highly
customizable without editing the shared .psd1. Precedence: CLI > overlay > .psd1 > limited session fallback.
Almost all phases already read the merged `$Config`, so the overlay reaches them
all (owner, window, Mark.Boxes, Mail, Reviewer, Df, etc.). Centralized
`ExpectedTime` defaults + `-TimeFormat`; `Clone.SourceDir`; project-level
`Workbook.ExcelPrefix`; new pure unit-tested lib `ConfigOverlay.ps1`.

v2.6: incremental mapping `-Add` (grow the map day by day, keep progress).
Workbook filename prefix now lives in work-folder config; legacy mapping
`Excel_Prefix` is still honored only as a compatibility/per-row override.

Major refactor: shared MappingStore, plan-driven Replace, recovery + monitoring.
Pure (COM-free) libs are unit-tested via `Tests\Run-Tests.ps1`; COM/Edge phases
validated by static analysis only (no PowerShell/Excel in the cloud build env)
and need a Windows + Excel 2019 run to confirm end to end.

Phases: Mapping, InitConfig (new), ExcelSnap (legacy), GiftHmSnap, GiftMqSnap, GiftJenkins,
GiftJenkinsNoFile, GfixHmSnap, GfixJenkins, GfixLogDownload, DfSnap,
Clone, **Align (new)**, ReplaceGift/Gfix/Df, MarkGift/Gfix/Df,
ReviewGift/Gfix/Df, ReviewEvidence, **Comments (new, review-note list)**,
**CheckSheet (new, fill review check sheet)**, **DeliverMail (new, review-request mail)**,
Validate, RepairMapping, ProbeShapes, Crop, **WatchProgress (new)**.
The GFIX-log highlight is folded into MarkGfix; `MarkGfixLog.ps1` remains only as
a standalone re-highlight utility (reachable by name, no mapping column).

Replace is now plan-driven (EvidencePlan + EvidenceExecutor), correl-major per the
review standard. GFIX log is matched by `GfixLog.ps1` (whole file pasted) instead
of the old TODO placeholder. All mapping I/O goes through MappingStore (atomic
writes). Every phase appends events to `status\progress.jsonl`; watch them live
with `Watch-MappingProgress.ps1` (read-only, never locks the CSV).

To run the tests on Windows: `powershell -File Tests\Run-Tests.ps1` (parse-checks
every .ps1 + runs the unit tests). Encoding check: `powershell -File Check-Encoding.ps1`.

## Known issues / open points

- **Not yet run on Windows/Excel**: this refactor was authored in a Linux cloud
  env without PowerShell or Excel. Run `Tests\Run-Tests.ps1`, then smoke-test
  ReplaceGift/Gfix/Df and DfSnap on a copy before trusting them in production.
- **Align full branching** needs two domain facts: which FROM_sys/TO_sys literals
  mean "Host" (set `Align.HostSystemTypes` in VerifyConfig), and confirmation of
  the per-migration-type sheet sets in `AlignCompare.ps1`. Until then Align uses
  the send-result sheets (send[2], send[3]) and warns.
- **Align recv sheets are never synced** — recv sheets hold operator-captured evidence;
  only the host-team-managed send sheets are fetched from J4.
- **Align -Apply** syncs values + formats (Range.Copy) and is experimental.
- JenkinsSnap.ps1 matches the known-good repo logic (real Common.ps1 helpers); the
  earlier `Get-EdgeHwnd`/`Capture-Window` phantom-function risk is resolved.

## TODOs

- **SnapVerify M1 + M2 done** — M1: `SnapVerify.ps1` pure library +
  `Tests/Test-SnapVerify.ps1` unit tests + `SnapVerify` config section in
  `VerifyConfig.psd1`. M2: `MqSnap.ps1` migrated to MappingStore/ProgressLog and
  wired to F2 (page-text poll, page-kind sentinel, MQ verdict ok=1/ng=2, batch
  `Expected_Time` prompt); two new pure helpers (`ConvertTo-ExpectedDateTime`,
  `Set-EmptyRunTimeCells`) are unit-tested. **M3 done** -- `JenkinsSnap.ps1` wired
  to F3 (GiftRecv/GfixRecv NG=2 + summary, batch time prompt, sentinel,
  `Test-JenkinsSnapDone`); NoGfix stays pure-screenshot until M6. **M4 (HM wiring +
  HmSnap migration), M5 (pixel localisation), M6 (NoGfix annotation) remain.**
  When wiring M3/M4, copy MqSnap's `Test-MqSnapDone` pattern (done == exactly '1')
  so NG='2' rows stay pending -- `Get-PendingRows`/`Test-SnapDone` treat any
  non-'0' value as done and would hide NG rows. Design + open questions (only Q5,
  Rtncd/Rsncd semantics, is non-blocking) live in `docs/SnapVerify-Plan.md`.

- **Generate-HostOpenMapping `-Add` cannot filter by owner at the same time** —
  the daily flow adds new JOB_NAMEs incrementally with `-Add`, but owner
  filtering is not applied in that mode; fix so `-Add` + owner filter compose.

- **GfixLogDownload: auto-set GoAnywhere max rows to 100**
  Currently requires manual setup (default GoAnywhere list shows 20 rows — not enough for
  busy BIZ codes). Future: use SendKeys / UI automation to set the rows-per-page dropdown
  to 100 automatically after `Switch-ToEdge`, before the per-row search loop.

- **DfSnap: DfExePath not yet configured** — set `Df.ExePath` in `VerifyConfig.psd1` to
  the full path of df.exe so the prompt is skipped. Or just type it when prompted each run.

- **DfSnap region calibration** — default capture is `region` (x=120,y=280,w=1250,h=657
  for ~1980x1020). Tune `Df.RegionX/Y/Width/Height` and per-direction
  `Df.CropLeft/Top/Right/Bottom` (the window shadow is asymmetric). A pixel-color
  auto-detect of the window edge is a future option (no vision in a PS script).

- **GfixLogDownload max-rows** — still relies on manual "rows=100" setup.
  - **SS_CODE override DONE**: ReplaceGfix now reads an optional `SS_CODE` mapping
    column and threads it through the plan (`Build-GfixEvidencePlan -CorrelToSs`
    -> log op `SsCode` -> `Find-GfixLogForCorrel`). When the column is present and
    non-empty it wins; otherwise `GfixLog.ps1` infers SS from `Correl_ID_S`
    (5th char, or `J` for `J<biz>LxxS` jobs) exactly as before. Add an `SS_CODE`
    column to the mapping to take effect.

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
- `Owner` — mapping owner suffix (no personal default)
- `WindowWidth`, `WindowHeight`, `CropPx` — screenshot window size
- `CursorCell` — review cursor cell (default: A3)
- `CloneSourceDir` — external path for Clone phase
- `EvidenceDir` — evidence output folder (default: `<WorkDir>\evidence`)
