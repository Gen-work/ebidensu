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
GfixJobList.ps1         pure parser for the GoAnywhere completed-jobs LIST page text
                        (Ctrl+A/Ctrl+C capture): ConvertFrom-GfixJobListText (tab-
                        delimited rows keyed by JobNo, data rows identified by a
                        numeric JobNo regex -- no Japanese literals needed) +
                        Get-GfixJobListRowsForIf (filter by normalized IF_NO,
                        receive-side only by default). No COM. Unit-tested. Lets
                        GfixLogDownload fetch every job matching a needed IF_NO
                        (job numbers are unique; IF_NO/project-name text is not).
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
SnapVerify.ps1          pure snap-phase NG detection + localisation library
                        (no COM, no SendKeys). ASCII source (Japanese via [char]).
                        ConvertFrom-HmPageText / Test-HmAbend (F1),
                        ConvertFrom-MqPageText / Test-MqRecord (F2),
                        ConvertFrom-JenkinsListText / Test-JenkinsFile (F3/F4),
                        Get-SnapPageKind (A3 sentinel), Resolve-SnapRunTime (2.2),
                        and M5/F5 pixel localisation: Get-MatchedRowIndex /
                        Get-RowPixelRect / Get-JenkinsHighlightRect /
                        New-SnapLocRect / Save-SnapLocSidecar.
                        Unit-tested via Tests\Test-SnapVerify.ps1.
SnapLocalize.ps1        M5/F5 wiring glue (NOT pure: System.Drawing + the
                        Find-ActiveHighlightRow scan). Write-SnapLocalize turns a
                        verdict into a snap\<folder>\<correl>.loc.json sidecar via
                        the pure SnapVerify geometry; swallows all errors (never
                        blocks snapping). Dot-sourced by Hm/Mq/JenkinsSnap when
                        SnapVerify.Localize.Enabled. No param() = safe to dot-source.
WorkbookResolver.ps1    dot-source helper: evidence/J4 workbook filename resolution
                        (prefix + Excel_NAME stem) plus reusable full-width
                        ASCII filename fallback (`FullWidthFilenameResolver`).
                        Unit-tested.
OwnerFilter.ps1         pure WBS owner-cell matching (Test-OwnerMatch: exact /
                        owner<-other / other->owner; reverse dir = not owned)
                        + Select-JobsByOwner (filter explicit -Add JOB_NAMEs by
                        WBS owner; jobs absent from WBS kept as temp). No Excel.
                        Unit-tested (Tests\Test-OwnerFilter.ps1).

Clone.ps1               Phase Clone
Align.ps1               Phase Align/Precheck: compare work evidence vs J4 baseline
ReplaceEvidence.ps1     Phase ReplaceGift / ReplaceGfix / ReplaceDf (plan-driven)
Mark.ps1                Phase MarkGift / MarkGfix / MarkDf. Each Mark.Boxes
                        entry may add a 'Template' key to try image-recognition
                        placement (Locate-ByImage.ps1 LockBits match against
                        the source snap PNG) before falling back to the fixed
                        OffsetX/OffsetY box -- see mark_templates/README.txt.
                        -Mode Gfix also highlights the GFIX log Command: row,
                        auto-sized to the row's actual text width (GfixLog.
                        AutoHighlightWidth, capped at HighlightColEnd).
mark_templates/         reference images for Mark.ps1's optional image-match
                        box placement (see mark_templates/README.txt). Ships
                        empty; populated per project on an office PC.
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
DeliverFiles.ps1        Phase DeliverFiles: replaces the 3 delivery-scope
                        sheets (GIFT/GFIX recv result + GIFT-vs-GFIX data
                        compare -- Align.ps1's Get-AlignRecvSheets set) in
                        the corresponding J4 workbook with the matching
                        work sheets, in place (other J4 sheets untouched);
                        first delivery for an Excel_NAME copies the whole
                        file instead. Also copies DATA\GFIX/GIFT. Never
                        deletes source files. Sets isFilesDelivered.
BackupJ4.ps1            Phase BackupJ4 ("bk"): read-only against J4 --
                        copies each targeted Excel_NAME's current J4
                        workbook into a local, timestamped backup folder
                        (default <WorkDir>\bk). Run before DeliverFiles to
                        keep a local rollback point (DeliverFiles now edits
                        J4 workbooks in place instead of always overwriting
                        the whole file).
Validate.ps1            Phase Validate (read-only diagnostic)
Watch-MappingProgress.ps1  read-only progress monitor (does NOT lock mapping)
Check-Encoding.ps1      read-only encoding policy checker + label self-test
Tests/                  Run-Tests.ps1 (parse-check all + units) + Test-*.ps1

JenkinsSnap.ps1         Phase GiftJenkins / GfixJenkins / GiftJenkinsNoFile
HmSnap.ps1              Phase GiftHmSnap / GfixHmSnap. MappingStore + ProgressLog
                        + SnapVerify F1 detection (page-text poll, page-kind
                        sentinel, HM abend verdict ok=1/ng=2/ask, newest-wins
                        within the time window, batch Expected_Time prompt).
                        Per-TO_code appl grouping (one HM page per appl).
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
Locate-ByImage.ps1      C#-compiled LockBits template matcher; called by
                        Mark.ps1 (see the Mark.ps1 entry above) for optional
                        image-recognition box placement
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
`MappingStore.ps1`, `GfixLog.ps1`, `GfixJobList.ps1`, `EvidencePlan.ps1`,
`EvidenceExecutor.ps1`, `ProjectLabels.ps1`, `ProgressLog.ps1`, `ScreenRegion.ps1`,
`AlignCompare.ps1`, `ConfigOverlay.ps1`, `Common.ps1`, `WorkbookResolver.ps1`,
`SendMetadata.ps1`, `OcrWindows.ps1`, `EvidenceImageExport.ps1`, `SnapVerify.ps1`,
`SnapLocalize.ps1`, `OwnerFilter.ps1`, `Find-ActiveHighlightRow.ps1`. All phase
scripts have `param()` and are called via `& $path @args`.

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

### Config overlay groups must track VerifyConfig.psd1

`ConfigOverlay.ps1`'s `Get-ConfigOverlayGroups` is a second, hand-maintained
index of `VerifyConfig.psd1`'s top-level sections (used by the `-Phase
InitConfig -Interactive` grouped field walker and by
`Get-ConfigOverlayReadmeText`). `New-ConfigOverlaySnapshot`/
`Update-ConfigOverlayData` read `VerifyConfig.psd1` generically (every
top-level key is captured and schema-repaired automatically), so a new
top-level section written into `.psd1` is silently correct at the JSON/repair
layer but invisible in the grouped editor and README until someone also adds
it to a NAMED group in `Get-ConfigOverlayGroups` -- it stays reachable only
via the catch-all `all` group, unlabeled. This actually happened:
`SnapVerify` (added v2.9.4, a major feature spanning six changelog entries)
had no named group until this was caught and fixed. **Whenever a phase gains
a new top-level `VerifyConfig.psd1` config section, add it to the most
relevant group in `Get-ConfigOverlayGroups` (and mention it under "Common
fields" in `Get-ConfigOverlayReadmeText`) in the same change.**
`Tests\Test-ConfigOverlay.ps1` has a schema-drift guard that fails the build
when a snapshot field is reachable only via `all` -- run
`Tests\Run-Tests.ps1` after any `VerifyConfig.psd1` structural change. That
same test file also repairs a reduced copy of the REAL `VerifyConfig.psd1`
defaults (not just hand-built fixtures) to confirm `-Phase InitConfig`
repair never drops an operator value and never throws against the actual
production config shape.

## Current state (last bump: 2026-07-06 v2.9.30)

v2.9.30 (Mark image-match: anchor-only sizing + snap-time template-hit
sidecar; DfSnap default to window capture; FillCheckSheet write verification):
**Fixed** -- (1) a `Mark.Boxes` entry with a `Template` match used to size
the drawn red box from the TEMPLATE CROP'S OWN pixel dimensions (scaled +
padded) whenever a match was found, ignoring the box's configured
`Width`/`Height`. This was fine for a box whose whole point was the crop
(`GIFT_noGfixfile`'s `StampImage`), but broke as soon as `Template` was added
to an existing fixed-size box (`GIFT_Jenkins`/`GFIX_Jenkins`, previously
`OffsetX/OffsetY/Width=288.8/Height=18.8`): the drawn box shrank to match the
anchor crop's own size instead of staying `288.8x18.8`. `Find-MarkBoxByImage`
now checks for `Width`/`Height` on the box FIRST -- when either is present,
Template only supplies the anchor (top-left corner, still shiftable via
`PadX`/`PadY`) and the configured `Width`/`Height` are used as-is; a box with
neither keeps the old crop-derived-size behavior. This lets an anchor crop be
a small, stable, unique landmark (e.g. a fixed label near the real target)
without forcing the box to that crop's own dimensions -- expand out from the
matched top-left corner with `Width`/`Height`/`PadX`/`PadY` instead of only
`PadX`/`PadY` growing the crop's own size. (2) `FillCheckSheet.ps1`'s cell
writes (date + 5 other columns) were each a bare `try { ... } catch {}`
around `.Value2 = ...` with no readback -- any exception, or any silent no-op
(e.g. a write that doesn't error but doesn't stick), was swallowed with zero
diagnostic, and the run still printed `[OK] wrote N row(s)` even when column
B (date) came back blank. New `Set-CellChecked` writes then reads back every
cell and compares; a mismatch or exception is now logged (`[WARN] <label>
write failed/did not verify (row N)`) and the row is flagged not-OK, so the
final `[OK] wrote N row(s)` line and the `status\progress.jsonl` event now
reflect what the workbook actually contains instead of what the code
attempted.
**Added** -- (1) snap-time template-hit sidecar: `JenkinsSnap.ps1` (GiftJenkins
/ GfixJenkins / GiftJenkinsNoFile) accepts `-MarkBoxes` (`Config.Mark.Boxes[
<folder>]`, threaded from `VerifyTool.ps1`) and, right after each screenshot
is saved, runs the same `Locate-ByImage.ps1` match for every box carrying a
`Template` key -- against the page as JUST captured, while it's known-good --
and caches any hits to `snap\<folder>\<correl>.tplhit.json`
(`SnapLocalize.ps1`'s new `Write-MarkTemplateHits`, same
never-blocks-the-caller contract as the existing `Write-SnapLocalize`).
`Mark.ps1`'s `Find-MarkBoxByImage` now reads this sidecar first
(`Get-MarkTemplateHitFromSidecar`) instead of re-scanning the archived PNG,
falling back to a live match whenever the sidecar is missing, has no entry
for that box, names a different Template (stale after a config edit), or was
recorded against a different-sized PNG -- so Mark never depends on the
sidecar existing. Zero new config fields: this reuses the same `Mark.Boxes[
<folder>].Template` key that already opts a box into image-match, so a
folder with no Template-carrying boxes is entirely unaffected. Console lines
now show which anchor source fired: `[MARK-IMG] ... (sidecar)` /
`(live)` / `[STAMP-IMG] ... (sidecar)` / `(live)`. (2) `Df.CaptureMode`
default changed from `'region'` to `'window'`: DF snap now auto-fits
whatever size the operator actually sized the df.exe window to, instead of
assuming a fixed `1250x657` rectangle. The existing invalid-handle/invalid-
rect fallback to `region` is unchanged and is now the safety net rather than
the default; set `Df.CaptureMode = 'region'` to restore the old behavior on
a machine where `window` proves unreliable.
**Notes** -- Static-checked only (no Windows/Excel/Edge in this dev
environment): confirm the anchor-only sizing on `GIFT_Jenkins`/`GFIX_Jenkins`
once real `Template` crops are added, confirm the `.tplhit.json` sidecar is
written at snap time and actually read (not silently falling back to live
every time) on an office PC, confirm `window` capture mode against a real
df.exe window, and confirm `FillCheckSheet`'s verified writes on the real
check-sheet workbook.

v2.9.29 (FillCheckSheet: on-disk prefix fallback + CheckSheetPath remembered):
**Fixed** -- (1) check-sheet column F (review target) could list a filename
that didn't match the real workbook: `Resolve-ExcelPrefix` (mapping row's
legacy `Excel_Prefix` column, else `Workbook.ExcelPrefix`) had no way to
tell "deliberately no prefix" from "nothing configured for this row" -- a
row lacking both (older-vintage mapping row, or a run where
`Workbook.ExcelPrefix` isn't set) silently produced a bare, unprefixed name
even while sibling rows in the same run got the full prefix from their own
legacy column value. `FillCheckSheet.ps1` gained an `-EvidenceDir` param
(default `<WorkDir>\evidence`, same convention as DeliverFiles/DeliverMail)
and, whenever the resolved prefix is blank, looks up the real evidence file
on disk (`Find-WorkbookByExcelName`) and recovers its actual prefix via
`Get-PrefixFromFilename` (existing helper, previously only used by
DeliverFiles' bare-name fallback, read here in the opposite direction) --
so the check sheet always lists what's actually on disk. (2) `CheckSheetPath`
prompted on every run even after answering it: the prompt lived inside
`FillCheckSheet.ps1`, so the operator's answer was a local variable that
vanished when the script returned, never reaching `verify_session.json`
(config *was* being read correctly -- `CheckSheet.Path` was just genuinely
unset). Moved the prompt into `VerifyTool.ps1`'s `CheckSheet` dispatch,
mirroring the existing `DfExePath` first-run-prompt-then-remember pattern:
State (CLI/session/menu `k`) wins, then config `CheckSheet.Path`, and only
if both are empty does it prompt once and immediately persist the answer.
**Added** -- the check-sheet row preview now prints column B's date
(`yyyy/MM/dd`) next to each planned row so the operator can confirm it
without opening the workbook (behavior itself is unchanged: a real date
value, format mirrored from the row above, equivalent to Ctrl+;). Static-
checked only (no Windows/Excel in this dev environment) -- confirm the
on-disk prefix fallback and the CheckSheetPath remember-once flow on an
office PC.

v2.9.28 (InitConfig: fix GetNewClosure() losing ConfigOverlay.ps1 functions
+ silent non-terminating failure):
**Fixed** -- the first real Windows/PowerShell run of `-Phase InitConfig`
threw `Get-ConfigOverlayJson`/`Get-ConfigOverlayReadmeText : ... Command
NotFoundException`, then a `WriteAllText` `IOException` (verify_config.json
open in another program), yet still printed `[OK] wrote/updated work-folder
config overlay`. Root cause: the repair/full-snapshot writer (`$writeOverlay`,
built via `{...}.GetNewClosure()`) is invoked from more than one call site
(the direct non-interactive path and the `-Interactive` editor's save
command), so it needs its own captured copy of `$dest`/`$dryRunFlag`/etc --
but `GetNewClosure()` only snapshots *variables* from the defining scope into
a detached session state; it does not carry over *functions*. `Get-Config
OverlayJson`/`Get-ConfigOverlayReadmeText` only exist because `ConfigOverlay.
ps1` is dot-sourced at the top of `VerifyTool.ps1`, so they were unreachable
from inside the closure. The same detachment also silently dropped this
script's `$ErrorActionPreference = 'Stop'` override (the closure's own scope
falls back to PowerShell's true default `Continue`), turning both failures
into non-terminating errors the closure just printed past. Fix: capture
`Resolve-ToolPath $Config 'ConfigOverlay'` (a plain string, which
`GetNewClosure()` DOES preserve) and re-dot-source it at the top of the
closure body, and set `$ErrorActionPreference = 'Stop'` there too, so the
functions always resolve and a real failure (e.g. a locked file) now stops
and reports instead of limping on to a false `[OK]`. See CHANGELOG.md for
the full writeup. No PowerShell in this dev environment to re-run
`Tests\Run-Tests.ps1` -- confirm both the default repair path and
`-Interactive` save on an office PC.

v2.9.27 (Mark.Boxes: StampImage -- image-recognition-only stamp for
GIFT_noGfixfile, no fixed-offset fallback):
**Added** -- a `StampImage` key usable alongside `Template` on any
`Mark.Boxes` entry: when the `Template` crop matches on the source snap PNG
(via the existing `Locate-ByImage.ps1` LockBits scan, same as the plain
Template box path), `StampImage` is inserted (native size, `ZOrder`
bring-to-front) at the matched+scaled location instead of a red rectangle.
Deliberately has NO `OffsetX/OffsetY` fallback -- unlike a plain Template
box, where "no match" falls back to a fixed guess, a StampImage box's whole
point is "only appear when the target pattern is actually found"; no match
means nothing is drawn at all (`[SKIP-STAMP]` in the console, not a
warning). New `ExcelHelpers.ps1` `Insert-PictureAtPointBringToFront` (same
raw-point-coordinate shape as `Add-RedRectangle`, but for a whole picture).
Wired the default `Mark.Boxes.GIFT_noGfixfile` (was `@()`) to
`@{ Template = 'NoGfixHit.png'; StampImage = 'already_exists.png' }` --
this is a simpler, self-contained alternative to v2.9.26's `Mark.NoteStamps`
for the same past-data-hit use case: it runs directly against the source
snap PNG via live template matching, with no dependency on
`SnapVerify.Localize.Enabled` being on or a `.note.json` sidecar existing
(both of which are off/absent by default, which is why NoteStamps alone
never fired for the operator in practice). Both mechanisms coexist --
NoteStamps still applies when a `verifyNote` payload exists; StampImage
applies to the plain-metadata `GIFT_noGfixfile` pictures that are the common
case. Inserted stamps are named `<NamePrefix><folder>_<correl>_<idx>` so
`Remove-MarkShapes` cleans them up on every idempotent Mark re-run like any
other mark shape. Documented in `mark_templates/README.txt`,
`ConfigOverlay.ps1`'s InitConfig readme text, and `verify_config.example.json`.
Static-checked only (no Windows/Excel in this dev environment) -- needs a
real `NoGfixHit.png` crop (from an actual past-data-hit screenshot) and
`already_exists.png` dropped into `mark_templates/` before this does
anything on an office PC.

v2.9.26 (Mark: NoteStamp images on verifyNote annotations -- GIFT_noGfixfile):
**Added** -- `Mark.NoteStamps` config: a new opt-in way to insert a whole
stamp image (e.g. `already_exists.png`, dropped into `mark_templates/`) next
to a `verifyNote` annotation instead of just the existing red rectangle +
`SnapVerify.NoGfixNoteColumn` text. Keyed by the note's `Folder` value (the
first `|` field of the AltText payload `EvidenceExecutor.ps1` stamps via
`Set-ShapeMetadata 'verifyNote'`), so more note kinds can register a stamp
later without touching `Mark.ps1`; only `GIFT_noGfixfile` (the F4/M6
past-data hit) is wired today. Each entry is `@{ Image; Column; RowOffset }`
(default `AF`, `RowOffset=0`). Implementation deliberately reuses the pixel
rect the `verifyNote` payload already carries (sourced from the snap-time
`<correl>.loc.json` / `.note.json` sidecars, M5/M6) instead of re-scanning
the source PNG for the orange Ctrl+F highlight band a second time at Mark
time -- the existing `Mark.ps1` verifyNote block already scales that rect's
Top to a sheet-space Y (`$top`); this just also runs `Get-RowAtOrBelow` on
it (with the same `-1` off-by-one correction `Get-PictureBottomRow` already
established, since `Get-RowAtOrBelow` returns the row *after* the target
pixel) to get an Excel row, then inserts the image at `(row + RowOffset,
Column)` via a new `ExcelHelpers.ps1` `Insert-PictureBringToFront` (same
shape as `Insert-PictureSendToBack` but `ZOrder(0)`, since a stamp must sit
on top). The inserted picture is named `<NamePrefix>verifyNoteStamp_<correl>_0`
so it is cleaned up by the existing `Remove-MarkShapes` idempotent-rerun pass
like every other mark shape. A stamp failure (image not found, bad config)
only warns -- it never fails the surrounding verifyNote mark or blocks
`isMarked`. Threaded `VerifyConfig.psd1`/`verify_config.json` ->
`VerifyTool.ps1` (`-Mode Gift` only, mirroring `NoGfixNoteColumn`) ->
`Mark.ps1 -NoteStampConfig`. Static-checked only (no Windows/Excel in this
dev environment) -- confirm the row math and stamp placement on an office PC
once `already_exists.png` is added to `mark_templates/`.

v2.9.25 (DeliverFiles unzip delivery; Review -J4; config field merge +
_SCHEMA-based repair + InitConfig walker rework):
**Added** -- (1) `DeliverFiles` now also delivers the DATA unzip subfolders
(DfSnap isZip extractions): `work\DATA\GIFT\unzip` / `work\DATA\GFIX\unzip`
files (per-correl `<correl>*` filter, same as the plain DATA files) are
copied into `J4\DATA\GIFT\unzip` / `J4\DATA\GFIX\unzip`, creating the J4
subfolder on first delivery; the specs are Optional (absent locally =
silent). (2) All four review phases (`ReviewGift/Gfix/Df/Evidence`) gained
`-J4` (menu toggle `j4`, shown as `ReviewJ4`): the workbook opens from the
DELIVERED J4 folder instead of `work\evidence` (saves land on the J4 file;
mapping bits update as usual); the J4 folder resolves via the new canonical
config below and a missing folder errors with the where-to-set-it hint.
(3) Duplicated config fields merged into canonical TOP-LEVEL fields:
`J4EvidenceDir` (was `DeliverFiles.J4EvidenceDir` = `Mail.EvidenceFolder`;
read by DeliverFiles / BackupJ4 / DeliverMail body `{2}` / Review -J4) and
`Address` (was `Reviewer.Address`; DeliverMail To). New pure unit-tested
`Get-ConfigJ4EvidenceDir` / `Get-ConfigReviewerAddress` (`ConfigOverlay.ps1`)
-- a non-empty LEGACY field still wins (old configs unchanged), but
InitConfig no longer emits the legacy duplicates and migrates a legacy value
into the canonical field on snapshot (re)generation without mutating the
live runtime config. **Fixed** -- (a) the v2.9.22 walker was actually
broken: `Expand-ConfigWalkPath`/`Get-ConfigWalkLeaves` returned `,$array`
into `@(...)`-wrapping callers, which NESTS (the v2.8.1 rule), so `w`
collapsed a whole group into ONE "System.Object[]" pseudo-field; both now
return plain arrays (verified by running the editor with scripted input
under portable pwsh 7 in this dev env). (b) InitConfig repair used to dump the whole
snapshot into a sparse overlay (it could not tell "new to the tool" from
"deliberately omitted"). Snapshots now stamp a `_SCHEMA` dotted-path
inventory (metadata: runtime-stripped, walker-excluded); the reworked
`Update-ConfigOverlayData` appends ONLY fields absent from the stamp, never
re-adds operator-deleted fields (union stamp refresh), and a stamp-less file
is only STAMPED on first repair (nothing added; f=Force / Interactive for
the full set). New pure `Get-ConfigSchemaPaths` + hashtable path helpers;
repair unit tests rewritten. **Changed** -- the `-Interactive` editor: `w`
is now a walk LOOP (pick group -> field-by-field Enter/value/`-d` delete
(y/N confirm)/q -> then "next group / s=save / Enter=back"); saving happens
inside the editor via a writer callback (`Invoke-ConfigEditorSave`), so a
locked/open verify_config.json no longer loses edits -- close it then
r=retry / Enter=back (edits kept) / q=discard; unsaved edits warn on quit.
Same writer serves repair/Force/dry-run paths. Pure logic + the editor loop
+ repair flow + DeliverFiles DATA/unzip copy were exercised under portable
pwsh 7 in this dev env (unit tests green; scripted-input editor harness;
fake-workdir DeliverFiles run); Excel COM paths (Review -J4 open) remain
static-checked -- confirm on an office PC, and re-run Tests\Run-Tests.ps1
under Windows PS 5.1.

v2.9.24 (GFIX log font size + highlight measurement fixes; InitConfig repair
mode; DfSnap isZip unzip-compare):
**Added** -- (1) `ReplaceGfix` now forces font SIZE as well as name on every
pasted GFIX receive-log line: new `Replace.GfixLogFontSize` config (default
`11`; `0` = leave workbook default), threaded `VerifyTool` ->
`ReplaceEvidence -GfixLogFontSize` -> `Invoke-EvidencePlan` ->
`Write-LogLines` (new `FontSize` param). (2) `-Phase InitConfig` on an
EXISTING `verify_config.json` now defaults to REPAIR/UPDATE: the operator's
file is kept exactly as-is (values untouched, sparse stays sparse) and only
newly-added config fields are appended (each added dotted path listed);
`f=Force` = old full-snapshot regenerate (.bak kept either way); an
unparseable file is never touched. New pure `Update-ConfigOverlayData`
(`ConfigOverlay.ps1`, unit-tested). (3) `DfSnap` rows with mapping
`isZip`=1 no longer hand df.exe the two ZIP binaries: each side's zip is
extracted into `DATA\GIFT\unzip` / `DATA\GFIX\unzip` (file named after the
correl id, SendVsGift `data\unzip` convention; discovery/entry-selection
logic mirrors SendVsGift's) and df.exe compares the extracted files. No
readable zip on a flagged side -> warn + fall back to the plain data file;
extraction failure -> row fails (`unzip/fail` progress event). **Fixed** --
the GFIX-log highlight auto-width was "longer or shorter" than the pasted
text: `Get-TextPixelWidth` now measures with GenericTypographic (+
MeasureTrailingSpaces) instead of padded plain `MeasureString`;
`Get-AutoHighlightColEnd`/`Invoke-GfixLogHighlight` accept
`FontName`/`FontSize` overrides so MarkGfix/MarkGfixLog measure with the
font the log was PASTED in (`Replace.GfixLogFontName/Size`, threaded by
VerifyTool) instead of trusting the cell-font read; the column-width read
uses `$ws.Columns.Item($c)`. (Beware: PS variables are case-insensitive --
the measurement locals are `$measureFont`/`$measureSize` because
`$fontName` would silently overwrite the `$FontName` param.) Pure logic
unit-tested; zip helpers exercised with real archives under PS 7; COM/GDI+
paths static-checked only -- confirm size-11 paste, highlight width,
InitConfig repair, and DfSnap unzip on an office PC.

v2.9.23 (Mark: image-recognition box placement + GFIX highlight auto-width +
forced GFIX log font):
**Added** -- (1) `Mark.Boxes` entries can now carry an optional `Template`
key (see `mark_templates/README.txt`); when set, `Mark.ps1` tries
`Locate-ByImage.ps1` (LockBits template match) against the original snap PNG
before drawing the red rectangle, scales the hit to the inserted picture's
on-sheet point size, and falls back to the existing fixed
`OffsetX/OffsetY/Width/Height` box on any miss (no Template, missing file, no
match) -- this never blocks Mark. Per-box `Tolerance`/`PadX`/`PadY`
overrides; new `Mark.TemplateDir` / `Mark.ImageMatch.Tolerance` config, new
`mark_templates/` folder (ships empty; needs real reference crops captured
on an office PC). (2) `-Mode Gfix`'s GFIX-log yellow highlight (also
`MarkGfixLog.ps1` standalone) is now sized to the Command: row's ACTUAL
pasted text width by default (`GfixLog.AutoHighlightWidth`, default `$true`)
instead of always filling the fixed `HighlightColStart..HighlightColEnd`
range -- new `ExcelHelpers.ps1` helpers `Get-TextPixelWidth` (GDI+
`MeasureString`, needs `Add-Type -AssemblyName System.Drawing`, now loaded by
both `Mark.ps1` and `MarkGfixLog.ps1`) and `Get-AutoHighlightColEnd` (walks
the sheet's real column widths to find where the measured text lands, capped
at `HighlightColEnd` so this only ever tightens the old fixed-width
behavior, plus `GfixLog.HighlightPadCols` extra columns of padding); the
column-accumulation math is split into a pure `Get-ColumnsForWidth` and
unit-tested (`Tests\Test-ExcelHelpers.ps1`). `AutoHighlightWidth = $false`
restores the old fixed-width behavior. (3) `ReplaceGfix` now forces every
pasted GFIX receive-log line to a fixed font -- `Write-LogLines`
(`ExcelHelpers.ps1`) gained a `FontName` parameter, threaded through
`EvidenceExecutor.ps1`'s `Invoke-EvidencePlan` and a new
`ReplaceEvidence.ps1 -GfixLogFontName` parameter (default `'MS Gothic'`, the
ASCII-typeable alias Windows/Excel resolve to the Japanese fixed-width font
"MS ゴシック" -- kept ASCII per this repo's source-encoding rule rather than
embedding the raw full-width Japanese name); new `Replace.GfixLogFontName`
config field (blank leaves the workbook's default font untouched).
`VerifyTool.ps1` threads all of the above from `VerifyConfig.psd1`/
`verify_config.json` into `Mark`/`MarkGfixLog`/`Replace` dispatch;
`ConfigOverlay.ps1`'s `excel` editor group now includes `GfixLog`, and its
README text + `verify_config.example.json` document the new fields. Pure
logic (`Get-ColumnsForWidth`) is unit-tested; the COM/GDI+ paths (template
match scaling, column-width reads, font assignment, MeasureString DPI
assumption) are static-checked only (no Windows/Excel in this dev
environment) -- confirm image-match placement (needs real templates first),
the auto-width highlight, and the forced log font on an office PC.

v2.9.22 (InitConfig editor: grouped field walker -- no manual path typing):
**Added** -- the `-Phase InitConfig -Interactive` editor gained a `w` command
next to the existing `v`/`e`/`d`/`s`/`q` ones. Picking `w` and a group (by
number or key, same lookup as the other group commands) now WALKS every
editable field in that group one at a time -- showing the current value and
prompting Enter=keep / a value=set it / `-del`=delete it / `q`=stop walking --
instead of requiring the operator to already know and type the exact dotted
JSON path (e.g. `Mark.Boxes.GIFT_HM.0.OffsetX`) for every field they want to
touch. New pure-ish helpers in `VerifyTool.ps1`: `Expand-ConfigWalkPath`
recurses a group's top-level paths down to leaf fields (hashtables always
recurse into every key; an array recurses by index only when every element is
itself a hashtable -- structured records like `Mark.Boxes` entries or
`PhaseOrder` rows -- otherwise the whole array, e.g. `Mail.BodyLines`, is one
atomic JSON-edit leaf, same as before); `Get-ConfigWalkLeaves` expands a
group's `Paths` (including the `all` group's `*`) into the leaf list;
`Invoke-ConfigFieldWalk` drives the per-field prompt loop (edits apply to the
in-memory `$Data`, same as `e`/`d`; still only written to disk via the main
menu's `s` + typed `YES` confirmation) and asks for confirmation before
walking a group with more than 30 fields (e.g. `phase`/`PhaseOrder`, which
expands to one leaf per `Key`/`Field`/`Label`/`Status`/`BitValue` per phase
row). `v`/`e`/`d` are unchanged and still take a manual path for operators who
already know exactly what to touch. Updated the editor's intro text, the
`InitConfig` phase notes, `ConfigOverlay.ps1`'s `_README` snapshot text and
`Get-ConfigOverlayReadmeText`, `README.md`, and `verify_config.example.json`
to describe `w` alongside `v`/`e`/`d`. Pure logic (`Expand-ConfigWalkPath` /
`Get-ConfigWalkLeaves`) is straightforward path-tree recursion with no
Excel/COM involvement; `Invoke-ConfigFieldWalk`'s console I/O is static-
checked only (no PowerShell in this dev environment) -- confirm the walk
prompts/confirmations on an office PC.

v2.9.21 (DeliverFiles: sheet-level replace instead of whole-file overwrite; new BackupJ4 phase):
**Changed** -- `DeliverFiles.ps1`'s evidence-Excel step no longer overwrites
the whole J4 workbook. It now works like `Align.ps1` but in the opposite
direction: it opens the CORRESPONDING J4 workbook and replaces its GIFT
受信結果 / GFIX受信結果 / GIFTデータvsGFIXデータ sheets (the operator's own
captured evidence -- exactly `Align.ps1`'s `Get-AlignRecvSheets` set) with
the matching sheets from the work evidence workbook, in place
(`Sync-DeliverSheets`, mirroring `Align.ps1`'s `Sync-Sheet` with source/dest
swapped: the J4 sheet is deleted and the work sheet copied into its old slot
to preserve position, or appended when J4 doesn't have it yet). Every other
J4 sheet (host-managed send sheets, etc.) is left untouched. When J4 has no
workbook yet for an Excel_NAME (first delivery), the phase falls back to the
old whole-file copy as a bootstrap. `-Backup` still backs up the whole J4
file (into `J4EvidenceDir\_bak`) before its sheets are replaced. **Added** --
new standalone `BackupJ4.ps1` phase (alias `Bk`/`BkJ4`, menu key `bk`):
read-only against J4, copies each targeted Excel_NAME's current J4 workbook
into a local, timestamped folder (`DeliverFiles.BackupLocalDir`, default
`<WorkDir>\bk`) so the operator can keep a local rollback point before
running the now-in-place-editing DeliverFiles. `VerifyTool.ps1` wires
`-Phase BackupJ4` the same way as `DeliverFiles` (J4EvidenceDir resolution
falls back to `Mail.EvidenceFolder`). Static-checked only (no
Windows/Excel/Edge in this dev environment) -- confirm the sheet-replace
(including the first-delivery bootstrap fallback and the same-leaf-filename
Excel-COM handling) and the BackupJ4 copy on an office PC.

v2.9.20 (DeliverFiles rework -- no source deletion + full-width J4 dedup; DeliverMail/DeliverFiles config error messages; GfixLogDownload log naming):
**Changed** -- `DeliverFiles.ps1` no longer has a "move" mode at all: it only
ever **copies**, and source DATA files (GIFT/GFIX) are never deleted. The old
`-MoveData` switch (and `DeliverFiles.MoveData` config key) is removed --
replaced by `-SkipExcel` / `-SkipData` (default: copy BOTH the evidence Excel
and the DATA files; either can be excluded for a given run) and a new
`-Backup` switch that copies any J4 file this phase is about to
overwrite/remove into `J4EvidenceDir\_bak\<name>.<timestamp>.bak` first.
**Fixed** -- (1) if the local evidence Excel was saved without the configured
`Workbook.ExcelPrefix` (e.g. cloned before the prefix was set), DeliverFiles
now falls back to finding it by the bare `Excel_NAME` and still writes the J4
copy under the fully-prefixed name (previously it copied under whatever name
the source happened to have, so J4 could end up with an unprefixed file).
(2) J4 can carry a stale copy of the same workbook typed with full-width
ASCII characters (e.g. `０` vs `0`) while the work-folder copy is half-width;
`Find-WorkbookByExcelName`'s tolerant match already handles this on the READ
side, but on delivery this used to leave BOTH the old full-width J4 file and
the new half-width copy sitting side by side. DeliverFiles now scans J4 for a
same-stem full-width variant (`WorkbookResolver.Get-FullWidthWorkbookCandidates`,
already used elsewhere for this), and -- after asking the operator to confirm
-- removes it so only the half-width (work-folder) name remains; `-Backup`
saves the removed file first. (3) The interactive/CLI dispatch in
`VerifyTool.ps1` now fails DeliverFiles with an explicit, actionable message
when `DeliverFiles.J4EvidenceDir` / `Mail.EvidenceFolder` are both unset
("set DeliverFiles.J4EvidenceDir ... in verify_config.json -- run -Phase
InitConfig") instead of only the child script's bare `-J4EvidenceDir is
required` error -- these two fields (like `Reviewer.Address` for
DeliverMail) are intentionally blank in the committed `VerifyConfig.psd1`
(scrubbed of personal defaults so nobody's real path/address ships in the
repo) and must be set per work folder via `verify_config.json`; DeliverMail's
missing-reviewer error got the same "where to set it" treatment.
`Show-PhaseNotes` for DeliverMail now spells out every `{n}` placeholder
in `Mail.SubjectTemplate` / `Mail.BodyLines` and which config field feeds it.
**Fixed** -- `GfixLogDownload.ps1` named downloaded logs
`<JobNo>_<timestamp>_<originalName>.log`, but GoAnywhere itself names the
downloaded file after the job number, so `<originalName>` was always
`<JobNo>.log` too -- the first and last filename fields were always
identical. The job number is now replaced with the mapping `JOB_NAME`(s)
that needed it (joined with `+` when more than one correl shares a job, the
duplicate-IF_NO case from v2.9.18), falling back to the job number only if
no JOB_NAME is known. Static-checked only (no Windows/Excel/Edge in this dev
environment) -- confirm the full-width J4 prompt, the prefix fallback, and
the new GfixLogDownload log filenames on an office PC.

v2.9.19 (Common.ps1: Edge activation robustness -- promote JenkinsSnap's process-handle fix):
**Fixed** -- the operator-reported symptom of "auto-switch to GoAnywhere after
pressing Enter" occasionally not working traces to `Common.ps1`'s shared
`Activate-EdgeWindow` (called by `Switch-ToEdge`, which `GfixLogDownload` /
`MqSnap` / `HmSnap` all use): it activated Edge purely via
`$Shell.AppActivate("Microsoft Edge")` -- a substring match against the
window TITLE -- and discarded the success/failure return value entirely
(`$null = $Shell.AppActivate(...)`). If Edge's title text doesn't contain
"Microsoft Edge" verbatim at that moment (title format changes across Edge
versions, an app-mode/PDF tab with a different title, multiple Edge windows
open), `AppActivate` silently returns `$false` and the old code proceeded to
"activate" whatever window already happened to be in the foreground, with no
warning that anything went wrong. `JenkinsSnap.ps1` had already hit and fixed
this exact flakiness for itself, independently, with a process-name lookup
(`Get-EdgeMainWindowHandle` scanning for `msedge.exe` by `MainWindowHandle`,
title match only as a fallback) -- but that fix was never carried back into
the shared `Common.ps1` helper, so GfixLogDownload/MqSnap/HmSnap were still
on the old, title-only path. Promoted the process-handle-first approach into
`Common.ps1.Activate-EdgeWindow` verbatim (byte-for-byte the same logic
JenkinsSnap already had proven); removed JenkinsSnap's now-redundant local
copy (`Activate-JenkinsEdgeWindow`) in favor of the shared one. See the new
TODO below for the related next-stage plan (ReplaceGfix duplicate-candidate
confirmation). Static-checked only (no Windows/Edge in this dev environment)
-- confirm on an office PC.

v2.9.18 (GfixLogDownload: duplicate-IF_NO fix -- download by job number, verify by content):
**Fixed** -- one IF_NO commonly feeds more than one downstream SS_CODE receive
job (e.g. `JIDSC02S` and `JIDSC03S` both off `IF5001_001`), so the GoAnywhere
completed-jobs list carries duplicate rows with the identical project name.
`GfixLogDownload.ps1` used to find its target row by `Ctrl+F`-searching that
project-name text, which cannot tell duplicate rows apart -- both correls
landed on the SAME physical job (confirmed from a real run: both `JIDSC02S`
and `JIDSC03S` "moved" the identical `1000002995601.log`), so one correl's
real log was never downloaded, and -- worse -- the old code marked
`GFIX_log=1` on ANY new file appearing in Downloads with no content check, so
the miss was silent (only surfacing later as a `[MISS-REQ]` in `ReplaceGfix`,
by which point the mapping already claimed the row was done). Rewritten:
Ctrl+A the LIST page once (`Read-PageText.ps1`, same technique the SnapVerify
phases already use), parse it with the new pure `GfixJobList.ps1`
(`ConvertFrom-GfixJobListText`, data rows identified by a numeric JobNo regex
-- no Japanese literals needed), group pending mapping rows by normalized
IF_NO, and for each IF_NO download **every** matching job (`Get-
GfixJobListRowsForIf`, receive-side only) keyed by its unique job number --
not just the first Find-in-page hit. Detail pages now open via `Ctrl+F
<job number>, Enter, Esc, Enter` (job number is itself the link target; no
Shift+Tab needed, confirmed against the real page). Logs are named
`<JobNo>_<timestamp>_<originalName>.log`, never after a correl. After
downloads, every pending correl is resolved by content match via the
existing, unit-tested `Find-GfixLogForCorrel` (`GfixLog.ps1`, unchanged --
its any-`*.log`-in-dir fallback already does exactly the disambiguation
needed): a real match sets `GFIX_log=1`; otherwise `GFIX_log=2` with the
match error logged. An IF_NO with zero matching rows in today's list is a
hard miss -- resolved immediately to `GFIX_log=2` (no interactive prompt,
since there is no GoAnywhere state to retry against) and the run continues
with the next IF_NO. `GFIX_log` is a plain value column (not a bitmask) with
the same NG=2-still-pending convention as `SendVsGift`/Mq/Hm/JenkinsSnap: a
local `Test-GfixLogDone` (`== '1'`) replaces the generic `Get-PendingRows`
call, since `2` is non-empty/non-'0' and would otherwise read as done. New
pure lib `GfixJobList.ps1` + `Tests\Test-GfixJobList.ps1` (fixtures include
the real duplicate-IF_NO case from a live run). COM/Edge parts are static-
checked only (no PowerShell/Excel/Edge in this dev environment); confirm the
job-number Ctrl+F open sequence and the end-to-end download+match flow on an
office PC before trusting it in production.

v2.9.17 (Align same-name workbook open fix + sheet-order preserve):
**Fixed** -- (1) Align reported every sheet "missing in J4" with an empty
"J4 sheets present:" list. Excel cannot open two workbooks with the same leaf
filename in one instance, and the J4 baseline shares the *identical* filename
with the work evidence by design. With `DisplayAlerts=$false` the second
`Workbooks.Open` (J4) returns `$null` instead of erroring, so `$j4Wb` was null
and every sheet read as missing (no `StrictMode`, no `[FAIL]`). The only
workbook that synced (`JJODWDB2`) just happened to have a full-width `W` in its
J4 filename, so its leaf name differed and both could open. `Align.ps1` now
opens J4 via `Open-J4Safely`: on a leaf-name collision (or same path) it copies
J4 to `%TEMP%\verify_j4_<guid>.xlsx` and opens the copy (removed in `finally`);
a null-workbook guard now throws instead of misreporting. (2) `Sync-Sheet`
preserves sheet order -- it copies the J4 sheet to *before* the work sheet then
deletes the work sheet, instead of delete-then-insert-after (which shifted
indices and scrambled the order). COM/Excel parts static-checked only; confirm
on an office PC + Excel.

v2.9.15 (Align Host->Open default + J4 no-content guard + picture-aware diff):
**Fixed** -- (1) Align always reported every sheet "missing in J4". With
`Align.HostSystemTypes` unset the migration type is `Unknown`, whose scope was
the three *receive* sheets -- but J4 baselines never carry recv sheets, so
nothing ever synced. `Align.ps1` now defaults an unclassifiable migration to
`Align.DefaultMigrationType` (default `HostToOpen`): delete + copy the work
send-data / GIFT-send-result / GFIX-send-result sheets from J4, in order.
(2) The image-based send-data sheet compared as "same" (value-only
`Compare-SheetGrid`) and was skipped; new picture-aware `Compare-AlignSheet`
also compares pasted-picture count so it syncs, and an already-aligned sheet is
correctly `[same]`/skipped. **Added** -- J4 "no contents" guard
(`Get-AlignSheetKind` + `Test-J4SheetPrepared`, unit-tested): a send-data sheet
needs >=1 picture and a send-result sheet needs > `Align.MinSendResultRows`
(default 3) text rows, else `[NO CONTENTS] ... replace skipped`. New COM
`Get-SheetMetrics` (rows/cols/flat + PictureCount + TextRowCount); new config
keys `Align.DefaultMigrationType` / `Align.MinSendResultRows` threaded from
VerifyTool. Pure logic + tests via `Tests\Run-Tests.ps1`; COM paths
static-checked only -- confirm on an office PC + Excel.

v2.9.14 (DfSnap df.exe path: configurable default + first-run prompt):
**Added** -- `Df.DefaultExePath` (default `C:\tools\DF\DF.exe`) is the df.exe
path the first-run prompt pre-fills (Enter accepts); `Df.ExePath` stays empty by
default (= ask on first run) and, when set, locks the path so the prompt is
skipped. VerifyTool now remembers the resolved path in `verify_session.json`
(`DfExePath`) and reloads it on startup, so the operator is prompted only on the
first DfSnap run. Resolution: CLI `-DfExePath` > session > `Df.ExePath` >
prompt(`Df.DefaultExePath`). The first-run prompt + persistence live in
VerifyTool's DfSnap dispatch; `DfSnap.ps1` gained a `-DefaultExePath` param so its
standalone prompt offers the same default. COM/Excel parts static-checked only;
confirm the prompt + persistence on an office PC.

v2.9.13 (snap TimeCheck menu toggle + -Add owner filter):
**Added** -- (1) the HM/MQ/Jenkins snap phases now expose a `tc` interactive
menu option that toggles the run-time window check per run. It seeds from
`SnapVerify.TimeCheck` (still usually off) and threads `$State.TimeCheck` into
the three snap dispatch blocks, so the operator can turn the window check on for
a single run without editing config. (2) `Generate-HostOpenMapping -Add` now
composes with the owner filter: explicit `JOB_NAME` / `Correl_ID_M` /
`Excel_NAME` selectors are looked up in the WBS (col A) and dropped when their
owner cell (col P) belongs to another operator; a JOB_NAME absent from the WBS
is kept (temp / not-yet-listed) and reported in the warnings. Owner matching
moved to a new pure `OwnerFilter.ps1` (`Test-OwnerMatch` + `Select-JobsByOwner`,
unit-tested in `Tests\Test-OwnerFilter.ps1`); the WBS scan helper
(`Build-WbsJobOwnerMap`) is COM and static-checked only. Pure logic + tests run
via `Tests\Run-Tests.ps1`; confirm the snap toggle + Excel scan on an office PC.

v2.9.12 (SnapVerify field fixes after the first office-PC run of M6):
**Changed** -- the run-time window check is now OFF by default on every snap
phase. New `SnapVerify.TimeCheck` (default `$false`, threaded to Hm/Mq/JenkinsSnap):
detection still flags missing data / abends / missing files, but skips the
run-time prompt and +-tolerance compare unless `TimeCheck = $true` (the window is
mostly nice-to-have and the prompt slowed every run). **Fixed** -- (1)
`Resolve-SnapRunTime` now also accepts time-only input `HH:mm:ss` / `HH:mm`
(1- or 2-digit hour, anchored to today) on top of the `yyyy/MM/dd HH:mm[:ss]`
forms, validates input, and a blank/garbage tolerance no longer zeroes the
default (new unit tests). (2) After the HM `o/n/s` ask and the HM/MQ page-kind
sentinel retry the shell is foreground; the phase now `Switch-ToEdge` before
continuing so subsequent keystrokes hit the Edge page, not the CLI. (3) NoGfix
poll no longer waits out the full timeout: `Wait-JenkinsPageReady -RequireTerm`
is `$false` for NoGfix, so a loaded list page is ready the moment it classifies
as a Jenkins result (the correl is expected absent). (4) A NoGfix row that reads
OK deletes any stale `<correl>.note.json`; MarkGift's past-data note now comes
from `ProjectLabels.ps1` (`NoGfixPastData`) and its `verifyNote` branch is guarded
to `-Mode Gift`. COM/Excel parts are static-checked only; confirm on an office PC.

v2.9.11 (SnapVerify M6: NoGfix past-data annotation -- by codex): F4 wired end to
end. `GiftJenkinsNoFile` runs detection (`Test-JenkinsFile -ExpectExists:$false`):
an unexpected file sets `GIFT_noGfixfile_snap = 2` and, when `Localize.Enabled`,
writes `snap\GIFT_noGfixfile\<correl>.note.json` (PixelRect / FileDateTime /
Reason). ReplaceEvidence stamps the NoGfix picture's AltText `verifyNote|folder|
correl|x,y,w,h|imageWidth|fileDateTime`; MarkGift scales pixel->point via
`Shape.Width / imageWidth`, draws the red box on the file-time field, and writes
`過去分データー` to `SnapVerify.NoGfixNoteColumn` (default `AZ`) on the picture's
row. Annotation requires `SnapVerify.Localize.Enabled = $true` (default `$false`).

v2.9.10 (JenkinsSnap page-body click no longer navigates into a queued job):
**Fixed** -- consecutive Jenkins screenshots intermittently opened a job page
(`.../job/sc_str1_50_21_stop_appserver/`) and the second file in a `TO_code`
group came out unscreenshotted. The per-row flow used `Click-PageBody` (a fixed
`Left+150, Top+150` left-click) to focus the page before `Ctrl+F` and before the
page-text read. `(150,150)` lands in the Jenkins LEFT sidebar, and when a build
is queued the "Build Queue" (`実行予定のビルド`) widget shows a job hyperlink there,
so the click navigated Edge into that job; later captures in the group then shot
the wrong page and the correl `Ctrl+F` matched nothing. The queue widget only
renders while something is queued -- hence the "didn't used to happen"
intermittency. The `(150,150)` click was doing double duty -- focusing the page
AND collapsing the previous row's `Ctrl+A` select-all so it was not captured in
the next screenshot -- so it is REPLACED (not removed) with a new
`Click-JenkinsPageCenter` that clicks the window centre, which on these Jenkins
pages carries no hyperlink (confirmed by the operator). The centre click still
clears the selection and focuses the page; `Esc` does NOT clear an Edge text
selection (an Esc-only fix would leave the select-all highlight in the next
capture). Used before `Ctrl+F` (clears the prior selection before the
screenshot) and in `Get-JenkinsPageTextOnce` (document focus for the read) --
same approach as `MqSnap`'s `Click-MqPageCenter`. `MqSnap`/`HmSnap` keep their
own clicks (MQ/HM pages are text, not hyperlink lists, so no navigation hazard).
Also removed mojibake comment decorations repo-wide: box-drawing `─` (U+2500) and
em-dash `—` (U+2014) comment rules, which mojibake to `笏笏`-style garbage on
CP932, were replaced with ASCII across all 12 affected scripts (JenkinsSnap 645x,
MarkGfixLog/Validate 267x, ExcelHelpers 270x, Mark 182x, ExcelSnap, Common,
ReviewEvidence, Generate-HostOpenMapping, GfixLogDownload, ReplaceEvidence,
JenkinsDownload). ~13 older scripts still carry raw Japanese comments (e.g.
`# 正常終了`) -- a separate, larger `[char]` migration, untouched here.
This is COM/Edge wiring, static-checked only; confirm on an office PC + Excel.

v2.9.9 (SnapVerify ASCII fix + M5 pixel localisation):
**Fixed** -- `SnapVerify.ps1` threw `The variable '$script:SV_Abend' cannot be
retrieved because it has not been set.` on the JP-locale host: the file still
held ~18 raw-Japanese comments, and PS 5.1 reading a no-BOM `.ps1` as CP932
misreads those multibyte bytes, shifting tokenisation (operator's stack showed
line 75 vs the clean file's 79) and dropping the top-of-file `$script:SV_Normal`
/ `$script:SV_Abend` assignments; under `Set-StrictMode` (JenkinsSnap sets
`Latest`) the unset read throws instead of yielding `$null`. Replaced every
non-ASCII char in `SnapVerify.ps1` + `Find-ActiveHighlightRow.ps1` with ASCII
(runtime Japanese still comes from `[char]`), so the file tokenises identically
on every codepage -- same class of fix as v2.9.8's Test-SnapVerify. **Added M5
(F5 pixel localisation)**: pure unit-tested `Get-MatchedRowIndex` (screen row of
the verdict's newest-wins row), `Get-RowPixelRect` (HM/MQ `Row1Top+(n-1)*RowHeight`
geometry, same as Find-Abend), `Get-JenkinsHighlightRect` (orange Ctrl+F band ->
rect), `New-SnapLocRect` + `Save-SnapLocSidecar` (write `snap\<folder>\<correl>.loc.json`
with x/y/w/h + imageWidth for Mark's pixel->point scaling). Non-pure glue
`SnapLocalize.ps1` (`Write-SnapLocalize`, System.Drawing + highlight scan, swallows
all errors) is dot-sourced by Hm/Mq/JenkinsSnap and writes the sidecar after each
verdict; VerifyTool threads the new `SnapVerify.Localize` config block (Enabled
`$false` by default; HM/MQ geometry zeros until `Calibrate-HmGeometry.ps1` runs,
so the leg is inert until opted in -- Jenkins needs no geometry). M6 (NoGfix
annotation: consume the sidecar AltText -> Mark + AZ note) remains. Pure logic +
tests run via `Tests\Run-Tests.ps1`; the COM/GDI+ wiring is static-checked only
and needs an office-PC run to confirm.


v2.9.8 (SnapVerify M4 -- HM instant NG detection + Test-SnapVerify parse fix):
`HmSnap.ps1` rewritten from the legacy bare-CSV version to the modern stack --
MappingStore (atomic writes), ProgressLog events, and the pure `SnapVerify.ps1`
detection library -- while keeping HM's per-`TO_code` appl grouping (one HM page
opened per appl). It now runs F1 (plan 2.3): after the search it polls the page
text (A2), classifies the page (`Get-SnapPageKind -Phase Hm`, sentinel A3),
archives the Ctrl+A text as `snap\<Stage>_HM\<correl>.txt` (A1), screenshots, then
`ConvertFrom-HmPageText` + `Test-HmAbend` decide: ok->`<Stage>_HM_snap`=1 (newest
run in the time window ended normally; earlier in-window abends become
retried-ok warnings), ng->2 (newest in-window run is an abend), ask->operator
chooses o=OK(1)/n=NG(2)/s=skip(pending)/q=quit (0 rows, no in-window rows, or
no-time-mode abend -- plan 4.F1). Out-of-window historic abends only warn (never
auto-NG). NG=`2` stays pending (re-offered next run) + end-of-run NG summary. A
one-time batch run-time prompt (`Resolve-SnapRunTime`) fills empty
`Expected_Time` cells (plan 2.2). Off-page kinds (OuterFrame/Empty/Unknown) stop
and ask `r=retry / s=skip / q=quit`. Pending uses a local `Test-HmSnapDone`
(done == '1') so NG rows aren't hidden; VerifyTool's GiftHmSnap/GfixHmSnap
dispatch passes the SnapVerify + ExpectedTime config (mirrors GiftMqSnap).
`SnapVerify.Enabled=$false` reverts HmSnap to pure screenshot. The focus-safe
pattern from v2.9.7 is preserved (per-row `Reset-FocusToBody` only; `Switch-ToEdge`
only from a console-foreground point). F1's pure helpers + unit tests shipped in
M1; this is wiring only. **Also fixed**: `Tests\Test-SnapVerify.ps1` failed to
parse on the JP-locale host -- three `Assert-Equal` messages embedded raw
Japanese *inside single-quoted strings* (plus raw CJK comments); under PS 5.1's
ANSI codepage (CP932) a Shift-JIS lead byte swallowed the closing quote, running
the string away until the parser reported a bogus "missing terminator" far below.
Replaced all non-ASCII with ASCII per the ASCII-source rule (assertions still
compare the `[char]`-built values, so test logic is unchanged). M5 (pixel
localisation) and M6 (NoGfix annotation) remain.

v2.9.7 (MqSnap focus regression fix): `MqSnap.ps1` no longer `Switch-ToEdge`s at
the top of each per-row attempt (its `Alt+Tab` toggled to the console after the
previous screenshot, so `Click-PageBody` clicked the wrong window). Restored the
known-good pattern: `Switch-ToEdge` once before the loop and only inside the
interactive branch; per-row refocus is `Reset-FocusToBody` (AppActivate by title
+ `Click-PageBody`) only. Detection/screenshot behavior unchanged.

v2.9.6 (SendVsGift OCR-dropout tolerance + clean-read preference): the JP OCR
recognizer drops runs of characters from long ASCII record strings (field: ~12
digits lost off the first record), which symmetric edit-distance scored as a hard
`mismatch`. `Compare-SendRecordCheck` now has an OCR-dropout tier -- after the
exact / prefix / compact similarity checks fail, a materially-shorter side that is
almost entirely an in-order subsequence of the other (length ratio `<= 0.85`,
`LCS / shorter >= 0.9`) scores `fuzzy` not `mismatch`; too short to judge (`< 6`
chars) scores `unknown` (never a false mismatch). New pure `Get-SendLcsLength`
(longest common subsequence) backs it (unit-tested). `Find-SendRecordByRowNumber`
now returns the LONGEST record after a row label -- each image is OCR'd with both
`ja` + `en-US` and the lines are merged, so the fullest read wins and a clean
en-US read is not shadowed by a dropped ja one. The verdict rule is now explicit:
a matching row count (max row label == gift `MaxRowNumber`) keeps the verdict
`ok` even if a record disagrees -- record text is too OCR-noisy to auto-NG; the
disagreement is still surfaced in the per-field `Checks`. Pure-logic only; run
`Tests\Run-Tests.ps1` on Windows to confirm.

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

- **NEXT: ReplaceGfix duplicate-candidate confirmation** — since v2.9.18,
  `GfixLogDownload` deliberately downloads *every* GoAnywhere job matching a
  needed IF_NO (not just one), because duplicate-IF_NO rows are common and
  content matching (`Find-GfixLogForCorrel`) is what actually decides which
  correl a log belongs to. This means `log\` can now legitimately hold more
  than one candidate log for a single correl (e.g. genuine retries of the
  same job) more often than before. Today `Find-GfixLogForCorrel` silently
  picks the newest candidate by its `Command:` line timestamp and only prints
  a `[WARN] N logs matched; chose newest (...)` — it never stops for operator
  confirmation. Planned next step: when `ReplaceGfix` (or `GfixLogDownload`'s
  finalize step) hits a multi-candidate `Warning`, show the operator each
  candidate's file name + parsed timestamp and require an explicit pick
  (Enter = accept newest, or choose another) before the log is pasted into
  the evidence workbook / before `GFIX_log`/`isReplaced` is marked done,
  instead of trusting "newest wins" silently. Needs: (1) deciding where the
  prompt belongs (`GfixLogDownload` finalize vs `ReplaceGfix`'s log op in
  `EvidenceExecutor.ps1` — the latter runs later and closer to when the log
  is actually inserted, so may be the more meaningful place to ask), (2) a
  non-interactive fallback (keep "newest wins" under `-NonInteractive`, same
  as the rest of this codebase's interactive/non-interactive split).

- **Mark: image-recognition placement for the red rectangle -- WIRING DONE**
  (v2.9.23), calibration still open. `Mark.Boxes` entries can now add a
  `Template` key (filename resolved against `Mark.TemplateDir`, then
  `mark_templates/`); when present, Mark.ps1 calls the existing
  `Locate-ByImage.ps1` (LockBits template match) against the original snap
  PNG (`<WorkDir>\snap\<folder>\<correl>.png`, the same file
  ReplaceEvidence pasted) instead of trusting a fixed offset, scales the hit
  from source-PNG pixels to the inserted picture's on-sheet point size, and
  falls back to the configured `OffsetX/OffsetY/Width/Height` box whenever
  there is no Template, the file is missing, or no match is found -- so this
  degrades gracefully and never blocks Mark. Per-box `Tolerance`/`PadX`/`PadY`
  overrides; console lines are tagged `[MARK-IMG]` (matched) vs `[MARK]`
  (fixed offset fallback) so a run makes it obvious which path was used.
  Still needs: real reference template PNGs per mark target (a small,
  visually distinctive crop of the target field -- see
  `mark_templates/README.txt` for the how-to) captured from real evidence,
  and an office-PC/Excel session to calibrate and confirm the pixel->point
  scaling -- no Windows/Excel in this dev environment, so `mark_templates/`
  ships empty and this stays fixed-offset-only until templates are added.

- **GiftJenkinsNoFile: callout bubble on the past-data mark** — SnapVerify M6
  (v2.9.11) already detects an unexpected *old* file in the no-GFIX-expected
  case (`Test-JenkinsFile -ExpectExists:$false`), draws the red box on the
  file's timestamp field, and stamps `過去分データー` (`ProjectLabels.NoGfixPastData`)
  into `SnapVerify.NoGfixNoteColumn` (default `AZ`). Requested follow-up: add
  an actual callout/comment-bubble shape next to the mark (not just the AZ
  column text) so the "this is old/past data" note is visible directly on the
  evidence picture itself. Needs a design decision on the shape to use (Excel
  `msoShapeCallout` via COM ~= `Shapes.AddCallout`, sized/positioned relative
  to the existing `verifyNote` AltText rect) plus an office-PC/Excel session
  to confirm placement -- no Windows/Excel in this dev environment.

- **Edge activation robustness DONE** (v2.9.18) — `Common.ps1`'s
  `Activate-EdgeWindow` (used by `Switch-ToEdge`, which `GfixLogDownload` /
  `MqSnap` / `HmSnap` all call after the operator presses Enter) used to
  activate Edge purely via `$Shell.AppActivate("Microsoft Edge")`, a
  title-substring match, and silently discarded its success/failure return
  value. `JenkinsSnap.ps1` had already independently fixed this exact
  flakiness for itself with a process-name-based lookup
  (`Get-EdgeMainWindowHandle` / `Activate-JenkinsEdgeWindow`, msedge.exe by
  process rather than window title), but that fix never made it into the
  shared `Common.ps1` helper the other phases use. Promoted the
  process-handle-first / title-match-fallback approach into
  `Common.ps1.Activate-EdgeWindow` (title match is now only a fallback, and a
  real `[WARN]` is printed when both paths fail instead of silently
  "activating" whatever window already happened to be foreground);
  `JenkinsSnap.ps1`'s duplicate local copy was removed in favor of the shared
  one. Static-checked only (no Windows/Edge in this dev environment) --
  confirm on an office PC that `Switch-ToEdge` reliably reaches GoAnywhere
  again.

- **SnapVerify M1–M5 done** — M1: `SnapVerify.ps1` pure library +
  `Tests/Test-SnapVerify.ps1` unit tests + `SnapVerify` config section in
  `VerifyConfig.psd1`. M2: `MqSnap.ps1` migrated to MappingStore/ProgressLog and
  wired to F2 (page-text poll, page-kind sentinel, MQ verdict ok=1/ng=2, batch
  `Expected_Time` prompt); two new pure helpers (`ConvertTo-ExpectedDateTime`,
  `Set-EmptyRunTimeCells`) are unit-tested. M3: `JenkinsSnap.ps1` wired
  to F3 (GiftRecv/GfixRecv NG=2 + summary, batch time prompt, sentinel,
  `Test-JenkinsSnapDone`); NoGfix stays pure-screenshot until M6. **M4 done** --
  `HmSnap.ps1` migrated to MappingStore/ProgressLog and wired to F1
  (page-text poll, page-kind sentinel, `Test-HmAbend` verdict ok=1/ng=2/ask with
  newest-wins in the time window, batch `Expected_Time` prompt, local
  `Test-HmSnapDone`); per-`TO_code` appl grouping preserved; VerifyTool dispatch
  passes SnapVerify+ExpectedTime config (mirrors MqSnap). **M5 done** -- F5 pixel
  localisation: pure `Get-MatchedRowIndex` / `Get-RowPixelRect` /
  `Get-JenkinsHighlightRect` / `New-SnapLocRect` / `Save-SnapLocSidecar` in
  `SnapVerify.ps1` (unit-tested) produce a `snap\<folder>\<correl>.loc.json` rect
  for the verdict's row; non-pure glue `SnapLocalize.ps1` (`Write-SnapLocalize`,
  System.Drawing + `Find-ActiveHighlightRow` scan) is dot-sourced by the three
  snap scripts and writes the sidecar after each verdict when
  `SnapVerify.Localize.Enabled` (default `$false`; HM/MQ geometry must be
  calibrated first, Jenkins uses the orange highlight). **M6 done** (v2.9.11) --
  NoGfix annotation: `GiftJenkinsNoFile` detects an unexpected file
  (`Test-JenkinsFile -ExpectExists:$false`), writes `<correl>.note.json` when
  `Localize.Enabled`, ReplaceEvidence stamps `verifyNote` AltText, MarkGift
  pixel->point scales + draws the red box and writes `過去分データー`
  (`ProjectLabels.NoGfixPastData`) to `SnapVerify.NoGfixNoteColumn` (default `AZ`).
  v2.9.12 field fixes: `TimeCheck` default-off, time-only run-time input, Edge
  refocus after prompts, NoGfix poll `-RequireTerm $false`, stale-note cleanup.
  M3/M4 copied MqSnap's `Test-MqSnapDone` pattern (done == exactly '1')
  so NG='2' rows stay pending -- `Get-PendingRows`/`Test-SnapDone` treat any
  non-'0' value as done and would hide NG rows. Design + open questions (only Q5,
  Rtncd/Rsncd semantics, is non-blocking) live in `docs/SnapVerify-Plan.md`. The
  M5 COM/GDI+ wiring is static-checked only; confirm on an office PC + calibrate
  `SnapVerify.Localize.*Row1Top/*RowHeight/*ColLeft/*ColWidth` before trusting it.

- **Generate-HostOpenMapping `-Add` + owner filter compose DONE** (v2.9.13) —
  explicit `-Add` selectors (`JOB_NAME` / `Correl_ID_M` / `Excel_NAME`) used to
  bypass the WBS owner-match scan, so jobs were added regardless of owner. They
  are now looked up in the WBS (col A) via the new `Build-WbsJobOwnerMap` and
  filtered through pure `Select-JobsByOwner` (`OwnerFilter.ps1`): a job whose WBS
  owner cell (col P) belongs to another operator is dropped (warned); a job
  absent from the WBS is kept as a temp/not-yet-listed job and reported. The
  WBS-range `-Add` path already owner-filtered (Step C) and is unchanged. Pure
  logic unit-tested in `Tests\Test-OwnerFilter.ps1`; the COM scan needs an
  office-PC run to confirm.

- **GfixLogDownload: auto-set GoAnywhere max rows to 100**
  Currently requires manual setup (default GoAnywhere list shows 20 rows — not enough for
  busy BIZ codes). Future: use SendKeys / UI automation to set the rows-per-page dropdown
  to 100 automatically after `Switch-ToEdge`, before the per-row search loop.

- **DfSnap: DfExePath configurable + first-run prompt DONE** — `Df.ExePath`
  (empty by default) holds a locked path that skips the prompt entirely; the new
  `Df.DefaultExePath` (default `C:\tools\DF\DF.exe`) is the suggestion the
  first-run prompt pre-fills (Enter accepts). Resolution is CLI `-DfExePath` >
  `verify_session.json` > `Df.ExePath` > prompt(`Df.DefaultExePath`). VerifyTool
  prompts once on the first DfSnap run, remembers the answer in
  `verify_session.json` (`DfExePath`), and passes both values to `DfSnap.ps1`
  (which keeps its own default-pre-filled prompt for standalone use). To lock a
  path and never be prompted, set `Df.ExePath` directly. COM/Excel parts are
  static-checked only; confirm the prompt + persistence on an office PC.

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
- `DfExePath` — df.exe path; remembered after the first DfSnap run so the
  prompt fires only once (seeded from `Df.DefaultExePath`)
