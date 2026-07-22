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

Versioning: use `MAJOR.MINOR.PATCH`; release headings in `CHANGELOG.md` carry the current `vX.Y.Z`. See `docs/Versioning.md` for bump rules and release automation guidance.

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
ScreenRegion.ps1        pure screen-region clamp math + Resolve-DirectionalCrop
                        (four-side snap crop resolution: CropPx + per-side +
                        per-folder overrides). Unit-tested. Dot-sourced by
                        VerifyTool.ps1.
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
ProcessTimeParse.ps1    pure HM processing start/end/duration helpers for the
                        ProcessTime phase: Get-ProcessDurationText (HH:mm:ss,
                        not clamped to 24h), ConvertTo-ProcessTimeNormalizedLine
                        (repairs OCR-injected spaces INSIDE time tokens:
                        '10 :58 :20' / '00 :00 : 0 1' -> HH:mm:ss),
                        ConvertFrom-ProcessTimeOcrLines (anchors on normalized
                        datetime tokens per OCR'd row instead of column
                        position; fuzzy '...shuuryo' status match; emits
                        PageDuration / CorrelSeen / Partial for cross-checks),
                        Select-ProcessTimeRow + Get-ProcessTimeRowRank (full >
                        partial, correl-seen > unseen, newest among equals),
                        Get-ProcessTimeOcrMissNote (why a read yielded nothing),
                        Get-NewestProcessTimeRow (newest-by-StartTime;
                        -MinimumTimeOfDay default 09:00, v2.14.0),
                        Resolve-ProcessTimeRowPlan (-Stage + sidecar-exists +
                        ProcessTime_Inserted bitmask bits + -Force -> per-row
                        NeedsOcr/NeedsWrite; v2.13.0, bitmask v2.15.0) +
                        Get-ProcessTimeMigratedInsertedValue (legacy plain
                        '1' -> bitmask '3'), Get-ProcessTimeOutputTag /
                        Get-ProcessTimeOutputFileName / Resolve-ProcessTime
                        OutputDir (config-driven output tag classification,
                        not hardcoded to JDL/JRV; -DeriveFromName derives an
                        unlisted tag from Excel_NAME chars 2-4, v2.15.2) and
                        Get-ProcessTimeCheckSummaryLine (end-of-run
                        manual-check summary line), all v2.15.0.
                        ConvertTo-ProcessTimeDateTimeValue / ConvertTo-
                        ProcessTimeDurationValue (v2.15.2) parse the
                        sidecar's formatted stamps back into real Excel
                        date/time serials for the output workbook's value
                        cells + check-formula columns.
                        ConvertTo-ProcessTimeCorrelKey
                        also strips OCR-inserted whitespace before folding
                        (v2.14.0); Select-ProcessTimeRow's -MinimumTimeOfDay
                        (default 09:00) drops HM history rows; Get-ProcessTime
                        RecordCount falls back to the count immediately before
                        the result diamond when no datestamp anchors it (JDL).
                        No Excel/OCR. Unit-tested (Tests\Test-ProcessTimeParse.ps1).
ProcessTimeCheck.ps1    pure ProcessTime output-workbook audit ("check") column
                        module (dot-source, no param(), no COM): Get-ProcessTime
                        CheckColumnSpec returns the data-driven spec for the
                        columns appended after A..H -- I 処理時間(検算) (=E-D),
                        J チェック (T/F compare of written vs re-derived
                        duration), K 件数チェック (record-count check; first
                        version flags a blank/zero/non-numeric count) -- and
                        New-ProcessTimeCheckFormula fills a formula template's
                        {0} with a row number. ProcessTime.ps1's COM-side
                        Set-ProcessTimeCheckColumns walks the spec to write the
                        headers/formulas/number-formats uniformly after the
                        data rows. Japanese headers via [char]. Unit-tested
                        (Tests\Test-ProcessTimeCheck.ps1). v2.16.0.

Clone.ps1               Phase Clone
Align.ps1               Phase Align/Precheck: compare work evidence vs J4 baseline
ReplaceEvidence.ps1     Phase ReplaceGift / ReplaceGfix / ReplaceDf (plan-driven)
ProcessTime.ps1         Phase ProcessTime: extracts each correl's HM batch
                        processing start/end time (GIFT + GFIX) and derives
                        the duration -- archived snap\GIFT_HM|GFIX_HM\
                        <correl>.txt first (ConvertFrom-HmPageText), else OCR
                        of the HM screenshot already inserted into the
                        evidence workbook: content-validated candidate
                        pictures (section, below-label, above-label; relaxed
                        candidates must show the correl id in their OCR text)
                        each read via Invoke-WinOcrFile +
                        ConvertFrom-ProcessTimeOcrLines / Select-ProcessTimeRow.
                        Writes one row per GIFT/GFIX side per correl to
                        <label>(<Tag>).xlsx evidence workbooks under
                        ProcessTime.OutputDirectory (v2.14.0; classified per
                        ProcessTime.OutputTags, default JDL/JRV but
                        extendable, e.g. JDS -- v2.15.0; a row matching no
                        tag goes to UnclassifiedTag instead of aborting the
                        run; OutputMode 'Single' writes one untagged
                        workbook instead; OutputDirectoryByTag routes a tag
                        to its own destination directory). Run after
                        ReplaceGift/ReplaceGfix. Sets ProcessTime_Inserted,
                        a bitmask since v2.15.0 (bit 1 = OCR'd, bit 2 =
                        written; a legacy plain '1' is migrated to '3').
                        -Stage Ocr|Write|Both (v2.13.0) runs the
                        extract-and-cache-to-sidecar step and the
                        write-the-output-workbooks step independently; each
                        correl's OCR result (both sides, combined) is cached
                        at snap\ProcessTime\<correl>\result.json so a
                        Write-only rerun opens no evidence workbook at all.
                        Prints an end-of-run "needs manual check" summary
                        listing every correl whose GIFT and/or GFIX side was
                        not matched (v2.15.0).
Mark.ps1                Phase MarkGift / MarkGfix / MarkDf. Each Mark.Boxes
                        entry may add a 'Template' key to try image-recognition
                        placement (Locate-ByImage.ps1 LockBits match against
                        the source snap PNG) before falling back to the fixed
                        OffsetX/OffsetY box -- see mark_templates/README.txt.
                        A box may also add BaseRow/RowHeight (GIFT_MQ) to
                        shift OffsetY for correls whose page shows a
                        different record count than the calibrated baseline
                        row; the target row/count come from a snap-time
                        <correl>.mqrow.json sidecar, else a re-parse of
                        <correl>.txt, else English OCR of <correl>.png.
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
                        Also writes <correl>.mqrow.json (the verdict's target
                        row index + record count, via Get-MatchedRowIndex)
                        so Mark.ps1 can shift the GIFT_MQ red box when a
                        correl shows other than the usual 2 records.
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
Probe-SheetFormat.ps1   read-only cell-FORMAT probe (calibration aid, v2.15.2):
                        dumps a workbook's / one sheet's column widths, row
                        heights and distinct format signatures (NumberFormat,
                        font, colors as raw BGR Longs, alignment, borders)
                        with sample addresses; optional -Json report. Use to
                        match generated output (e.g. ProcessTime) to a
                        delivery template. Has param() -> call via &.
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
`SnapLocalize.ps1`, `OwnerFilter.ps1`, `Find-ActiveHighlightRow.ps1`,
`ProcessTimeParse.ps1`, `ProcessTimeCheck.ps1`. All phase scripts have
`param()` and are called via `& $path @args`.

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

`GIFT_ProcessTime` / `GFIX_ProcessTime` are plain, informational per-side
value columns (NOT a bitmask, NOT this phase's `Get-PendingRows` field):
`0` not yet attempted, `1` start/end extracted, `2` not found. The
`ProcessTime` phase gates its own reprocessing on `ProcessTime_Inserted`,
which **is** a bitmask (v2.15.0, matching the `isReplaced`/`isMarked`/
`isReviewed` convention above): bit 1 (1) = this correl's OCR result has
been extracted and cached (per-correl sidecar under
`snap\ProcessTime\<correl>\result.json`); bit 2 (2) = the row has been
written into an output workbook; `3` = both done. A pre-v2.15.0 mapping's
plain `1` (the old "written" flag) is migrated to `3` once, on load
(`Get-ProcessTimeMigratedInsertedValue`, `ProcessTimeParse.ps1`) — a legacy
write could only ever happen after OCR succeeded, so it is never migrated to
just bit 2. `-Force` redoes whichever stage(s) `-Stage` selects regardless of
its bit. Output is written per configurable tag (`ProcessTime.OutputTags`,
default `JDL`/`JRV` but not limited to them — e.g. add `JDS`) into
`<ProcessTime label>(<Tag>).xlsx`; a result row matching no configured tag is
routed to `ProcessTime.UnclassifiedTag` (default `Other`) instead of
aborting the whole write, and `ProcessTime.OutputMode = 'Single'` writes
every row into one untagged workbook instead. `ProcessTime.
OutputDirectoryByTag` routes a tag to its own destination directory.

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

## Current state (last bump: 2026-07-22 v2.16.0)

v2.16.0 (ProcessTime: check-formula module + unified generation, pure
refactor -- no OCR change): the ProcessTime output workbook's audit
formulas -- previously inlined into `Write-ProcessTimeWorkbook`'s per-row
write loop (col I `=E-D`, col J `=IF(ROUND(F*86400)=ROUND(I*86400),"T","F")`)
where they were untestable and coupled to the data write -- are extracted
into a data-driven module and generated in one uniform pass. **Added** --
new pure `ProcessTimeCheck.ps1` (dot-source, no param(), no COM):
`Get-ProcessTimeCheckColumnSpec` returns the ordered spec for the columns
appended after the A..H data -- I 処理時間(検算) (=E-D), J チェック (T/F
compare), and a new K 件数チェック that formalizes the operator's manual
"count check" (first version: the per-row record count in col G parses to a
positive number, thousands commas stripped; blank -> blank, zero/non-numeric
-> NG; a stricter GIFT-vs-GFIX cross-row equality check is a documented
follow-up because the vertical layout groups all GIFT then all GFIX rows per
job, so a single-{0} template cannot reference the paired row). Japanese
headers via `[char]` (ASCII source). `New-ProcessTimeCheckFormula
-Template -Row` fills a template's `{0}` with the row number (pure, unit-
tested, e.g. row 5 -> `=E5-D5`). The spec formulas are SELF-GUARDING (blank/
text source cells -> "" in the cell) so the "partial row stays blank" rule
holds with no per-row inspection. **Changed** -- `Write-ProcessTimeWorkbook`
writes A..H data only; a new COM finalizer `Set-ProcessTimeCheckColumns`
walks the spec after all data rows exist to write the check headers, per-row
formulas and number formats, gated on the new `ProcessTime.EmitCheckColumns`
config (default `$true`; `$false` writes A..H only). Every table
range/border/fill/width now uses a dynamic last-column letter (A..H or
A..K). **Added (reserved)** -- `ProcessTime.OcrPreprocessBinarize` /
`OcrPreprocessThreshold` config: carried through config + CLI + docs but NOT
wired into the OCR image pipeline yet (a placeholder for a later stage).
**Notes** -- pure logic (`ProcessTimeCheck.ps1`) is unit-tested
(`Tests\Test-ProcessTimeCheck.ps1`), the COM finalizer
`Set-ProcessTimeCheckColumns` and the dynamic-range formatting are
static-checked only -- confirm the I/J/K columns' formulas + values (and
that a partial row stays blank) against a real evidence run on an office PC
with Excel.

v2.15.3 (ProcessTime: code-review hardening): three review findings against
v2.15.2. **Changed** -- tag auto-derivation is now gated on a STRICT
whole-name regex (`^[0-9A-Za-z][A-Za-z]{3}[0-9A-Za-z]{4}$` -- the full
`?XXX????` shape, exactly 8 alphanumerics with letters in the tag slot);
any non-conforming name fails safe to UnclassifiedTag instead of minting a
junk tag from a blind substring. **Added** -- OCR image preprocessing
(`ConvertTo-ProcessTimeOcrImage`, System.Drawing) as the ROOT-CAUSE fix for
the ja `9`->`3` digit misreads: every picture is upscaled (2x, auto-capped
below the WinRT OCR MaxImageDimension) + grayscaled + contrast-stretched
(1.3) at the single OCR choke point (`Read-ProcessTimeOcrLines`) before
either recognizer reads it; `<stem>_pre.png` artifacts live next to the OCR
dumps and any failure falls back to the original image. Config:
`ProcessTime.OcrPreprocess`/`OcrPreprocessScale`/`OcrPreprocessContrast`.
(A post-hoc regex check + bounding-box en-US re-read cannot catch this bug:
the misread yields a format-VALID timestamp.) **PS 5.1 note** -- the target
runtime is Windows PowerShell 5.1; pwsh 7 runs here prove parse + pure
logic only. The unit suite now bans the whole `@($var[index])` wrap shape
from ProcessTime.ps1 (both real "Argument types do not match" incidents
were that pattern); indexed-collection enumeration goes through
`ConvertTo-ProcessTimeBucketArray`. Confirm all three on an office PC.

v2.15.2 (ProcessTime: auto-derived output tags, snap-PNG OCR tier, real
date/time cells + check formulas): driven by a real 257-row JOD office-PC
run. **Fixed** -- (1) the `[FAIL] ... 引数の型が一致しません` on the 'Other'
workbook followed by a contradictory `[OK]` for the same path: the workbook
had SAVED fine -- the throw came after, from
`@($writtenRowsByCorrel[$r.CorrelId])` in the write-bit marking loop (the
PS 5.1 @()-over-hashtable-indexed-List[object] binder bug
`ConvertTo-ProcessTimeBucketArray` already works around, one call site
over), so the tag reported both FAILED and written and no write bit was
ever set; now routed through the helper + a source-guard unit test. (2) all
257 JOD rows piled into `処理時間(Other).xlsx`: new `ProcessTime.AutoDeriveTag`
(default `$true`) derives the tag from the Excel_NAME's own `?XXX????`
shape (chars 2-4: `CJODWDEJ` -> `JOD`) when no configured OutputTags entry
matches, so an unlisted project family still gets its own workbook
(`Get-ProcessTimeOutputTag -DeriveFromName`, unit-tested; UnclassifiedTag
now only catches non-conforming names). **Added** -- (a) snap-PNG OCR tier:
per-side source priority is now `snap\<Stage>_HM\<correl>.txt` -> OCR of
`snap\<Stage>_HM\<correl>.png` (the cleaner, per-correl-named original
screenshot; trusted like the section tier, source `ocr:snap-png`, dump
`<side>_<correl>_snapocr.ocr.txt`) -> evidence-workbook picture OCR.
(b) output workbook start/end are REAL Excel date/time values (OADate +
NumberFormat set before the value) and duration a real `[h]:mm:ss` serial,
with text fallback; new audit columns I `=E{r}-D{r}` and J
`=IF(ROUND(F*86400,0)=ROUND(I*86400,0),"T","F")` (blank on partial rows);
pure `ConvertTo-ProcessTimeDateTimeValue`/`ConvertTo-ProcessTimeDurationValue`
unit-tested. (c) new standalone `Probe-SheetFormat.ps1` (see file map).
**Notes** -- RECORDED recurring ja-OCR confusion: digit `9` read as `3`
(JIGPC06S: page `11:19:16/11:19:28/20260701111906/29,264` -> ja
`11:13:.../111306/23,264` while en-US read `11:19:` right); see the TODO
below for the planned en-US digit cross-check. COM paths static-checked
only -- confirm on an office PC against the same JOD folder.

v2.15.1 (ProcessTime: template-matched worksheet formatting + code-review
fixes): the office-PC formatting fix from the v2.15.0 follow-up (Yu Gothic
11pt template styling, A1-address ranges to dodge the `Cells.Item` COM
overload mismatch) went through a code review that caught three real bugs,
all fixed here. **Fixed** -- (1) the formatting block's single bare
`catch {}` silently skipped every remaining step (and printed no warning)
the moment one COM call failed, so a workbook could report success while
only partially formatted; split into one try/catch per concern (range
resolve / AutoFilter+borders / header fill+font / font+row-height /
alignment / GIFT-GFIX row fill / column widths), each logging
`Write-Warning` with the failing step and output path. (2)
`Generate-HostOpenMapping.ps1`'s snap `ID.txt`/`ID.png` bulk-ID selector --
documented as only for `-FromBizCode JOD -Owner all` -- actually fired for
ANY call omitting `-CorrelIdsM`/`-JobNames`/`-ExcelNames`, so a stale
`ID.txt` left over from an earlier JOD batch could silently hijack an
unrelated run (e.g. `-FromBizCode JRV -Owner AAA`) into a tiny ID-file-
limited temp mapping instead of the intended WBS+FromBizCode scan; gated
behind a new pure, unit-tested `Test-MappingIdBulkSelectorEnabled`
(`MappingInput.ps1`). (3) a duplicated `CHANGELOG.md` entry from the prior
commit was merged into one, versioned entry. **Notes** -- current
`ProcessTime.ps1` output formatting settings (colors/fonts/sizes/widths) are
now documented inline as a reference point; a dedicated formatting module
is still a TODO, and Start/End are written as plain formatted TEXT rather
than real Excel date/time values with a cell `NumberFormat`. Also removed an
accidentally committed `VerifyConfig_bk.psd1` backup and added a
`.gitignore` rule for `*_bk.ps1`/`*_bk.psd1`/`*.bak`; line-ending
normalization (`.gitattributes`) done as a separate commit per review
feedback. No PowerShell/Excel in this dev environment -- confirm the
per-step formatting warnings and the mapping selector gate on an office PC.

v2.15.0 and earlier (bitmask ProcessTime_Inserted + config-driven output
tags; the v2.14.x JDL/JRV split + OCR robustness; v2.13.0 staged Ocr/Write;
the v2.12.x ProcessTime OCR-parsing fixes; v2.11-v2.6 snap crop, Mark
image-match, DeliverFiles/BackupJ4, SnapVerify, config overlay, incremental
-Add, and the major MappingStore/plan-driven-Replace refactor) -- folded
here; see CHANGELOG.md for the full per-version history.

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

- **ProcessTime: ja-OCR digit 9->3 -- ADDRESSED via image preprocessing**
  (v2.15.3; originally recorded v2.15.2, recurring per operator) — the ja
  recognizer misread digit `9` as `3` on the HM font in time-of-day,
  14-digit datestamp AND record-count fields (JIGPC06S: real page
  `11:19:16 / 11:19:28 / 20260701111906 / 29,264`; ja read `11:13:16 /
  11:13:28 / 20260701111306 / 23,264`). Fixed at the root: every picture
  is preprocessed (2x bicubic upscale capped below the WinRT OCR
  MaxImageDimension + grayscale + 1.3 contrast stretch,
  `ConvertTo-ProcessTimeOcrImage`) before either recognizer reads it. A
  post-hoc regex check cannot catch this class of misread (the result is a
  format-valid timestamp), so prevention-before-recognition is the only
  clean fix. Needs office-PC confirmation on the JIGPC06S picture; if
  misreads persist there, the fallback plan remains pooling the en-US
  read's time-of-day/count digits as hints (same shape as
  `-StartDateHints`) and adopting them when the recognizers disagree only
  on 3/9 positions.

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

Machine/operator state ONLY -- see `docs/Configuration.md` for the full
layering rule (psd1 = shipped defaults, work-folder `verify_config.json` =
everything project-scoped, session = machine ephemera). Project-scoped
first-run prompt answers (e.g. `CheckSheet.Path`) persist to the work
folder's `verify_config.json` since v2.10.7, not here.

Remembered between runs:
- `WorkDir` — last work folder path
- `Owner` — mapping owner suffix (no personal default)
- `WindowWidth`, `WindowHeight`, `CropPx` — screenshot window size
- `CursorCell` — review cursor cell (default: A3)
- `CloneSourceDir` — external path for Clone phase
- `EvidenceDir` — evidence output folder (default: `<WorkDir>\evidence`)
- `DfExePath` — df.exe path; remembered after the first DfSnap run so the
  prompt fires only once (seeded from `Df.DefaultExePath`)
