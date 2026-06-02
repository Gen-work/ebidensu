# Changelog

Tracks iterations across Misaki's browser (work) ↔ IDE (home) workflow.
Bump the date heading whenever a new bundle is delivered.

## 2026-06-02 - Incremental mapping add + Excel_Prefix auto-capture

### Added
- **Incremental mapping add (`-Add`).** `Generate-HostOpenMapping.ps1 -Add`
  merges freshly-selected rows INTO an existing `mapping_<Owner>.csv` instead
  of overwriting it. Existing rows — and ALL their progress (snaps, isReplaced,
  isMarked, isReviewed, Excel_Prefix, delivery flags, comments) — are kept
  verbatim; only `Correl_ID_M` values not already present are appended. Use it
  to grow the map day by day:
  `-Add -JobNames CJODJDEU,CJODJDB5` / `-Add -CorrelIdsM JIDSC09M` /
  `-Add -ExcelNames LJRVWD64` / `-Add -WbsStartRow 2300 -WbsEndRow 2400`.
  Reachable from the menu (`add` toggle) and the `Mapping` phase wiring.
- **`-ExcelNames` selector** for Generate/VerifyTool — accepts Excel_NAME(s)
  and reverse-maps each to its JOB_NAME (index-5 `W`→`J`) to look the row up
  in GFIX. Menu key `ex`.
- **`Get-PrefixFromFilename`** in `WorkbookResolver.ps1` — inverse of
  `Get-ExcelFullStem`; recovers the J4 prefix that precedes `_<Excel_NAME>`.

### Fixed
- **JenkinsSnap keeps window ops on Edge (folded in from codex's fix).** After
  `Read-Host`, some terminals keep the console as the foreground window, so the
  old code resized/screenshotted the **CLI** instead of Edge. JenkinsSnap now
  resolves the real `msedge` MainWindowHandle (`Activate-JenkinsEdgeWindow`),
  drops the `Switch-ToEdge` Alt+Tab dance, and screenshots via the activated
  `$edgeHwnd` rather than `GetForegroundWindow()`. Returns a zero handle (and
  fails the row) instead of operating on the wrong window. New `.ps1` files
  normalized to no-BOM per encoding policy.

### Changed
- **Clone now captures `Excel_Prefix`.** When a workbook is cloned from
  `-SourceDir` (real filename `<prefix>_<Excel_NAME>.xlsx`), Clone extracts the
  prefix and writes it back into the mapping (per `Excel_NAME`, atomic write).
  This is the missing "input point" for `Excel_Prefix` — downstream phases
  (CheckSheet / DeliverMail / DeliverFiles) now get the exact J4 filename
  without any manual entry. Existing rows/progress are untouched.

## 2026-06-01 - Delivery: review check sheet fill + review-request mail

The final hand-off step. Two new phases close the workflow after review.

### Added
- **DeliverMail phase.** One Outlook **draft** per evidence Excel (grouped by
  `Excel_NAME`), built via Outlook COM `CreateItem` + `Display` — never
  auto-sent. Misaki eyeballs each draft, clicks Send by hand, then presses Enter
  in the shell to set the new `isDelivered` mapping flag (`1` = sent). `s` skips,
  `q` quits, and `-m "comment"` records a note in the new `DeliverComment` column
  (per `Excel_NAME`, like `ReviewComment`). Subject =
  `【GIFT廃止対応】<Phase>レビュー依頼(<Excel_NAME>)`; body + reviewer + UNC paths
  are all config-driven (`Mail` / `Reviewer` in VerifyConfig.psd1). Outlook is
  released but never Quit (it may be the operator's live session).
- **CheckSheet phase.** Appends one row per Excel to the shared review check
  sheet (sheet `Check Sheet_J4`): A No. (continued, only if blank), B 記入日
  (today, format copied from the row above), C `JAVA`, E `J4内部ﾚﾋﾞｭｰ`,
  F full evidence filename, G owner, H reviewer (加瀬). Because it is a public
  document the edit is **double-checked**: a TEMP copy is filled and opened for
  visual review; on Enter the original is re-stat'd and, only if it is unchanged
  since the preview began, the identical edits are committed — otherwise the
  write is **held** so the operator can re-check. Already-listed Excels are
  skipped unless `-Force`. Path comes from `CheckSheet.Path`; if missing the
  phase prompts (and remembers it in `verify_session.json`).

### Changed
- **MappingStore** now defaults two new columns: `isDelivered` (`0`) and
  `DeliverComment` (`''`). `isDelivered` is a `PhaseOrder` field, so it is
  auto-added to existing mappings on startup and shown in Status.

## 2026-06-01 - Review workflow: header-fill fix, MarkGfix merge, review comments, mode sheet

Host->Open focus. Five related improvements:

### Fixed
- **Replace no longer drops the N1:U1 header fill.** `Reset-SheetBelowRow`
  (ExcelHelpers.ps1) used `Range.Clear()` on `A<start>:T<last>`, which wipes the
  *entire* merged cell when a colored header banner merged above the start row
  dips into the clear range. It now snapshots the fill+value of any such
  above-anchored merge on the boundary row, clears, then restores (re-merging
  defensively). Shared by all three Replace modes, so the fix covers
  ReplaceGift/Gfix/Df uniformly.

### Changed
- **GFIX-log highlight folded into MarkGfix.** The yellow "Command:" highlight is
  now applied in the same pass that draws GFIX red rectangles (one workbook open,
  tracked by `isMarked` bit 2). The standalone `isGfixLogMarked` column and the
  `MarkGfixLog` phase entry are removed. Core logic moved to the shared
  `Invoke-GfixLogHighlight` in ExcelHelpers.ps1; `MarkGfixLog.ps1` stays as a
  by-name re-highlight utility (no mapping column).
- **Review opens the matching sheet.** ReviewGift/Gfix/Df now bring
  GIFT/GFIX jushin kekka (or the DF compare sheet) to the front on open.

### Added
- **Review comments.** At the per-workbook Review prompt, append `-m "comment"`
  (works with Enter / `s` / `q`) to record a note in the new `ReviewComment`
  column (per Excel_NAME group). Prior notes are shown when the workbook opens.
  New read-only **Comments** phase lists every recorded note.

## 2026-05-31 - Cleanup pass: paste residue, SnapConfig BOM, encoding policy, JenkinsSnap DryRun

Follow-up to the review of the shared-lib refactor. The big blockers
(GfixLogDownload mapping-truncation + failure recovery, DfSnap region capture,
plan-driven Replace) were already fixed in the prior PR; this pass closes the
gaps that survived.

### Fixed
- **Pack-LlmContext paste residue** removed from 6 files whose tails carried a
  truncated `` ` `` / ``` ``n ``` / `--- File: X ---` separator that breaks the
  PowerShell parser: `SnapConfig.psd1`, `Run-Snap.ps1`, `MqSnap.ps1`,
  `HmSnap.ps1`, `Crop-Snap.ps1`, `ExcelSnap.ps1`.
- **`SnapConfig.psd1`** holds raw Japanese labels but had NO BOM, so
  `Import-PowerShellDataFile` would mojibake them -> added the UTF-8 BOM
  (per the .psd1 rule in CLAUDE.md).
- **`Fix-Encoding.ps1`** rewritten: it used to add a BOM to *every* file,
  which violated the ".ps1 = no BOM" policy. It now enforces the documented
  policy (.ps1 -> strip BOM, warn on non-ASCII; .psd1 -> BOM only when it holds
  raw text; .json/.jsonl -> no BOM) and uses robust extension filtering instead
  of the fragile `-Include` (no `-Recurse`).
- **`JenkinsSnap.ps1 -DryRun`** no longer opens Edge or prompts for navigation;
  it now reports the per-row capture plan and skips all UI.

### Added
- **`Check-Encoding.ps1`** now also fails on Pack-LlmContext paste residue
  (`--- File:` separators / stray ``` ``n ``` markers) in any business script,
  so a truncated paste is caught in CI instead of at runtime.

## 2026-05-29 - Shared MappingStore + plan-driven Replace + recovery/monitoring

Large refactor toward stability, recoverability, encoding safety, and a
review-standard evidence layout. New pure (COM-free) libraries are unit-tested;
the rest validated by static analysis (no PowerShell/Excel in the build env).

### Added - shared libraries (dot-source, ASCII source, no BOM)
- `MappingStore.ps1` - single source of truth for mapping_<Owner>.csv:
  `Import-Mapping`, `Export-MappingAtomic` (temp file + Move with retry +
  backoff; never silent on a locked CSV), `Ensure-MappingColumns`,
  `ConvertTo-TargetIdList` (array / comma / trim), `Test-TargetRow`,
  `Get-PendingRows` (snap + bitmask), `Set-MappingBit`.
- `GfixLog.ps1` - pure GFIX receive-log matcher. `SS_CODE = Substring(4,1)`;
  fragment `/appl/<TO>/<TO>Ver1/gfix/recv/<id> <SS>`; 0 match = fail; many =
  newest by timestamp + warning; returns the whole file's lines.
- `EvidencePlan.ps1` - pure correl-major plan builders (`Build-Gift/Gfix/Df
  EvidencePlan`) encoding the review order; Jenkins / NoGfix are trailing
  sections; NoGfix pictures optional; `Select-ValidCorrelIds` drops #VALUE!/blank.
- `EvidenceExecutor.ps1` - walks a plan and performs the Excel inserts;
  reports MissingRequired / MissingOptional / Warnings.
- `ProgressLog.ps1` - append-only `status\progress.jsonl` (UTF-8 no BOM).
- `ProjectLabels.ps1` - Japanese sheet/label names built from `[char]` so
  consumers stay ASCII / codepage-agnostic.
- `ScreenRegion.ps1` - pure screen-region clamp math.
- `AlignCompare.ps1` - pure sheet compare + migration-type classification.
- `Tests\` - `Run-Tests.ps1` (parse-checks every .ps1 + runs units) and units
  for MappingStore / GfixLog / EvidencePlan / ScreenRegion / AlignCompare.
- `Check-Encoding.ps1` - enforces the no-BOM (.ps1) / BOM-required-Japanese
  (.psd1) policy and prints constructed Japanese labels.
- `Watch-MappingProgress.ps1` - read-only monitor; copies the mapping and tails
  progress.jsonl, so it never locks mapping_<Owner>.csv.
- `Align.ps1` - Align/Precheck phase: compare work evidence vs the J4 baseline
  workbook by Excel_NAME; DryRun report by default, `-Apply` syncs values
  (formats TODO). Migration-type branching (Host->Open = 3 receive sheets;
  Open->Open / Open->Host add the GIFT/GFIX send-result sheets).

### Changed
- `Generate-HostOpenMapping.ps1` - removed pasted garbage at EOF (would break
  the file); owner-match arrows built via `[char]` (raw arrows silently broke
  matching under no-BOM PS 5.1); added `isMarked` / `isGfixLogMarked` columns;
  on regenerate (`-Force`) old completion status is merged by Correl_ID_M so
  finished work is not wiped; added owner / GFIX-row match counts.
- `JenkinsSnap.ps1` - routes through MappingStore (atomic per-row write instead
  of full-CSV re-read/re-write); emits progress events. Stable Edge/URL/
  screenshot logic and operator prompts kept.
- `DfSnap.ps1` - `region` capture default (df.exe window handle is flaky); waits
  for MainWindowHandle; window mode validates the rect and falls back to region
  (not fullscreen); per-direction crop for the asymmetric shadow.
- `ReplaceEvidence.ps1` - rewritten plan-driven (correl-major), replacing the
  folder-major loop. Gift/Gfix/Df builders + executor; GFIX log via the matcher;
  `isReplaced |= bit` only when all REQUIRED pieces inserted; NoGfix optional
  (`-AllowMissingOptionalNoGfix`); per-correl misses -> progress.jsonl.
- `GfixLogDownload.ps1` - per-correl transaction; on a missing log it does NOT
  blind-navigate: interactive pauses for manual download/retry/skip,
  `-NonInteractive` stops the run. Log naming
  `<id>_<timestamp>_<orig>.log`; GFIX_log set only when a log is really found.
- `VerifyConfig.psd1` / `VerifyTool.ps1` - Df region/crop config; new `Align`
  and `WatchProgress` phases, scripts, aliases; Df region args passed through.

### Notes / open items
- Snap output paths are phase-only (`snap\<type>\<Correl_ID_S>.png`); no
  biz_code/to_code subfolders (TO_code still groups Jenkins navigation only).
- Align full branching needs two domain inputs: the FROM_sys/TO_sys literals
  meaning "Host" (`-HostSystemTypes`) and confirmation of per-type sheet sets.
  Until set, type is Unknown and Align uses the Host->Open scope with a warning.
- Replace partial state is tracked by row bits + progress.jsonl (no new columns).

## 2026-05-19 - ReviewEvidence live test + Apply-LlmPatch v3

### Added / Changed
- Added and wired `ReviewEvidence.ps1` manual review flow.
  - New review phases: `ReviewGift`, `ReviewGfix`, `ReviewDf`, and `ReviewEvidence`.
  - `isReviewed` now follows bitmask semantics: `1=GIFT`, `2=GFIX`, `4=DF`, `7=all`.
  - `ReviewEvidence.ps1` opens evidence workbooks by `Excel_NAME`, waits for manual review, then saves/closes and updates mapping.
- Updated `VerifyConfig.psd1` with review phase entries, aliases, `SaveWaitMs`, and review bit values.
- Updated `VerifyTool.ps1` so review phases route to `ReviewEvidence.ps1` with `ReviewBit`.
- Fixed `ReviewEvidence.ps1` `$pid` bug by avoiding PowerShell's built-in read-only `$PID` variable.
- Explicitly opened workbooks as read-write via `Workbooks.Open($file, 0, $false)` and added read-only detection.
- Moved cursor relocation to after manual review instead of immediately after opening.
  - After manual review, the script selects `A3` by default, or `A1` for sheets whose name contains the arrow marker.
  - It also scrolls the active window to the top-left area after selecting the target cell.
- Clarified review prompt wording: `update mapping` is used instead of `mark`, to avoid confusion with the red-rectangle Mark phase.

### Tooling
- Replaced `Apply-LlmPatch.ps1` with v3 behavior.
  - Resolves relative patch paths against the script/repo directory.
  - Trims one boundary newline inside `<search>` / `<replace>`.
  - Validates multiple patches against the in-memory updated file state.
  - Writes each changed file once, preventing later patches from overwriting earlier patches on the same file.
  - Preserves BOM and original line endings.
- Confirmed that patch text copied with Windows PowerShell must use `Get-Content -Raw -Encoding UTF8 | Set-Clipboard` to avoid Japanese mojibake.

### Live test notes
- `ReviewGift -TargetIds QJRVWD50` successfully opened the workbook and updated `isReviewed` bit `1`.
- `ReviewGift -TargetIds KJRVWD64` successfully opened the workbook, waited for manual review, saved/closed, and updated mapping.
- Current remaining issue: after `Ctrl+S` + `Esc`, Excel may still report `$wb.Saved = false`, causing an extra confirmation prompt. Next change should remove the second confirmation and use a longer wait before close.

### Pending / Next
- Change save routine to:
  1. send `Ctrl+S`;
  2. wait about 300 ms;
  3. send `Esc` to dismiss the GenBa macro prompt;
  4. wait about 5 seconds for network-share save to settle;
  5. if `$wb.Saved` is still false, log a warning only;
  6. close without second save to avoid a macro loop;
  7. update mapping and continue to the next workbook.
- Keep `Close($false)` for now. Do not switch to `Close($true)` unless the macro behavior is retested, because close-time save may trigger the same macro/prompt again.
- After the save-close flow stabilizes, run a multi-workbook review pass without `-TargetIds`.

## 2026-05-18 (Session 3) — AI Co-workflow Tooling & Encoding Fixes

### Added / Changed
- **AI Sync Infrastructure**: Added `Pack-LlmContext.ps1`, `Apply-LlmPatch.ps1`, and `Export-DailyPatch.ps1` to facilitate secure, clipboard-driven LLM interactions bypassing Genba network restrictions.
- **Encoding Fixes**: Enforced `-Encoding UTF8` across all `Get-Content` and `Set-Content` calls in the LLM toolchain to prevent Shift-JIS text corruption (文字化け).
- **Docs**: Planned deprecation of `HANDOFF.md`. Updated `CLAUDE.md` to track local build tools.

## 2026-05-18 (later2) — v2.2: Mark phase, auto column repair, shape probe

### Added

- **`Mark.ps1`** — Phase `MarkGift` / `MarkGfix` / `MarkDf`. Draws red rectangles around designated areas of each picture in the evidence workbook. Idempotent: shapes whose name starts with `verifyMark_` are cleared before drawing, so re-runs are safe.
  - Reads each picture's `AlternativeText` metadata (stamped by `ReplaceEvidence.ps1`) to know what the picture represents (`v1|<folder>|<correl>`).
  - For each picture, looks up `VerifyConfig.psd1 -> Mark.Boxes -> <folder>` to find the list of `(OffsetX, OffsetY, Width, Height)` rectangles to draw, in points, relative to the picture's top-left.
  - `isMarked` bitmask field (1=Gift, 2=Gfix, 4=Df) tracks completion. Bit only set when at least one mark was drawn and nothing failed.
  - All marks ZOrder = `msoBringToFront` so they sit on top of the pictures.
  - Line color = red BGR `0x0000FF`, line weight default `1.5` (override per-box or via `Mark.LineWeight`).

- **`Probe-Shapes.ps1`** — Calibration tool. Lists every Shape in a workbook with its name, type, `(Left, Top, Width, Height)`, and `AltText`. Color-codes output:
  - **green** = Picture with `v1|...` metadata (Mark anchors on these).
  - **yellow** = manual AutoShape (your reference rectangle — measure these).
  - Prints the calibration recipe at the end:
    ```
    OffsetX = AutoShape.Left - Picture.Left
    OffsetY = AutoShape.Top  - Picture.Top
    Width   = AutoShape.Width
    Height  = AutoShape.Height
    ```
  Usage: `.\VerifyTool.ps1 -Phase ProbeShapes -ProbeFile evidence\KJRVWD64.xlsx`

- **`ExcelHelpers.ps1`** additions:
  - `Set-ShapeMetadata` / `Get-ShapeMetadata` — store and read `"v1|<key>|<value>"` payload on a Shape's `AlternativeText`. Used to identify pictures across runs.
  - `Add-RedRectangle` — hollow red border rectangle at absolute `(Left, Top)`, configurable `Width`/`Height`/`LineWeight`/`Name`. Sets `msoBringToFront`.
  - `Remove-MarkShapes` — deletes every Shape on a sheet whose Name starts with a given prefix. Used by Mark for idempotent re-runs.

- **`ReplaceEvidence.ps1`** changed:
  - Every inserted Picture is now stamped with `Set-ShapeMetadata` immediately after insertion. ExcelSnap gets key `'excel'` and the `JOB_NAME` as value. Per-correl snaps get the source folder name and `Correl_ID_S` as value. NoGfix tail snaps stamped too. This is what makes Mark work.

- **`Ensure-PhaseColumns`** function in `VerifyTool.ps1`. Reads `PhaseOrder`, finds every non-empty `Field`, ensures the mapping CSV has that column (defaults missing cells to `'0'`). **Auto-called on every startup** so adding a new phase/field never causes `(value missing)` issues. Existing data is never modified.
  - Manual entry point: `.\VerifyTool.ps1 -Phase RepairMapping`.
  - Solves: "I shouldn't lose progress just because I changed the mapping module" — old mapping CSVs upgrade in place.

- **`VerifyConfig.psd1`** additions:
  - New `Mark` section with `NamePrefix`, `LineWeight`, and `Boxes` table per source folder. Default offsets are placeholders (all `OffsetX=0, OffsetY=0`); calibrate via Probe.
  - New `Scripts` entries: `Mark`, `Probe`.
  - New `PhaseOrder` entries: `MarkGift` (bit=1), `MarkGfix` (bit=2), `MarkDf` (bit=4), `RepairMapping`, `ProbeShapes`.
  - New aliases: `Mark` / `Mgift` / `Mgfix` / `Mdf` / `RepairMapping` / `Repair` / `EnsureCols` / `Probe`.

### How to calibrate Mark (first-time setup)

1. Run `ReplaceGift` once on a single Excel: `.\VerifyTool.ps1 -Phase ReplaceGift -TargetIds KJRVWD64`.
2. Open `evidence\KJRVWD64.xlsx` manually.
3. Switch to the `GIFT受信結果` sheet.
4. Draw a red rectangle by hand around the area you want Mark to circle on one of the pictures (the area depends on which picture's source folder — e.g. for `GIFT_HM` picture, circle the `処理状態` cell).
5. Save the workbook (leave it closed when you run probe).
6. `.\VerifyTool.ps1 -Phase ProbeShapes -ProbeFile <full path to that xlsx>` — note the Picture position and the AutoShape position.
7. Compute `OffsetX = AutoShape.Left - Picture.Left`, `OffsetY = AutoShape.Top - Picture.Top`, `Width = AutoShape.Width`, `Height = AutoShape.Height`.
8. Edit `VerifyConfig.psd1 -> Mark -> Boxes -> 'GIFT_HM'` (etc.) with the measured values.
9. Repeat per folder you care about (`GIFT_MQ`, `GIFT_Jenkins`, etc.).
10. Run `.\VerifyTool.ps1 -Phase MarkGift` — it will draw the same rectangle at the right place on every picture across every evidence workbook.

For folders you don't want marked (like `GIFT_noGfixfile` and `excel`), leave the list empty `@()`.

### Known limitations

- All pictures from the same source folder get the same rectangle. If the row position varies inside the picture (e.g. HM occasionally has the target row on row 2 instead of row 1), the rectangle will be off. OCR support is the planned escape hatch but not yet built — see `OcrLocate.ps1` in pending.
- `Mark` requires `ReplaceEvidence` to have run first (otherwise pictures have no metadata). If you re-run `ReplaceGift` after `MarkGift`, all marks are wiped by `Reset-SheetBelowRow`; just re-run `MarkGift` after.
- Multiple AutoShape candidates in the same Probe output (e.g. you drew 2 reference rectangles) are all listed — pick whichever is yours.

---

## 2026-05-18 (later) — v2.1: Validate phase + Claude Code context

### Added

- **`Validate.ps1`** — Phase `Validate` (aliases: `Check`, `Diagnose`). Read-only diagnostic. Scans `WorkDir` without touching Excel and reports:
  - Mapping file presence + required column check + `isReplaced` bitmask distribution (e.g. `none=4, Gift=2, all=1`).
  - Directory structure (`evidence`, all `snap/*` subfolders, `DATA`, `log`) with PNG/XLSX counts.
  - Template files (`template.xlsx`, `template_<bizcode>.xlsx`) present in WorkDir.
  - **Per Excel_NAME readiness matrix** — for each Excel_NAME, a Y/N per phase (Clone / Gift / Gfix / Df) plus a compact `missing` column showing what's blocking it (`evidence`, `excel`, `gHM`, `gMQ`, `gJk`, `GHM`, `GJk`, `DF`).
  - Aggregate readiness totals per phase (e.g. `ReplaceGift : 5/8 ready`).
  - Next suggested action with actual CLI commands to run.
  - `-Compact` suppresses the per-Excel matrix.
  - `-TargetIds` narrows the scan.

- **`CLAUDE.md`** — context file for Claude Code at home. Documents project purpose, file map, key conventions (dot-source rules, switch-flag pattern, encoding, bitmask), Excel COM rules, current state, planned phases, Misaki's preferences, and the cross-environment workflow pattern.

- **Clone next-step hints** — `Clone.ps1` now prints clear follow-up commands at the end. If anything was missed, it tells you which fallback locations were tried. If anything was cloned, it points to `Validate` and `ReplaceGift` as the next moves.

- **`VerifyConfig.psd1`** — `Scripts.Validate = 'Validate.ps1'`, `PhaseOrder` entry, aliases (`Validate`, `Check`, `Diagnose`).

### Why Validate matters

When entering a new work session, Misaki was previously guessing which phase to run next. Status only shows mapping CSV column states — it doesn't know whether the actual snap PNGs exist on disk. Validate fills the gap by cross-checking mapping rows against filesystem state in one pass, so the answer to "what should I run next?" is one command away.

---

## 2026-05-18 — v2: Clone + Replace + bitmask

### Added

- **`ExcelHelpers.ps1`** — dot-source library for Excel COM operations.
  - `New-ExcelApp` / `Close-ExcelApp` — lifecycle, ScreenUpdating off, alerts off.
  - `Open-Workbook` / `Close-Workbook` / `Get-SheetByName` / `Unhide-AllSheets`.
  - `Reset-SheetBelowRow` — deletes shapes whose top ≥ row's top, then clears values + formats + highlights in the range `A<startRow>:T<lastUsed>`. Uses `End(xlUp)` from the bottom of column A to find last used row (fixes the `xlDown` bug in the old script).
  - `Get-RowAtOrBelow` / `Get-NextAnchorRow` — anchor-row math using `shape.Top + shape.Height` against `Cells(r,1).Top`. Approximates start scan row from `shape.Top / 15` for speed.
  - `Insert-PictureSendToBack` — inserts picture at native size with z-order = `msoSendToBack` (1), so later Mark rectangles will stay on top.
  - `Write-PlainText` — plain label without bold / color / highlight.
  - `Write-LogLines` — paste multi-line log content into stacked cells.
  - `Get-BitValue` / `Set-BitValue` — read / `-bor` write on integer columns.
  - `Ensure-Column` — adds a missing column to every row before `Export-Csv` so it persists.

- **`Clone.ps1`** — Phase `Clone` (aliases: `MkExcel`, `RenameExcel`).
  - Groups mapping by `Excel_NAME`.
  - Per group, tries source paths in priority order:
    1. `<SourceDir>\<bizcode>\<Excel_NAME>.xlsx` for each bizcode candidate.
    2. `<WorkDir>\template_<bizcode>.xlsx`.
    3. `<WorkDir>\template.xlsx` (universal fallback).
  - Bizcode candidates come from `-BizCodes` if given, otherwise from row's `TO_code` then `FROM_code` (deduped).
  - Copies to `<WorkDir>\evidence\<Excel_NAME>.xlsx`.
  - Skip if dest exists unless `-Force`.
  - Prompts for `SourceDir` interactively if neither CLI flag nor session has it.

- **`ReplaceEvidence.ps1`** — Phases `ReplaceGift`, `ReplaceGfix`, `ReplaceDf`.
  - Single script, `-Mode Gift|Gfix|Df` switches behavior.
  - Per `Excel_NAME` group: open workbook → find sheet by mode → `Reset-SheetBelowRow 3` → insert ExcelSnap at B3 (Gift/Gfix only) → stacked per-correl snaps (`GIFT_HM`/`GIFT_MQ`/`GIFT_Jenkins` or `GFIX_HM`/`GFIX_Jenkins` or `DF`) → tail per mode → save.
  - Tail:
    - **Gift** — label `GFIX Jenkins フォルダ受信ファイルなし` once, then `GIFT_noGfixfile\<correl>.png` snaps stacked. NoGfix snaps may legitimately be absent; absence is logged as info, not a failure.
    - **Gfix** — per-correl label `GFIX受信log` + placeholder text. Stub function `Get-GfixLogLines` returns `<<TODO: GFIX 受信 log for <correl>>>` for now; replace with real grep against `work\log\` after `GfixLodDownload` lands (cf. old `Replace-GFIX.ps1` — `Find-RecvLogFile` / `Find-SendLogFile` / `Extract-SystemOutBlock`).
    - **Df** — nothing extra.
  - **All-OK contract** — `isReplaced |= bit` only if every required image existed and was inserted without error. Missing required images mark the run as failed for that group and the bit stays clear.

- **`isReplaced` bitmask field** in `mapping_<owner>.csv`.
  - `bit 0 (1)` = Gift done, `bit 1 (2)` = Gfix done, `bit 2 (4)` = Df done. Total 7 = all done.
  - One column instead of three flag columns, avoids schema bloat.
  - `Ensure-Column` adds the column to every row on first run so `Export-Csv` writes it.

- **`Get-FieldStats` bitmask support** in `VerifyTool.ps1`. New optional `BitValue` parameter. When `> 0`, "done" means `(value -band BitValue) -eq BitValue` instead of `eq '1'`.

- **`PhaseOrder` `BitValue` entries** in `VerifyConfig.psd1`. `Show-Status` and `Get-RecommendPhase` both read it. Status display shows `bit=N` next to the field name.

- **`CloneSourceDir` session persistence**. `VerifyTool.ps1` now has `-CloneSourceDir <path>` and `-BizCodes A,B` parameters; `CloneSourceDir` is saved into `verify_session.json` so subsequent runs reuse it without re-typing.

- **`Ask-RunOptions`** in `VerifyTool.ps1` gained two more interactive options: `d` to set `CloneSourceDir`, `b` to set `BizCodes`.

- **Aliases** in `VerifyConfig.psd1`:
  - `Clone` / `MkExcel` / `RenameExcel` → `Clone`
  - `Replace` / `ReplaceEvidence` → `ReplaceGift` (default; explicit per-mode names below)
  - `ReplaceGift` / `Rgift` → `ReplaceGift`
  - `ReplaceGfix` / `Rgfix` → `ReplaceGfix`
  - `ReplaceDf` / `Rdf` → `ReplaceDf`

- **`Replace` section** in `VerifyConfig.psd1`:
  - `StartRow = 3`, `ColAnchor = 2`, `BlankRowsBetween = 1`, `ClearEndColumn = 20`.
  - `GiftNoGfixLabel`, `GfixLogLabel`, `GfixLogTodoText`. Each can be edited without touching code. Empty values fall back to the in-code `[char]0xNNNN` defaults.

### Changed

- `Get-FieldStats` signature: added `[int]$BitValue = 0`. Default behavior (no `BitValue`) is unchanged — still `eq '1'`.
- `Show-Status` reads `BitValue` from each phase entry and prints `bit=N` next to the field name when present.
- `Get-RecommendPhase` likewise honors `BitValue` when picking the next pending phase.
- Old planned-only handler for `ReplaceEvidence` removed; `ReplaceEvidence` alias now points to `ReplaceGift` so old muscle memory still works.
- `Show-VerifyHelp` gained Clone / Replace examples and the bitmask legend.

### Conventions kept

- Switch flags resolved to `[bool]$X.IsPresent` before any dot-source, to avoid the `$Force` overwrite bug pattern.
- Tool scripts with `param()` are never dot-sourced; they are called via `& $p @args`. `ExcelHelpers.ps1` has no `param()` so dot-sourcing it from `ReplaceEvidence.ps1` is safe.
- Japanese text in `.ps1` files uses `[char]0xNNNN` for sheet names. Labels live in `VerifyConfig.psd1` as literal Japanese strings (UTF-8 BOM PSD1).
- English-only console output to avoid encoding issues in Windows shells.

### Pending / next iteration

- `Mark.ps1` — red rectangles, highlights overlaid on inserted pictures. Will dot-source `ExcelHelpers.ps1`.
- `TimeRange.ps1`.
- `OcrLocate.ps1`.
- `GfixLodDownload` real implementation (drives the `Get-GfixLogLines` stub in `ReplaceEvidence.ps1`).
- `DfSnap` real implementation.
- After `GfixLodDownload` lands: replace `Get-GfixLogLines` stub with grep against `work\log\` for recv/send log lines per `Correl_ID_S`. Reference old `Replace-GFIX.ps1` helpers (`Find-RecvLogFile`, `Find-SendLogFile`, `Extract-SystemOutBlock`).

## How to use this changelog

When iterating between browser (work) and IDE (home):

1. After making changes on either side, add an entry to this file under a new date heading describing what changed and why.
2. Sync the whole `VerifyTool` directory across.
3. The other side reads this file first to know what's new.

Suggested entry skeleton:

```markdown
## YYYY-MM-DD — short title

### Added / Changed / Fixed / Pending
- ...
```
