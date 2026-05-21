# Changelog

Tracks iterations across Misaki's browser (work) ↔ IDE (home) workflow.
Bump the date heading whenever a new bundle is delivered.

## 2026-05-19 — ReviewEvidence live test + Apply-LlmPatch v3

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
