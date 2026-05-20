# CLAUDE.md

You are working in **VerifyTool** — PowerShell automation for Misaki's GIFT→GFIX migration evidence collection at Honda Japan.

Read this file before doing anything else. Always read `CHANGELOG.md` next.

## What this tool does

Misaki manually verifies file-transfer migration evidence for ~hundreds of correlation IDs. The tool automates:

1. Generating a mapping CSV from the master `wipGFIX一覧.xlsx`.
2. Capturing Edge-browser screenshots from Jenkins / HostMonitor (HM) / MQ admin console for each Correl_ID_S.
3. Downloading transferred files from Jenkins folders into `work/DATA/`.
4. Cloning evidence Excel templates into `work/evidence/<Excel_NAME>.xlsx`.
5. Replacing the body of each evidence book with the captured snapshots (GIFT receive, GFIX receive, GIFT-vs-GFIX diff).
6. Visual review pass where Misaki opens each finished evidence book and stamps it reviewed.

Migration deadline: 2026-12 (MQFX EOS). Real, billable work — be careful.

# LLM output rules
---------------------------------------------------------
Never output a whole file unless explicitly asked.

Use one of these two patch formats only:

## A. XML search/replace patch

Default for small edits, docs, non-git files, or fragile exact replacements.

```xml
&lt;patch file="path/to/file"&gt;
&lt;search&gt;
exact old text, with enough context to be unique
&lt;/search&gt;
&lt;replace&gt;
new text
&lt;/replace&gt;
&lt;/patch&gt;
Rules:

Output patch blocks only. No prose inside the patch payload.

<search> must match exactly one place after LF normalization.

Include 1-2 context lines when useful.

For multiple edits, output multiple <patch> blocks.

If the patch body itself contains patch-like tags, escape them in the LLM answer as &lt;search&gt;, &lt;/search&gt;, &lt;replace&gt;, etc. Apply-LlmPatch.ps1 decodes entities before matching/writing.

Prefer XML patch when the target is not tracked by Git.

B. Git unified diff
Preferred for larger code edits inside this Git repo.

Rules:

Output raw unified diff from repo root, starting with diff --git ....

No prose inside the diff.

A single fenced ```diff block is acceptable, but raw diff is safer.

Apply-LlmPatch.ps1 runs git apply --check before git apply.

Git diff mode does not create PatchBak; use git diff / git checkout -- file for review and recovery.

Decision:

Small exact edit -> XML.

Big code change in repo -> Git diff.

Unsure -> XML, because failure is easier to diagnose.

Added Git unified-diff mode: clipboard diff --git -> git apply --check -> git apply.

XML patch mode now decodes entities before matching/writing, so docs/code can contain patch-like tags safely.

XML backups now live under PatchBak/.

Added write preflight and temp-file write. No full rollback transaction; keep it simple and Git-backed.

Updated CLAUDE.md and README.md patch output rules.
2026-05-19 — ReviewEvidence live test + Apply-LlmPatch v3

全コンテキストの抽出（初期設定・大規模型再構築用）

.\Pack-LlmContext.ps1

AIからのXMLパッチ自動適用

.\Apply-LlmPatch.ps1

退勤前の差分抽出（自宅のCursor/Claude同期用）

.\Export-DailyPatch.ps1

コンテキストをAIへ渡す

.\Pack-LlmContext.ps1

初期相談、大きめの設計変更、別環境のAIへ状況を渡すときに使います。

AIの修正を適用する

.\Apply-LlmPatch.ps1 -DryRun
.\Apply-LlmPatch.ps1

Apply-LlmPatch.ps1 はクリップボードを読み、形式を自動判定します。

XML patch: 小さい修正、README/CLAUDE/設定ファイル、Git外ファイル向け。

Git diff: repo内の大きめのコード修正向け。先に git apply --check を実行します。
XML patch のバックアップは PatchBak/ に作られます。repo直下を .bak で汚しません。

差分を確認・同期する

git diff
.\Export-DailyPatch.ps1

AI修正後は必ず git diff を見てから commit / 同期してください。

## Entry point

```powershell
.\VerifyTool.ps1                 # interactive menu
.\VerifyTool.ps1 -Help           # CLI examples
.\VerifyTool.ps1 -Phase Status   # quick mapping status
.\VerifyTool.ps1 -Phase Validate # read-only readiness report (use this first)
```

`VerifyTool.ps1` loads `VerifyConfig.psd1`, reads `verify_session.json` for remembered settings, and routes to per-phase scripts via `Invoke-ToolPhase`.

## Project file map

# --- AI Co-workflow (Local Build Tools) ---
Pack-LlmContext.ps1     packs project context to clipboard for LLM ingestion
Apply-LlmPatch.ps1      parses XML patches from clipboard and updates local files
Export-DailyPatch.ps1   extracts daily Git diff to clipboard for low-bandwidth sync
# ------------------------------------------

```
VerifyTool.ps1          main entry, menu, phase router, status display
VerifyConfig.psd1       paths, scripts, phase order, aliases, labels
verify_session.json     remembered settings (auto-generated, gitignore)

ExcelHelpers.ps1        dot-source library: Excel COM + bitmask helpers
Clone.ps1               Phase Clone (mkexcel / renameexcel)
ReplaceEvidence.ps1     Phase ReplaceGift | ReplaceGfix | ReplaceDf
Validate.ps1            Phase Validate (read-only diagnostic)
JenkinsSnap.ps1         Phase GiftJenkins | GfixJenkins | GiftJenkinsNoFile

(Invoked but not in this repo:
 HmSnap.ps1, MqSnap.ps1, ExcelSnap.ps1, ReviewEvidence.ps1,
 Common.ps1, Generate-HostOpenMapping.ps1, Crop-Snap.ps1)

CHANGELOG.md            iteration log (bump on every change)
README.md               user-facing docs
HANDOFF.md              older architecture notes
```

The `WorkDir` (outside this repo, on `\\fs-f3170-1\...`) contains:

```
work/
  mapping_<owner>.csv     primary data: 1 row per Correl_ID_S
  evidence/<Excel_NAME>.xlsx
  snap/excel/<JOB_NAME>.png
  snap/{GIFT_HM,GIFT_MQ,GIFT_Jenkins,GIFT_noGfixfile,GFIX_HM,GFIX_Jenkins,DF}/<Correl_ID_S>.png
  DATA/{GIFT,GFIX}/<Correl_ID_S>/  <-- Jenkins-downloaded files
  template.xlsx                    <-- universal evidence template
  template_<bizcode>.xlsx          <-- per-bizcode template
```

## Key conventions — do not break

### Dot-source vs invocation

Scripts with `param()` are **never dot-sourced**. They are always called via `& $path @args`. Otherwise the parent script's `[switch]` parameters leak into the child (the `$Force` overwrite bug — see HANDOFF.md).

`ExcelHelpers.ps1` is intentionally written **without** `param()` so dot-sourcing it from other tool scripts is safe. Do not add a `param()` block.

### Switch flags

In every script with switches, immediately convert to a plain bool before any dot-source:

```powershell
$forceFlag = [bool]$Force.IsPresent
```

Use `$forceFlag` downstream. Never read `$Force` after a dot-source.

### Encoding

| File type | Encoding |
|-----------|----------|
| `.ps1`, `.psd1` | UTF-8 BOM + CRLF |
| `.md` | UTF-8 no BOM, LF |
| Console output | English only |
| Japanese in `.ps1` code | `[char]0xNNNN` per character |
| Japanese in `.psd1` config | literal string (BOM PSD1 handles it) |

Example:

```powershell
$sheetGiftRecv = "GIFT" + [char]0x53D7 + [char]0x4FE1 + [char]0x7D50 + [char]0x679C  # GIFT受信結果
$Owner = [char]0x53B3                                                                 # 厳
```

### isReplaced bitmask

One integer column in `mapping_<owner>.csv`. Each mode owns one bit:

```
bit 0 (1) = GIFT replace done
bit 1 (2) = GFIX replace done
bit 2 (4) = DF replace done
total 7   = all three done
```

Use `Set-BitValue` / `Get-BitValue` / `Ensure-Column` from `ExcelHelpers.ps1`. Never write `'1'` to `isReplaced` directly.

`Get-FieldStats` in `VerifyTool.ps1` supports two modes:
- No `BitValue` arg → legacy `eq '1'` check.
- `BitValue > 0` → bitmask check `(value -band BitValue) -eq BitValue`.

### Phase config

A `PhaseOrder` entry in `VerifyConfig.psd1` can carry:

- `Key` — phase name (used by `-Phase` CLI arg).
- `Field` — mapping column. Empty `''` = no status tracking (Mapping / Status / Validate / Clone / Crop).
- `BitValue` — optional. When present, the field is read as a bitmask.
- `Label` — Japanese, shown in menu and help.
- `Status` — `implemented` | `legacy` | `planned`. `planned` calls `Show-PlannedPhase`.

To add a new phase:
1. Add a `PhaseOrder` entry.
2. Add an `Aliases` entry if it has shortcuts.
3. Add an `Invoke-ToolPhase` branch in `VerifyTool.ps1`.
4. Add a `Scripts` entry if it has its own `.ps1` file.

## Excel COM rules

1. Always use `New-ExcelApp` / `Close-ExcelApp` from `ExcelHelpers.ps1`. Close in `finally`.
2. After `AddPicture`, the returned Shape's `Top` and `Height` are reliable. Use `Get-NextAnchorRow` to compute the next stacking row — do not do manual row arithmetic.
3. `ZOrder(1)` = `msoSendToBack`. All inserted images go to the back so future `Mark.ps1` rectangles stay visible.
4. `Reset-SheetBelowRow $ws 3` is destructive but bounded — it deletes shapes whose top ≥ row 3's top, then clears values + formats + highlights in `A3:T<lastUsed>`. Row 3 is the conventional anchor; never reset below row 1 or 2.
5. Sheet names by mode (use `[char]` codes in code):
   - Gift → `GIFT受信結果`
   - Gfix → `GFIX受信結果`
   - Df → `GIFTデータvsGFIXデータ`

## When making changes

1. Read `CHANGELOG.md` to see what shifted recently.
2. Run `Validate` mentally — does the change make a phase ready or break readiness?
3. Don't dot-source `param()` scripts.
4. Don't write `'1'` to `isReplaced`; use bitmask helpers.
5. Don't reproduce the row-finding bug — use `Get-NextAnchorRow`.
6. Don't refactor unless asked. Don't expand scope.
7. When proposing edits, give line numbers and minimal diffs, not full file rewrites.
8. After any code change, add a `CHANGELOG.md` entry under a new date heading.

## Current state (last bump: 2026-05-19 v2.3)

**Implemented:** Mapping, ExcelSnap (legacy), GiftHmSnap, GiftMqSnap, GiftJenkins, GiftJenkinsNoFile, GfixHmSnap, GfixJenkins, Clone, ReplaceGift, ReplaceGfix (with TODO stub for log), ReplaceDf, MarkGift, MarkGfix, MarkDf, ReviewGift, ReviewGfix, ReviewDf, ReviewEvidence, Crop, Validate, RepairMapping (auto on startup + manual), ProbeShapes, Status.

**Mark architecture (current):**
- `ReplaceEvidence` stamps every inserted Picture with `AlternativeText = "v1|<folder>|<correl>"` (e.g. `"v1|GIFT_HM|JIDSC48S"`).
- `Mark.ps1` walks each evidence workbook's sheet, reads each Shape's metadata, and for every Picture whose folder is in the current mode's `Folders` list, draws rectangles per `VerifyConfig.psd1 -> Mark.Boxes -> <folder>`.
- Offsets are calibrated by running `Probe-Shapes.ps1` against a manually-marked reference Excel — see CHANGELOG.md for the recipe.
- Marks are idempotent: all shapes whose name starts with `verifyMark_` are deleted before redrawing.

**Auto column repair:**
- On every VerifyTool startup, if the mapping CSV exists, `Ensure-PhaseColumns` reads `PhaseOrder` and adds any missing column with default `'0'`. Existing data untouched.
- Manual: `.\VerifyTool.ps1 -Phase RepairMapping`.

**ReviewEvidence architecture (current live-tested state):**
- Review phases use `isReviewed` as a bitmask: `1=GIFT`, `2=GFIX`, `4=DF`, `7=all`.
- `ReviewGift`, `ReviewGfix`, `ReviewDf`, and `ReviewEvidence` all invoke `ReviewEvidence.ps1` with different `ReviewBit` values.
- `ReviewEvidence.ps1` opens evidence workbooks from `work/evidence/<Excel_NAME>.xlsx` in mapping order and waits for Misaki to review/adjust by hand.
- After Misaki returns to the shell and presses Enter, the script performs the final cursor relocation, sends save keys, closes the workbook, and updates mapping.
- Cursor relocation must happen after manual review, not immediately after opening.
- The script explicitly requests read-write open via `Workbooks.Open($file, 0, $false)` and must not update mapping if the workbook opens read-only.
- Avoid using `$pid` as a variable name in PowerShell; it collides with the built-in read-only `$PID`.
- Prompt wording should say `update mapping`, not `mark`, because `Mark` already means the red-rectangle phase.
- Current pending improvement: save flow should be `Ctrl+S -> short wait -> Esc -> about 5s wait -> close -> update mapping`. Do not add a second blocking confirmation just because `$wb.Saved` remains false; the GenBa save macro can make `$wb.Saved` unreliable.
- Prefer keeping close as `Close($false)` after explicit save, to avoid triggering a second macro save loop at close time.

**Apply-LlmPatch current notes:**
- Use the v3 patcher behavior: resolve relative paths against the script/repo directory, trim boundary newlines inside XML blocks, validate sequential patches in memory, and write each changed file once.
- When loading a patch text file into clipboard on Windows PowerShell, use `Get-Content -Raw -Encoding UTF8 | Set-Clipboard`; otherwise Japanese text may become mojibake and `<search>` will fail.

**Planned (still TODO):**
- `GfixLodDownload` — accept dirty multiline URL paste, parse links, download to `work/DATA/GFIX_LOD`, update `GFIX_log`. Drives the `Get-GfixLogLines` stub in `ReplaceEvidence.ps1`.
- `DfSnap` — open `C:\tools\DF\DF.exe`, input 2 file paths, capture to `snap/DF/<correl>.png`, update `DF_snap`.
- `OcrLocate.ps1` — Windows.Media.Ocr based bounding-box locator. Fallback for when the target row inside a picture is variable (current Mark assumes fixed offset).
- `TimeRange.ps1`.

After `GfixLodDownload` lands, replace `Get-GfixLogLines` stub in `ReplaceEvidence.ps1` with grep against `work/log/` — reference old `Replace-GFIX.ps1` (`Find-RecvLogFile`, `Find-SendLogFile`, `Extract-SystemOutBlock`).

## Misaki's preferences

- Discussion: Chinese. Brackets `「」` nested `『』`.
- Code & comments: Japanese or English. Console output: English.
- Concise, atomic modules. Single-purpose scripts.
- Don't refactor without being asked. Don't expand scope.
- Provide line numbers when proposing edits, not full file rewrites.
- "Caveman mode" (say `caveman` or `terse`): strip articles, filler, hedging from Chinese prose. Code blocks, JCL, error messages, Japanese workplace drafts untouched.
- If she gives feedback labeled `fb` with numbered points, do targeted fixes only.

## Cross-environment workflow

Misaki bounces between:
- **Work** — browser-only access to Claude (claude.ai). Real test data on `\\fs-f3170-1\...`. No CLI install.
- **Home** — IDE + Claude Code. No test data.

Pattern:
1. Work iteration: she pastes problem into claude.ai, gets new files, drops them onto the network share, tests.
2. After work session: she bumps `CHANGELOG.md` and pushes the directory home (file copy / git / share).
3. Home iteration: `cd VerifyTool_v2 && claude`. Claude Code reads this file + `CHANGELOG.md` and picks up where work left off. Refactor / refine / write new modules with no Excel COM available (mock as needed).
4. Push home changes back to work share.

Always read `CHANGELOG.md` first to know what shifted.
