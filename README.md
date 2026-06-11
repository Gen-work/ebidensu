# VerifyTool

`VerifyTool.ps1` is the main entry for the verification evidence workflow.

Daily use:

```powershell
.\VerifyTool.ps1
.\VerifyTool.ps1 -Help
```

## AI 協働ワークフロー

現場と自宅の環境間で安全にコードの差分を同期・更新するためのツールです。

1. 全コンテキストの抽出（初期設定・大規模型再構築用）
```
.\Pack-LlmContext.ps1
```

2. AIからのXMLパッチ自動適用
```
.\Apply-LlmPatch.ps1
```

3. 退勤前の差分抽出（自宅のCursor/Claude同期用）
```
.\Export-DailyPatch.ps1
```

## Per-work-folder config (verify_config.json)

`VerifyConfig.psd1` holds the project defaults. Each work folder may also carry
a `verify_config.json` that is **deep-merged over** those defaults at startup,
so a single case can customize almost everything without editing the shared
`.psd1`.

Precedence: **CLI args > work-folder `verify_config.json` > `VerifyConfig.psd1` > session fallback** (for the few values that still support session fallback).

Generate or refresh a work-folder file (pre-filled from the current effective
config, including existing overlay values plus any new defaults):

```powershell
.\VerifyTool.ps1 -Phase InitConfig
.\VerifyTool.ps1 -Phase InitConfig -Interactive  # grouped peek/edit/delete/save UI
.\VerifyTool.ps1 -Phase InitConfig -Force        # accepted for old habit; updates still keep a .bak
```

`InitConfig` also writes `verify_config.README.txt` next to the JSON with field
explanations, so the JSON can stay clean (standard JSON does not support
`//` comments). Interactive mode groups settings by `intro`, `phase`, `snap`,
`excel`, `wbs`, `path`, `mail`, and `all`; it lets you peek without changing,
edit any JSON path (for example `Window.Width` or `Mail.BodyLines`), delete a
path, then save only after typing `YES`. The `_README` introduction shown in
the JSON is included in the `intro` group and can be changed the same way. Then
edit values such as `DefaultOwner`, `Workbook.ExcelPrefix`, `Window`,
`Mark.Boxes` (red-rectangle positions), `Mail` (subject/body templates),
`Reviewer`, `CheckSheet`, `Df` (capture region) and `ExpectedTime`. Save as
UTF-8; Japanese text is fine. Re-run any phase; the banner prints
`Config overlay : ...` when it loaded.

See `verify_config.example.json` in the repo for a ready-to-copy starter.

## Folder layout

Expected work folder:

```text
work\
  mapping_<Owner>.csv
  wipGFIX一覧.xlsx
  template.xlsx                  (universal template, optional)
  template_<bizcode>.xlsx        (per-bizcode template, optional)
  snap\
    excel\
    GIFT_HM\
    GIFT_MQ\
    GIFT_Jenkins\
    GIFT_noGfixfile\
    GFIX_HM\
    GFIX_Jenkins\
    DF\
  DATA\
    GIFT\<Correl_ID_S>*
    GFIX\<Correl_ID_S>*
    GFIX_LOD\
  evidence\
    <Excel_NAME>.xlsx
```

### Full-width filename fallback

Some customer-provided files may use full-width ASCII characters in filenames
(for example `０` instead of `0`). Workbook lookup first tries the normal exact
and wildcard paths; if nothing is found, it scans for `.xlsx` names whose
full-width ASCII normalizes to the requested `Excel_NAME`, warns, and asks before
using the candidate. Non-interactive scripts/tests can use the resolver policy
`Prompt`, `Accept`, or `Reject`.

The reusable fallback lives in `WorkbookResolver.ps1`:

```powershell
# Generic file lookup after your normal not-found branch.
Resolve-FullWidthFileName -Dir $dir -Name 'report0.txt' -Filter '*.txt' -FullWidthFallback Prompt

# Workbook-specific lookup used by evidence phases.
Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem -FullWidthFallback Prompt
```

Jenkins-downloaded receive files are saved under `DATA\GIFT` or `DATA\GFIX` with their Jenkins filename, so downstream delivery can pick up `Correl_ID_S*` files. For example:

```text
C:\path\to\work\DATA\GIFT\JIDSK48S
```

`VerifyTool.ps1` remembers the last `WorkDir`, `Owner`, window size, crop size, evidence folder, review cursor cell, `CloneSourceDir`, and `J4BaseDir` in `verify_session.json`.

## Main commands

Status (CSV column scan):

```powershell
.\VerifyTool.ps1 -Phase Status
```

Validate (read-only readiness diagnostic -- run this first when entering a new work session):

```powershell
.\VerifyTool.ps1 -Phase Validate
.\VerifyTool.ps1 -Phase Validate -TargetIds JIGPL48S
```

Generate or refresh mapping:

```powershell
.\VerifyTool.ps1 -Phase Mapping -Force
```

Snap stages (existing):

```powershell
.\VerifyTool.ps1 -Phase ExcelSnap
.\VerifyTool.ps1 -Phase GiftHmSnap          -TargetIds JIGPL48S
.\VerifyTool.ps1 -Phase GiftMqSnap          -TargetIds JIGPL48S,JIDSL48S
.\VerifyTool.ps1 -Phase GiftJenkins         -RefreshUrls
.\VerifyTool.ps1 -Phase GiftJenkinsNoFile
.\VerifyTool.ps1 -Phase GfixHmSnap
.\VerifyTool.ps1 -Phase GfixJenkins
```

Clone evidence Excel (new):

```powershell
# Copy evidence templates / pre-existing files into work\evidence\<Excel_NAME>.xlsx
.\VerifyTool.ps1 -Phase Clone -CloneSourceDir D:\path\to\source
.\VerifyTool.ps1 -Phase Clone -BizCodes IGP2,ILP2
.\VerifyTool.ps1 -Phase Clone -Force
```

Align / J4 precheck (read-only by default):

```powershell
# If -J4BaseDir is omitted, VerifyTool uses the remembered J4BaseDir,
# then VerifyConfig.psd1 Align.J4BaseDir, then CloneSourceDir.
.\VerifyTool.ps1 -Phase Align -J4BaseDir D:\path\to\40.J4\07.GPCS\JRV
.\VerifyTool.ps1 -Phase Align
```

SEND data vs GIFT data metadata review (new):

```powershell
# Scans work\DATA\GIFT, writes work\data\gift_metadata.csv,
# then opens pending evidence Excel files for manual confirmation.
.\VerifyTool.ps1 -Phase SendVsGift
.\VerifyTool.ps1 -Phase SendVsGift -TargetIds JIGPC05S

# Stage 2 OCR (skeleton): exports the pictures on the send-data sheet,
# OCRs them with the built-in Windows engine (same family as the Snipping
# Tool text extraction, zero installs), writes work\data\send_metadata.csv
# and prints a field-by-field send-vs-gift verdict before the manual prompt.
# Enable per run with -Ocr, or persistently with SendVsGift.Ocr = $true in
# VerifyConfig.psd1 / verify_config.json.
.\SendVsGift.ps1 -WorkDir D:\work -Owner misaki -Ocr
```

See `docs/SendVsGift.md` for the Stage 1 metadata format and the Stage 2 OCR design.
The OCR verdict is advisory: Enter-to-mark remains the source of truth, and OCR
failure or low confidence simply falls back to the manual flow.

Replace evidence body (new):

```powershell
.\VerifyTool.ps1 -Phase ReplaceGift
.\VerifyTool.ps1 -Phase ReplaceGfix -TargetIds JIGPL48S
.\VerifyTool.ps1 -Phase ReplaceDf -Force

# Aliases also work:
.\VerifyTool.ps1 -Phase Rgift
.\VerifyTool.ps1 -Phase Rgfix
.\VerifyTool.ps1 -Phase Rdf
```

Visual review:

```powershell
.\VerifyTool.ps1 -Phase ReviewEvidence
.\VerifyTool.ps1 -Phase ReviewEvidence -CursorCell A1
.\VerifyTool.ps1 -Phase ReviewEvidence -TargetIds JIGPL48S
```

Delivery (final hand-off):

```powershell
.\VerifyTool.ps1 -Phase CheckSheet                          # fill the review check sheet
.\VerifyTool.ps1 -Phase CheckSheet -CheckSheetPath "\\srv\...\check.xlsx"
.\VerifyTool.ps1 -Phase DeliverMail                         # one Outlook draft per Excel
.\VerifyTool.ps1 -Phase DeliverMail -TargetIds SJRVWD64
```

## Common options

```powershell
-WorkDir <path>           # work folder. If omitted, the last one is reused.
-Owner <Owner>            # mapping owner suffix. No personal default is configured.
-TargetIds A,B            # limit by Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.
-CloneSourceDir <path>    # external path for Clone (existing evidence per bizcode).
-J4BaseDir <path>         # J4 baseline root for Align; defaults to config/CloneSourceDir/session.
-BizCodes A,B             # override bizcode candidate list for Clone.
-Force                    # redo rows that are already marked done.
-Interactive              # ask before each row where supported.
-WindowWidth 1050 -WindowHeight 761 -CropPx 6
-NoResize                 # do not resize Edge.
-RefreshUrls              # recapture Jenkins folder URLs.
-DryRun                   # print child-script arguments instead of running.
-ExcelPrefix <text>        # prefix before _<Excel_NAME>. CLI overrides config.
```

## Clone behavior

Phase: `Clone` (aliases: `MkExcel`, `RenameExcel`).

For each unique `Excel_NAME` in mapping (groups all rows sharing it):

1. Try `<SourceDir>\<bizcode>\<Workbook.ExcelPrefix>_<Excel_NAME>.xlsx` and
   `<SourceDir>\<bizcode>\<Excel_NAME>.xlsx` for each bizcode candidate.
2. Fallback to `<WorkDir>\template_<bizcode>.xlsx`.
3. Universal fallback to `<WorkDir>\template.xlsx`.
4. Copy -> `<WorkDir>\evidence\<Workbook.ExcelPrefix>_<Excel_NAME>.xlsx`
   (or `<Excel_NAME>.xlsx` when the prefix is blank).

Bizcode candidates come from `-BizCodes`, otherwise from the row's `TO_code` and `FROM_code` (deduplicated).

Skipped if the destination already exists, unless `-Force`.

## Replace behavior

Phases: `ReplaceGift`, `ReplaceGfix`, `ReplaceDf`. All call `ReplaceEvidence.ps1` with a different `-Mode`.

Per unique `Excel_NAME` (groups all `Correl_ID_S` sharing it):

1. Open `work\evidence\<Workbook.ExcelPrefix>_<Excel_NAME>.xlsx` (or the legacy row `Excel_Prefix` override if present).
2. Find target sheet by mode:
    - Gift -> `GIFT受信結果`
    - Gfix -> `GFIX受信結果`
    - Df   -> `GIFTデータvsGFIXデータ`
3. Reset row 3 downward (delete shapes, clear values + formatting + highlighting).
4. Insert images stacked at column B with blank rows between, picture z-order = `msoSendToBack` so later mark rectangles stay visible.
5. Tail per mode:
    - Gift -> label `GFIX Jenkins フォルダ受信ファイルなし` once, then `GIFT_noGfixfile\<correl>.png` stacked.
    - Gfix -> per-correl label `GFIX受信log` then placeholder text `<<TODO: GFIX 受信 log>>` (real implementation pending `GfixLodDownload`).
    - Df   -> nothing extra.
6. Save the workbook.
7. On all-OK: `isReplaced |= bit` (Gift=1, Gfix=2, Df=4) for every row in the group.

### isReplaced bitmask

A single integer column. Bits set independently per mode:

```text
bit 0 (1) : GIFT replace done
bit 1 (2) : GFIX replace done
bit 2 (4) : DF replace done
total 7    : all three done
```

Status display recognizes the bitmask via `BitValue` entries in `PhaseOrder`. `(value -band BitValue) == BitValue` means "done for that mode."

Bits are only set when every image for that mode actually existed and was inserted without error. If any image was missing or any step failed, the bit stays cleared and the run is reported as failed for that group.

## ReviewEvidence behavior

`ReviewEvidence.ps1` opens files from `work\evidence` according to `mapping_<Owner>.csv`.

For each pending row:

1. Find evidence Excel by `Excel_NAME`, fallback to `JOB_NAME`, `Correl_ID_S`, `Correl_ID_M`.
2. Open workbook through Excel COM and jump to the current ID on the send-data sheet (`送信データ` column A exact match, falling back to `-CursorCell`).
3. You check that ID manually.
4. Press Enter in the shell. The script marks only that ID as reviewed.
5. If the same workbook still has pending IDs, Excel stays open and the cursor jumps to the next ID; it does not reset all sheet cursors, save, or close yet.
6. After all IDs in the workbook are reviewed, the script selects `A3` by default on every sheet from last to first, sends Ctrl+S and Esc, waits briefly, closes, and updates the mapping.

Use `-CursorCell A1` if the fallback cursor rule changes.

## CheckSheet behavior

Phase: `CheckSheet` (aliases: `FillCheckSheet`, `RvCheck`). Appends one row per
evidence Excel (grouped by `Excel_NAME`) to the shared review check sheet, sheet
`Check Sheet_J4`. Columns written: A No. (continued from the last numeric No.,
only when blank), B 記入日 (today, number format copied from the row above),
C `JAVA`, E `J4内部ﾚﾋﾞｭｰ`, F full evidence filename
(`<Workbook.ExcelPrefix>_<Excel_NAME>.xlsx`; legacy row `Excel_Prefix` still overrides when present), G owner, H reviewer (configured reviewer). D/I/J~ are
left blank.

Because the check sheet is a public document the write is double-checked:

1. Snapshot the original's timestamp + size.
2. Copy it to a TEMP file, fill the planned rows there, and open it (visible) for review.
3. Press Enter to commit, or `q` to abort (nothing written).
4. The original is re-stat'd; the identical edits are committed **only if it is
   unchanged** since the preview began. If it changed, the write is held so you
   can re-run against the new content.

Already-listed Excels (matched on column F) are skipped unless `-Force`. The
path comes from `CheckSheet.Path` in `VerifyConfig.psd1`; if it does not exist
the phase prompts and remembers the answer in `verify_session.json` (or pass
`-CheckSheetPath`).

## DeliverMail behavior

Phase: `DeliverMail` (aliases: `Mail`, `SendMail`, `Deliver`). Builds one Outlook
**draft** per `Excel_NAME` via Outlook COM (`CreateItem` + `Display`) — it never
sends automatically. Subject is
`【GIFT廃止対応】<Phase>レビュー依頼(<Excel_NAME>)`; the body, reviewer (`To`),
and UNC paths are all config-driven (`Mail` / `Reviewer` in `VerifyConfig.psd1`).

For each pending group:

1. The draft opens in Outlook. You eyeball it and click **Send** yourself.
2. Return to the shell and press Enter to set `isDelivered = 1` for that Excel.
   `s` skips, `q` quits, and `-m "comment"` records a note in `DeliverComment`.

Outlook is released at the end but never Quit (it may be your live session).

## Phase status

Implemented:

```text
Mapping
GiftHmSnap, GiftMqSnap, GiftJenkins, GiftJenkinsNoFile
GfixHmSnap, GfixJenkins
Clone
ReplaceGift, ReplaceGfix, ReplaceDf
ReviewEvidence
Crop
Validate    (read-only diagnostic)
Status
```

Legacy but callable:

```text
ExcelSnap
```

Registered but not implemented yet:

```text
GfixLodDownload
DfSnap
```

`ReplaceGfix` writes a placeholder for the receive-side log section. Once `GfixLodDownload` lands, replace `Get-GfixLogLines` in `ReplaceEvidence.ps1` with grep against `work\log\` (see the function header comment).

## Design rule

`Common.ps1` stays primitive: window activation, screenshot helpers, SendKeys, small shared utilities.

`ExcelHelpers.ps1` is a dot-source library for Excel COM operations. It has no `param()` block so dot-sourcing is safe. Tool scripts with `param()` are never dot-sourced; they are called via `& $path @args`.

Project defaults live in `VerifyConfig.psd1`. Workflow decisions live in `VerifyTool.ps1`. Feature-specific automation lives in dedicated scripts (`HmSnap.ps1`, `JenkinsSnap.ps1`, `ReviewEvidence.ps1`, `Clone.ps1`, `ReplaceEvidence.ps1`).

This avoids the dot-source bug pattern where a child utility script can accidentally overwrite parent switch parameters.

## File list

```text
VerifyTool.ps1          main entry, menu, phase router, status display
VerifyConfig.psd1       project config (paths, scripts, phase order, aliases, labels)
verify_session.json     last-used cache; WorkDir fallback and legacy/session fallbacks

ExcelHelpers.ps1        dot-source library: Excel COM helpers, bitmask helpers
Clone.ps1               Clone phase
ReplaceEvidence.ps1     ReplaceGift / ReplaceGfix / ReplaceDf
Validate.ps1            Validate phase (read-only diagnostic)

JenkinsSnap.ps1         Jenkins-side capture + download
(HmSnap.ps1, MqSnap.ps1, ExcelSnap.ps1, ReviewEvidence.ps1, Common.ps1,
 Generate-HostOpenMapping.ps1, Crop-Snap.ps1 -- kept as before)

CLAUDE.md               context file for Claude Code (read it first when opening from IDE)
README.md               this file
CHANGELOG.md            iteration log across browser-IDE bridge

Pack-LlmContext.ps1     プロジェクト全体のコンテキストをクリップボードにコピー
Apply-LlmPatch.ps1      AIが生成したXMLパッチをローカルファイルに自動適用
Export-DailyPatch.ps1   本日のGit差分のみを抽出してクリップボードにコピー
```
