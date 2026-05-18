# VerifyTool

`VerifyTool.ps1` is the main entry for the verification evidence workflow.

Daily use:

```powershell
.\VerifyTool.ps1
.\VerifyTool.ps1 -Help
```

## Folder layout

Expected work folder:

```text
work\
  mapping_厳.csv
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
    GIFT\<Correl_ID_S>\
    GFIX\<Correl_ID_S>\
    GFIX_LOD\
  evidence\
    <Excel_NAME>.xlsx
```

Jenkins-downloaded files keep their DATA layout, for example:

```text
\\fs-f3170-1\12_生産管理\00121.GPCS\31.NII\other：その他\個人用ワーク\厳\Work\0514_JRV-IDS,IGP2\DATA\GIFT\JIDSK48S
```

`VerifyTool.ps1` remembers the last `WorkDir`, `Owner`, window size, crop size, evidence folder, review cursor cell, and `CloneSourceDir` in `verify_session.json`.

## Main commands

Status (CSV column scan):

```powershell
.\VerifyTool.ps1 -Phase Status
```

Validate (read-only readiness diagnostic — run this first when entering a new work session):

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

## Common options

```powershell
-WorkDir <path>           # work folder. If omitted, the last one is reused.
-Owner 厳                 # mapping owner suffix. Default: 厳.
-TargetIds A,B            # limit by Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.
-CloneSourceDir <path>    # external path for Clone (existing evidence per bizcode).
-BizCodes A,B             # override bizcode candidate list for Clone.
-Force                    # redo rows that are already marked done.
-Interactive              # ask before each row where supported.
-WindowWidth 1050 -WindowHeight 761 -CropPx 6
-NoResize                 # do not resize Edge.
-RefreshUrls              # recapture Jenkins folder URLs.
-DryRun                   # print child-script arguments instead of running.
```

## Clone behavior

Phase: `Clone` (aliases: `MkExcel`, `RenameExcel`).

For each unique `Excel_NAME` in mapping (groups all rows sharing it):

1. Try `<SourceDir>\<bizcode>\<Excel_NAME>.xlsx` for each bizcode candidate.
2. Fallback to `<WorkDir>\template_<bizcode>.xlsx`.
3. Universal fallback to `<WorkDir>\template.xlsx`.
4. Copy → `<WorkDir>\evidence\<Excel_NAME>.xlsx`.

Bizcode candidates come from `-BizCodes`, otherwise from the row's `TO_code` and `FROM_code` (deduplicated).

Skipped if the destination already exists, unless `-Force`.

## Replace behavior

Phases: `ReplaceGift`, `ReplaceGfix`, `ReplaceDf`. All call `ReplaceEvidence.ps1` with a different `-Mode`.

Per unique `Excel_NAME` (groups all `Correl_ID_S` sharing it):

1. Open `work\evidence\<Excel_NAME>.xlsx`.
2. Find target sheet by mode:
    - Gift → `GIFT受信結果`
    - Gfix → `GFIX受信結果`
    - Df   → `GIFTデータvsGFIXデータ`
3. Reset row 3 downward (delete shapes, clear values + formatting + highlighting).
4. Insert images stacked at column B with blank rows between, picture z-order = `msoSendToBack` so later mark rectangles stay visible.
5. Tail per mode:
    - Gift → label `GFIX Jenkins フォルダ受信ファイルなし` once, then `GIFT_noGfixfile\<correl>.png` stacked.
    - Gfix → per-correl label `GFIX受信log` then placeholder text `<<TODO: GFIX 受信 log>>` (real implementation pending `GfixLodDownload`).
    - Df   → nothing extra.
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

`ReviewEvidence.ps1` opens files from `work\evidence` according to `mapping_厳.csv`.

For each pending row:

1. Find evidence Excel by `Excel_NAME`, fallback to `JOB_NAME`, `Correl_ID_S`, `Correl_ID_M`.
2. Open workbook through Excel COM.
3. You check it manually.
4. Press Enter in the shell.
5. The script selects `A3` by default on every sheet from last to first, saves, closes, and updates `isReviewed = 1`.

Use `-CursorCell A1` if the review rule changes.

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
verify_session.json     remembers last settings between runs

ExcelHelpers.ps1        dot-source library: Excel COM helpers, bitmask helpers
Clone.ps1               Clone phase
ReplaceEvidence.ps1     ReplaceGift / ReplaceGfix / ReplaceDf
Validate.ps1            Validate phase (read-only diagnostic)

JenkinsSnap.ps1         Jenkins-side capture + download
(HmSnap.ps1, MqSnap.ps1, ExcelSnap.ps1, ReviewEvidence.ps1, Common.ps1,
 Generate-HostOpenMapping.ps1, Crop-Snap.ps1 — kept as before)

CLAUDE.md               context file for Claude Code (read it first when opening from IDE)
README.md               this file
CHANGELOG.md            iteration log across browser-IDE bridge
HANDOFF.md              architecture notes for any follow-up engineer
```
