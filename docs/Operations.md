# Operations reference

Day-to-day reference for running VerifyTool. Start at [`README.md`](../README.md)
for what the tool is; read [`CLAUDE.md`](../CLAUDE.md) for how it is built.

Normal use is the interactive menu:

```powershell
.\VerifyTool.ps1
.\VerifyTool.ps1 -Help
```

Everything below is either something the menu asks you anyway, or a flag for
scripting a run non-interactively.

---

## Contents

- [Work folder layout](#work-folder-layout)
- [The mapping CSV](#the-mapping-csv)
- [Per-work-folder config](#per-work-folder-config-verify_configjson)
- [Common options](#common-options)
- [Phases](#phases)
- [Phase behaviour in detail](#phase-behaviour-in-detail)
- [Full-width filename fallback](#full-width-filename-fallback)

---

## Work folder layout

```text
work\
  mapping_<Owner>.csv            job list + state (one row per data transfer)
  verify_config.json             optional per-folder config overlay
  wipGFIX一覧.xlsx                source workbook the mapping is generated from
  template.xlsx                  universal evidence template (optional)
  template_<bizcode>.xlsx        per-bizcode evidence template (optional)
  snap\
    excel\  GIFT_HM\  GIFT_MQ\  GIFT_Jenkins\  GIFT_noGfixfile\
    GFIX_HM\  GFIX_Jenkins\  DF\
    ProcessTime\<correl>\        per-correl OCR cache + exported pictures
  DATA\
    GIFT\<Correl_ID_S>*          receive files downloaded from Jenkins
    GFIX\<Correl_ID_S>*
    GFIX_LOD\
  log\                           GFIX receive logs (GoAnywhere downloads)
  evidence\
    <Excel_NAME>.xlsx            the deliverable workbooks
  status\
    progress.jsonl               append-only event log (tail it live)
  bk\                            timestamped J4 backups (BackupJ4)
```

`VerifyTool.ps1` remembers the last `WorkDir`, `Owner`, window size, crop size,
evidence folder, review cursor cell, `CloneSourceDir`, `J4BaseDir`,
`CheckSheetPath` and `DfExePath` in `verify_session.json` — machine/operator
state only. Project-scoped answers persist to the work folder's
`verify_config.json` instead (see [`Configuration.md`](Configuration.md)).

## The mapping CSV

`mapping_<Owner>.csv`, UTF-8 **with** BOM (Excel needs it for Japanese). One row
per `Correl_ID_S`. Phase state lives in dedicated columns; the tool adds any it
is missing on startup (`RepairMapping`).

### Bitmask columns

Three columns pack three modes into one integer:

| Column | bit 1 | bit 2 | bit 4 | all done |
|---|---|---|---|---|
| `isReplaced` | GIFT | GFIX | DF | 7 |
| `isMarked` | GIFT | GFIX | DF | 7 |
| `isReviewed` | GIFT | GFIX | DF | 7 |

Done for a mode means `(value -band bit) -eq bit`. A bit is set only when
**every** required piece for that mode succeeded; which correl/step/file failed
is recorded in `status\progress.jsonl`, not in extra columns.

`ProcessTime_Inserted` is also a bitmask: bit 1 = this correl's OCR result is
extracted and cached, bit 2 = the row has been written to an output workbook,
3 = both.

### Plain value columns

| Column | Meaning |
|---|---|
| `<phase>_snap` | `0`/empty pending · `1` OK · `2` NG (SnapVerify flagged it). NG stays pending and is listed in the end-of-run summary. |
| `SendVsGift` | `0`/empty pending · `1` OK · `2` NG. `2` is re-offered next run. |
| `isDelivered` | `0`/`1` per `Excel_NAME`, set when you confirm the review mail was sent. |
| `GIFT_ProcessTime` / `GFIX_ProcessTime` | informational: `0` not attempted · `1` extracted · `2` not found. |
| `ReviewComment` / `DeliverComment` | free text, captured with `-m "comment"` at the prompt. |
| `SS_CODE` | optional override for GFIX log matching; when present and non-empty it wins over the inferred SS code. |
| `Expected_Time` | expected run time for the snap time-window check. |

Watch progress from a second shell without locking the CSV:

```powershell
.\Watch-MappingProgress.ps1 -WorkDir <path> -Owner <owner>
```

## Per-work-folder config (verify_config.json)

`VerifyConfig.psd1` holds the project defaults. Each work folder may carry a
`verify_config.json` that is **deep-merged over** them at startup, so one case
can customize almost everything without editing the shared `.psd1`.

Precedence: **CLI args > work-folder `verify_config.json` > `VerifyConfig.psd1` >
session fallback** (for the few values that still support session fallback).

```powershell
.\VerifyTool.ps1 -Phase InitConfig               # no file yet -> full snapshot; file exists -> REPAIR
.\VerifyTool.ps1 -Phase InitConfig -Interactive  # grouped walk/peek/edit/delete/save UI
.\VerifyTool.ps1 -Phase InitConfig -Force        # full regenerate (keeps a .bak)
```

When the file already exists the default run is a **repair/update**: your values
are untouched, a sparse hand-written file stays sparse, and only fields the tool
has gained since the file was last written are appended — each one listed on the
console. Repair knows what is new from the hidden `_SCHEMA` inventory the file
carries (leave `_SCHEMA` alone; it is ignored at runtime). A file with no
`_SCHEMA` stamp is only stamped on its first repair, so a full snapshot is never
dumped into a sparse file.

Interactive mode groups settings by `intro`, `phase`, `snap`, `excel`, `wbs`,
`path`, `mail`, `all`. Pick `w` to **walk** a group field by field (Enter =
keep, a value = set, `-d` = delete, `q` = stop); `v`/`e`/`d` peek/edit/delete by
JSON path (`Window.Width`, `Mail.BodyLines`) when you know what to touch.
Changes land on disk only after `s` and typing `YES`; if the write fails because
the JSON is open elsewhere, nothing is lost — `r` retries, Enter returns to the
menu with edits kept.

`InitConfig` also writes `verify_config.README.txt` next to the JSON with field
explanations, since standard JSON has no comments.
`verify_config.example.json` in the repo is a ready-to-copy starter.

Duplicated fields were merged into single canonical top-level ones:
`J4EvidenceDir` (was `DeliverFiles.J4EvidenceDir` / `Mail.EvidenceFolder`) and
`Address` (was `Reviewer.Address`). Old files setting the legacy fields keep
working — a non-empty legacy value wins — but InitConfig no longer generates
them.

## Common options

```powershell
-WorkDir <path>           # work folder. If omitted, the last one is reused.
-Owner <Owner>            # mapping owner suffix. No personal default is configured.
-TargetIds A,B            # limit by Correl_ID_S / Correl_ID_M / JOB_NAME / Excel_NAME.
-Force                    # redo rows already marked done.
-Interactive              # ask before each row where supported.
-DryRun                   # print child-script arguments instead of running.
-CloneSourceDir <path>    # external path for Clone (existing evidence per bizcode).
-J4BaseDir <path>         # J4 baseline root for Align; falls back to config/CloneSourceDir/session.
-BizCodes A,B             # override the bizcode candidate list for Clone.
-ExcelPrefix <text>       # prefix before _<Excel_NAME>. CLI overrides config.
-WindowWidth 1050 -WindowHeight 761 -CropPx 6
-CropLeft/-CropTop/-CropRight/-CropBottom <px>   # per-side crop (-1 = inherit CropPx).
                          # Per-snap-folder overrides: Window.CropByFolder in verify_config.json.
-NoResize                 # do not resize Edge.
-RefreshUrls              # recapture Jenkins folder URLs.
```

## Phases

Run any of these from the menu, or as `.\VerifyTool.ps1 -Phase <Key>`. Most
accept aliases (`Rgift` for `ReplaceGift`, `Pt` for `ProcessTime`, …); `-Help`
lists them.

| Key | What it does |
|---|---|
| `InitConfig` | create/repair the work folder's `verify_config.json` |
| `Mapping` | generate or grow the job list from the WBS workbook |
| `GiftHmSnap` / `GfixHmSnap` | capture the HM batch-status page per correl |
| `GiftMqSnap` | capture the MQ transfer-status page |
| `GiftJenkins` / `GfixJenkins` | capture the Jenkins file list and download receive files |
| `GiftJenkinsNoFile` | the "no GFIX file expected" case |
| `GfixLogDownload` | pull the GFIX receive logs from GoAnywhere |
| `DfSnap` | run `df.exe` on the GIFT/GFIX data pair and capture the diff |
| `Clone` | create one evidence workbook per `Excel_NAME` |
| `Align` | compare work evidence against the J4 baseline (precheck) |
| `SendVsGift` | send-data vs GIFT-data metadata review (see [SendVsGift.md](SendVsGift.md)) |
| `ReplaceGift` / `ReplaceGfix` / `ReplaceDf` | insert the captures into the evidence workbook |
| `ProcessTime` | extract each batch's start/end time and build the processing-time workbooks |
| `MarkGift` / `MarkGfix` / `MarkDf` | draw the red rectangles |
| `ReviewGift` / `ReviewGfix` / `ReviewDf` / `ReviewEvidence` | human review walk-through |
| `Comments` | list the review notes (read-only) |
| `CheckSheet` | append rows to the shared review check sheet |
| `BackupJ4` | copy the current J4 workbooks to a local timestamped folder |
| `DeliverFiles` | replace the delivery-scope sheets in the J4 workbooks + copy DATA |
| `DeliverMail` | one Outlook draft per `Excel_NAME` |
| `Validate` | read-only readiness diagnostic — run this first in a new session |
| `Status` | mapping column scan |
| `RepairMapping` | add missing columns (runs automatically on startup) |
| `ProbeShapes` | list evidence-workbook shapes (calibration aid) |
| `Crop` | bulk-crop existing PNGs |
| `WatchProgress` | live progress monitor (read-only, never locks the CSV) |
| `ExcelSnap` | legacy, kept callable |

---

## Phase behaviour in detail

### Capture phases and SnapVerify

Every capture is checked against the page's own text before the row is marked.
The phase asks once per run for a run time (`[Enter]` = now, an explicit
`yyyy/MM/dd HH:mm:ss`, or `n` to skip the time check) and a tolerance, fills any
empty `Expected_Time` cells on the pending rows, then after each capture reads
the page text back and judges it.

A failure sets the snap column to `2` (NG) rather than `1`: the row stays
pending, is re-offered next run, and appears in the end-of-run NG summary. The
captured page text is saved next to the PNG as `<correl>.txt`. If the captured
text is not the expected page at all, the phase stops and asks `r`/`s`/`q`.

What counts as NG per page: HM — abend status, no matching correl, a run time
outside the window. MQ — "No Data!", no matching correl, receive time outside
the window, non-zero Rtncd/Rsncd. Jenkins — the expected file absent from the
list (or, for the no-GFIX case, unexpectedly present).

Set `SnapVerify.Enabled = $false` to go back to plain screenshots.
Full rules: [`SnapVerify-Plan.md`](SnapVerify-Plan.md).

### Clone

For each unique `Excel_NAME` in the mapping (grouping all rows that share it):

1. Try `<SourceDir>\<bizcode>\<ExcelPrefix>_<Excel_NAME>.xlsx`, then
   `<SourceDir>\<bizcode>\<Excel_NAME>.xlsx`, for each bizcode candidate.
2. Fall back to `<WorkDir>\template_<bizcode>.xlsx`.
3. Fall back to `<WorkDir>\template.xlsx`.
4. Copy to `<WorkDir>\evidence\<ExcelPrefix>_<Excel_NAME>.xlsx` (or
   `<Excel_NAME>.xlsx` when the prefix is blank).

Bizcode candidates come from `-BizCodes`, otherwise from the row's `TO_code` and
`FROM_code` (deduplicated). Skipped when the destination exists, unless `-Force`.

### Replace

`ReplaceGift` / `ReplaceGfix` / `ReplaceDf` all call `ReplaceEvidence.ps1` with a
different `-Mode`. Work is **plan-driven**: `EvidencePlan.ps1` builds a pure,
correl-major plan encoding the review order, and `EvidenceExecutor.ps1` walks it
and performs the inserts. Per unique `Excel_NAME`:

1. Open the evidence workbook.
2. Pick the target sheet by mode (GIFT receive result / GFIX receive result /
   GIFT-vs-GFIX data compare).
3. Reset row 3 downward — delete shapes, clear values, formatting, highlighting.
4. Insert the images stacked at column B with blank rows between; picture
   z-order `msoSendToBack` so the later mark rectangles stay visible.
5. Mode tail — GIFT: the "no GFIX file" label once, then the `GIFT_noGfixfile`
   captures. GFIX: the per-correl receive-log label, then the matched log pasted
   whole (`GfixLog.ps1` decides which log belongs to which correl by its
   `Command:` line, not by file name).
6. Save.
7. On all-OK, set the mode's `isReplaced` bit for every row in the group.

A snap file saved under a batch-stamped correl id resolves against a workbook
that shows the plain id, and vice versa.

### ProcessTime

Extracts each correl's HM batch start/end time for both sides and derives the
duration, then writes one row per side per correl into
`処理時間(<Tag>).xlsx` workbooks.

Source priority per side, best first:

1. `snap\<Stage>_HM\<correl>.txt` — archived Ctrl+A page text. Not OCR, so
   immune to digit misreads; trusted absolutely.
2. OCR of `snap\<Stage>_HM\<correl>.png` — the original per-correl screenshot.
3. OCR of the picture already inserted in the evidence workbook, validated by
   content (the correl id must appear in the OCR text unless the picture sits in
   the trusted section position).

`-Stage Ocr|Write|Both` splits extraction from writing; the OCR result is cached
per correl at `snap\ProcessTime\<correl>\result.json`, so a `Write`-only rerun
opens no evidence workbook at all.

**The 3↔9 digit problem.** The Japanese OCR engine confuses MS Gothic `9` and
`3`, and a misread produces a perfectly valid-looking timestamp. The tool's rule
is to act only where the answer is forced (`TimeDigitVerify.ps1`):

- A digit whose field falls out of range is corrected when exactly one 3↔9
  substitution brings it back — a minute reading `93` can only have been `33`.
  Ambiguous cases are left alone.
- Start, end and the page's own printed processing-time column are three
  readings of one fact. When they disagree and exactly one substitution
  reconciles all three, that fix is applied (and noted). When several fixes or
  none would work, **nothing is rewritten** and the row is flagged 要確認 in the
  verify column.
- Every remaining ambiguous `3`/`9` — one sitting where the opposite digit would
  also be legal — turns the cell **red** via conditional formatting on columns
  D/E/F, so you can check it against the snap image. The correl-id cell
  hyperlinks straight to that image.

Turn the red marking off with `ProcessTime.EmitDigitFormat = $false`.

Output columns: A–H data, then the audit columns I 処理時間(検算) `=E-D`,
J チェック (T/F compare), K 件数(参照), L 件数チェック, then the 検証 verify
column. Output routing is config-driven (`ProcessTime.OutputTags`,
`OutputMode`, `OutputDirectoryByTag`).

### Mark

Draws the red rectangles reviewers look for. Each `Mark.Boxes` entry may add a
`Template` key to attempt image-recognition placement (`Locate-ByImage.ps1`
matches the reference crop against the source snap PNG) before falling back to
the fixed `OffsetX`/`OffsetY` box. Console lines are tagged `[MARK-IMG]` when a
template matched and `[MARK]` when the fixed offset was used, so a run makes the
path obvious. `mark_templates/` ships empty — see its README for how to capture
templates on an office PC.

`MarkGfix` also highlights the GFIX log `Command:` row, auto-sized to the row's
actual text width.

### Review

`ReviewEvidence.ps1` opens workbooks from `work\evidence` per the mapping. For
each pending row:

1. Find the evidence Excel by `Excel_NAME`, falling back to `JOB_NAME`,
   `Correl_ID_S`, `Correl_ID_M`.
2. Open it through Excel COM and jump to the current id on the send-data sheet
   (column A exact match, falling back to `-CursorCell`, default `A3`).
3. You check that id.
4. Press Enter. Only that id is marked reviewed. `-m "comment"` records a note.
5. If the workbook still has pending ids, Excel stays open and the cursor jumps
   to the next one — no save or close yet.
6. Once every id in the workbook is done, the cursor is reset on every sheet
   from last to first, Ctrl+S and Esc are sent, and the workbook closes.

All review phases accept `-J4` (menu option `j4`): the workbook opens from the
delivered J4 folder instead of `work\evidence`, for re-checking delivered copies
after `DeliverFiles`. Saves land on the J4 file; the local mapping updates as
usual.

### CheckSheet

Appends one row per evidence Excel (grouped by `Excel_NAME`) to the shared
review check sheet, sheet `Check Sheet_J4`. Written: A No. (continued from the
last numeric No., only when blank), B 記入日 (today, number format copied from
the row above), C `JAVA`, E `J4内部ﾚﾋﾞｭｰ`, F the full evidence filename, G owner,
H the configured reviewer. D/I/J~ are left blank.

Because the check sheet is a shared document, the write is double-checked:

1. Snapshot the original's timestamp and size.
2. Copy it to a TEMP file, fill the planned rows there, and open it for review.
3. Enter commits, `q` aborts (nothing written).
4. The original is re-stat'd, and the identical edits are committed **only if it
   is unchanged** since the preview began. If it changed, the write is held so
   you can re-run against the new content.

Already-listed Excels (matched on column F) are skipped unless `-Force`. The
path comes from `CheckSheet.Path`; if it does not exist the phase prompts and
remembers the answer, or pass `-CheckSheetPath`.

### BackupJ4 → DeliverFiles → DeliverMail

Run `BackupJ4` ("bk") first: it is read-only against J4 and copies each targeted
`Excel_NAME`'s current J4 workbook into a local timestamped folder (default
`<WorkDir>\bk`), giving you a rollback point.

`DeliverFiles` then replaces the three delivery-scope sheets (GIFT/GFIX receive
result + the GIFT-vs-GFIX data compare) in the corresponding J4 workbook with
the matching work sheets, **in place** — other J4 sheets are untouched. The
first delivery for an `Excel_NAME` copies the whole file instead. `DATA\GFIX`
and `DATA\GIFT` are copied too. Source files are never deleted. Sets
`isFilesDelivered`.

`DeliverMail` builds one Outlook **draft** per `Excel_NAME` (`CreateItem` +
`Display`) — it never sends. Subject, body, reviewer and UNC paths are all
config-driven. For each group the draft opens, you eyeball it and click Send
yourself, then return to the shell and press Enter to set `isDelivered = 1`
(`s` skips, `q` quits, `-m "comment"` records a note). Outlook is released at
the end but never Quit — it may be your live session.

## Full-width filename fallback

Customer-provided files sometimes use full-width ASCII in filenames (`０`
instead of `0`). Workbook lookup tries the normal exact and wildcard paths
first; if nothing is found it scans for names whose full-width ASCII normalizes
to the requested one, warns, and asks before using the candidate.

```powershell
# Generic file lookup, after your own not-found branch.
Resolve-FullWidthFileName -Dir $dir -Name 'report0.txt' -Filter '*.txt' -FullWidthFallback Prompt

# Workbook-specific lookup used by the evidence phases.
Find-WorkbookByExcelName -Dir $evDir -ExcelName $fullStem -FullWidthFallback Prompt
```

Interactive tools keep the default `Prompt`; tests and non-interactive batch
flows should pass `Accept` or `Reject` explicitly.
