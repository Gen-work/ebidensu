# VerifyTool hand-off document

## Current goal

Build one human-friendly verification entry point. The operator starts one shell, picks a phase, and the tool guides them through evidence collection, download, replacement, and review.

Main entry:

```powershell
.\VerifyTool.ps1
```

Help:

```powershell
.\VerifyTool.ps1 -Help
```

## Core design

Do not put all parameters into `Common.ps1`.

Use this split instead:

```text
Common.ps1          stable primitives only
VerifyConfig.psd1   project defaults and phase registry
VerifyTool.ps1      interactive entry / stage decision / argument routing
*.ps1               phase-specific tools
```

Reason: `Common.ps1` is dot-sourced by many scripts. If it starts holding workflow state, it becomes another source of accidental variable pollution. The earlier `$Force` bug came from this class of problem.

## Implemented in this build

### VerifyTool.ps1

New main entrance. It:

- remembers last `WorkDir` in `verify_session.json`
- reads `mapping_厳.csv`
- displays phase status
- recommends next unfinished phase
- supports `-Help`
- routes old phase aliases to the new phase names
- keeps old scripts callable

Old aliases still work:

```text
Excel -> ExcelSnap
HmGift -> GiftHmSnap
MqGift -> GiftMqSnap
JenkinsGift -> GiftJenkins
NoGfix -> GiftJenkinsNoFile
HmGfix -> GfixHmSnap
JenkinsGfix -> GfixJenkins
Review -> ReviewEvidence
```

### ReviewEvidence.ps1

This is implemented because it is relatively safe and useful now.

Behavior:

1. Load `mapping_厳.csv`.
2. Open pending evidence Excel files under `work\evidence`.
3. Let the operator review by eye.
4. On Enter, select `A3` by default from last sheet to first sheet.
5. Save and close.
6. Update `isReviewed = 1`.

Useful calls:

```powershell
.\VerifyTool.ps1 -Phase ReviewEvidence
.\VerifyTool.ps1 -Phase ReviewEvidence -CursorCell A1
.\VerifyTool.ps1 -Phase ReviewEvidence -TargetIds JIGPL48S
```

### JenkinsSnap.ps1 path rule

Jenkins download output keeps the existing DATA layout:

```text
work\DATA\GIFT\<Correl_ID_S>
work\DATA\GFIX\<Correl_ID_S>
```

No-GFIX snap output folder also keeps the existing name:

```text
work\snap\GIFT_noGfixfile
```

Mapping field remains:

```text
GIFT_noGfixfile_snap
```

## Planned phase definitions

### ExcelSnap exact-range future version

Desired behavior:

1. Open `wipGFIX一覧.xlsx`.
2. Maximize / full screen.
3. Filter `JOB_NAME` at `O5`.
4. Snap exact range from `B4:O?`.
5. `?` is last result row, for example `O292`.
6. Save to `work\snap\excel\<JOB_NAME>.png`.
7. Update `Excel_snap = 1`.
8. Repeat until all rows complete.

Implementation options:

Option A, preferred if possible: use Excel COM `Range.CopyPicture()` instead of pixel screenshot. This avoids row-height measurement and full-screen dependency. The script can compute the visible filtered range and export it as an image.

Option B: use full-screen pixel capture. Then the script must:

- send Ctrl+Home
- know fixed top-left pixel for `B4`
- measure row height
- calculate bottom pixel by visible row count
- crop the screenshot

Option A is more stable. Option B is closer to current manual intuition but more fragile.

### GIFT HM snap

Desired behavior:

1. Remind operator to open HM.
2. Check `isMultiAppl` / `TO_code` grouping.
3. For each appl/TO, ask operator to switch to the correct HM page.
4. Search `Correl_ID_S`.
5. Save screenshot to `work\snap\GIFT_HM\<Correl_ID_S>.png`.
6. Update `GIFT_HM_snap = 1`.

Current `HmSnap.ps1 -Stage GIFT` mostly covers this.

### GFIX HM snap

Same as GIFT HM, but:

```text
work\snap\GFIX_HM\<Correl_ID_S>.png
GFIX_HM_snap = 1
```

Current `HmSnap.ps1 -Stage GFIX` mostly covers this.

### GIFT MQ snap

Desired behavior:

1. Remind operator to open MQ.
2. Press Enter to switch to MQ window.
3. Search evidence by `Correl_ID_S`.
4. Do not open detail.
5. Save to `work\snap\GIFT_MQ\<Correl_ID_S>.png`.
6. Update `GIFT_MQ_snap = 1`.

Current `MqSnap.ps1` mostly covers this, but tab counts may need adjustment.

### GIFT Jenkins snap/download

Desired behavior:

1. Open Jenkins folder page.
2. Ctrl+F.
3. Paste `Correl_ID_S`.
4. Screenshot list page.
5. Download file to `work\DATA\GIFT\<Correl_ID_S>`.
6. Save screenshot to `work\snap\GIFT_Jenkins\<Correl_ID_S>.png`.
7. Update `GIFT_Jenkins_snap = 1`.
8. If multi appl, ask the operator to switch page and repeat.

Current `JenkinsSnap.ps1 -Mode GiftRecv` mostly covers this.

### GIFT Jenkins no-GFIX-file snap

Desired behavior:

1. Use GFIX Jenkins folder page.
2. Search `Correl_ID_S`.
3. Capture no-file evidence.
4. Save to `work\snap\GIFT_noGfixfile\<Correl_ID_S>.png`.
5. Update `GIFT_noGfixfile_snap = 1`.
6. No download.

Current `JenkinsSnap.ps1 -Mode NoGfix` covers this with the new folder name.

### GFIX Jenkins snap/download

Same as GIFT Jenkins, but:

```text
work\DATA\GFIX\<Correl_ID_S>
work\snap\GFIX_Jenkins\<Correl_ID_S>.png
GFIX_Jenkins_snap = 1
```

Current `JenkinsSnap.ps1 -Mode GfixRecv` mostly covers this.

### GFIX LOD download

Not implemented.

Desired behavior:

1. Accept dirty multiline pasted input.
2. Ignore CR/LF noise until the operator types `EOF` on its own line.
3. Parse paths/URLs/log identifiers.
4. Download or collect LOD/log files.
5. Save under `work\DATA\GFIX_LOD`.
6. Update `GFIX_log` or related mapping field.

Important design point:

Implement input as a dedicated function, not `Read-Host` per line scattered across the script:

```powershell
function Read-MultilineUntilEof {
    $lines = @()
    while ($true) {
        $line = Read-Host
        if ($line -eq 'EOF') { break }
        if ($null -ne $line) { $lines += $line }
    }
    return ($lines -join "`n")
}
```

The exact regex should be decided after seeing real dirty input samples.

### DF snap

Not implemented.

Desired behavior:

1. Open `C:\tools\DF\DF.exe`.
2. Input two file paths.
3. Capture screen.
4. Save to `work\snap\DF\<Correl_ID_S>.png`.
5. Update `DF_snap = 1`.

Implementation note:

Prefer driving the app by stable keyboard sequence only if the UI is simple. If the app has stable window title or controls, use WinAPI targeting rather than foreground-only SendKeys.

### ReplaceEvidence

Not implemented.

Desired behavior:

1. Clone template file.
2. Rename to `Excel_NAME`.
3. Replace specific cells with snap/log paths or embedded images.
4. Highlight target cell.
5. If `Amount > 1`, add additional snaps below existing content.
6. Add red rectangles to newly inserted evidence.
7. Update `isReplaced = 1`.

Hard part:

Automatic red rectangle placement. Possible approaches:

1. Use named ranges in the template. Most stable.
2. Use fixed cell anchors by phase and evidence type. Acceptable if template layout is stable.
3. Use OCR to locate text in screenshots, then calculate rectangle coordinates. Powerful but fragile and slower.
4. Use screenshot timestamp / file name / range metadata to locate target. Useful as a fallback, not enough by itself.

Recommendation:

Start with named ranges. Only add OCR after the deterministic path works.

## Suggested next implementation order

1. Test `VerifyTool.ps1 -Help` and `VerifyTool.ps1 -Phase Status`.
2. Test one old-compatible phase with `-DryRun`.
3. Test `ReviewEvidence` on one `TargetIds` value.
4. Rewrite ExcelSnap exact-range using Excel COM `CopyPicture()`.
5. Add DF snap.
6. Add ReplaceEvidence with named ranges.
7. Add GFIX LOD parser after real input samples are available.

## Safety / maintenance notes

- Keep switch values copied into plain bool variables before dot-source.
- Do not dot-source tool scripts that have `param(...)` blocks.
- Do not put workflow state into `Common.ps1`.
- Prefer `TargetIds` for small test runs.
- Back up mapping before scripts that update many rows.
- Keep child scripts directly runnable, but route daily work through `VerifyTool.ps1`.
