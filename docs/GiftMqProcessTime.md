# GiftMqProcessTime -- fill the 処理時間 workbook from the GIFT MQ page

Standalone tool (`GiftMqProcessTime.ps1`, repo root). Not a VerifyTool
phase, not on the menu: it reads the operator's own `mapping.xlsx` and
writes the operator's own `処理時間(<Tag>).xlsx`, nothing else in the
work folder. v2.22.0.

## The hand procedure it replaces

Every day the leader posts which jobs run. The operator:

1. types `担当` / `GIFT実行日` / `GIFT TIME` into `mapping.xlsx` (sheet
   `mapping`, one row per job; the `ジョブ` column is the J-form name
   `CJODJCP1`, the `EXCEL` column the W-form `CJODWCP1`);
2. opens GIFT MQ > Transfer status > Inquiry in Edge and reads each
   job's **start** off the result LIST page (`Send date`), matching the
   row to the job by the scheduled time (a window: the real start is up
   to ~10 min early or late; the day's order settles the rest);
3. clicks each row's `Detail` button and reads the **end** off the detail
   page (`INSERTDATETIME`, the receive-log insertion stamp);
4. reads the **count** off the team's Teams chat
   (`ジョブ:QJODWCP1を実施します。(送信予定:338件)`);
5. types all of it into the `処理時間(BIX).xlsx` sheet:
   `No. / GIFT/GFIX / 相関ID / 開始日時 / 終了日時 / 処理時間 / 処理件数 / ジョブ`.

Steps 2, 3 and 5 are automated. Step 4 is automated as far as a pasted
text file goes (there is no Teams API access from the tool); otherwise the
operator types only the counts the end-of-run summary lists.

## Run

```powershell
# Edge: show the LIST page (Transfer status inquiry results) for the
# days you want, all records on ONE page (raise rows-per-page).
powershell -File GiftMqProcessTime.ps1 `
    -MappingXlsx C:\work\mapping.xlsx `
    -OutputXlsx  "C:\work\処理時間(BIX).xlsx" `
    -TeamsTextFile C:\work\teams.txt        # optional
```

1. It reads `mapping.xlsx` (COM, read-only) and keeps the rows with a
   `GIFT実行日`. With no `-FromDate`/`-ToDate` the page decides the window:
   only jobs scheduled on the days the page shows are in scope.
2. It stops at `Enter=OK / q=quit :` and waits. That prompt is a `Read-Host`
   **in the PowerShell window**, not in Edge: show the LIST page in Edge, then
   click the PowerShell window and press Enter there. It captures the page once
   (Ctrl+A/Ctrl+C, `Read-PageText.ps1`) and archives the text under
   `<output dir>\giftmq_text\list_<stamp>.txt` (reuse with
   `-PageTextFile`, e.g. to rerun without Edge).
3. It prints the match table (job / owner / scheduled / status / page No /
   send date). `ambiguous` (two candidates in the window) and `notime`
   (date but no time in the mapping) stop for a pick: `Enter` accepts the
   nearest, a page `No` picks that record, `s` skips.
4. For every row whose end time is still blank it opens the Detail page by
   keyboard, reads `INSERTDATETIME`, and comes back. **Every detail page
   is verified** against its record (`CORRELID(CHAR)` and `SENDDATETIME`
   to the second) before the value is trusted. See "Detail navigation".
5. It writes the sheet and saves. Then it lists what is still blank: end
   times it could not read, counts it had no Teams text for, scheduled
   jobs with no page record.

Switches: `-NoDetail` (start times only, no Edge clicks after the capture),
`-DryRun` (print the plan, write nothing), `-Force` (rewrite filled cells),
`-NonInteractive` (nearest wins, no prompts), `-Owner げ`, `-Jobs A,B`,
`-CorrelId JJPCRS12`, `-ToleranceMinutes 10`.

## What it writes

| col | value | source |
|-----|-------|--------|
| A `No.` | row-1 | only on appended rows |
| B | `GIFT` / `GFIX` | only on appended rows (the GFIX row is created blank) |
| C `相関ID` | the record's `Correlid` | LIST page |
| D `開始日時` | `Send date`, as a real Excel date/time (`yyyy/mm/dd hh:mm:ss`) | LIST page |
| E `終了日時` | `INSERTDATETIME` truncated to the second | Detail page |
| F `処理時間` | `=E-D` (kept / recreated on appended rows) | formula |
| G `処理件数` | `<n>件` (comma from 10,000) | Teams text, when given |
| H `ジョブ` | the mapping's `ジョブ` | mapping |

Existing GIFT rows (col H = job, col B = `GIFT`) are **updated, blank
cells only**. A job with no row gets a GIFT+GFIX pair: into the sheet's
spare placeholder rows (C..E, G, H blank) when it has them, else after
the last row with the previous pair's formats copied. The n-th match of a
job on the same day pairs with its n-th existing row (a job the leader
lists twice runs twice). The GFIX side is never touched.

Two conventions to know when reading a mixed sheet:

- rows typed by hand from the HM page carry a start about one second
  later than the MQ `Send date` (HM stamps its own start); the tool does
  not rewrite a filled start, so old and new rows can differ by a second;
- the end time is the receive-log insertion stamp, typically 3-8 s after
  the send, where the HM page's own end was ~3 s. The tool records what
  the MQ Detail says; the meaning of `処理時間` is the operator's call.

## Detail navigation

The list's `Detail` buttons are the only focusable controls in the table,
in page order. Two ways to reach record N's button, both keyboard-only:

- **Find** (default): `Ctrl+F` the record's own `Send date` text, `Esc`,
  `Tab`, `Enter`. Edge leaves the focus starting point at the find match,
  so the next `Tab` is that row's button.
- **Tab**: `Ctrl+F` the page title (`-AnchorText`), `Esc`, `Tab` N times
  (+ `-TabsBeforeFirstDetail`, the controls before the first button, e.g.
  a `Back` button; `-1` = auto: 0, 1, 2 are tried on the first record and
  the one that verifies is kept), `Enter`.

If the page reached is not the record's (the verification above), the
other method is tried, then the operator is asked (`r`/`s`/`q`). Leaving
the detail page: the page's own `Back` button (`Ctrl+F` title, `Tab`,
`Enter`; `-BackMethod Button`) or `Alt+Left` (`-BackMethod AltLeft`).
Every detail text reached is archived as
`giftmq_text\detail_<No>_<job>.txt`.

## Teams text

Paste the chat into a UTF-8 text file. Lines shaped
`ジョブ:<name> ... (送信予定:<n>件)` are picked up (ASCII or full-width
colon; a count on the next line attaches to the last job named). The W
name is mapped to the J job by replacing the 5th character. The LAST
mention of a job wins, so a corrected re-post overrides.

## Files

- `GiftMqProcessTime.ps1` -- the driver (COM + SendKeys glue; has
  `param()`, call via `&`/`-File`, never dot-source).
- `modules/verify/GiftMqProcessTime.ps1` -- pure library, no COM/UI:
  `ConvertFrom-GiftMqListText`, `ConvertFrom-GiftMqDetailText`,
  `Get-GiftMqDetailEndTime`, `Test-GiftMqDetailMatchesRecord`,
  `ConvertTo-GiftMqScheduledTime`, `Get-GiftMqMappingColumns`,
  `ConvertTo-GiftMqSchedules`, `Resolve-GiftMqJobMatches`,
  `ConvertFrom-GiftMqTeamsText`, `Get-GiftMqOutputPlan`,
  `Format-GiftMqStamp` / `Format-GiftMqCount`, `Get-GiftMqDetailTabCount`.
  Unit-tested: `Tests\Test-GiftMqProcessTime.ps1`.

## If pressing Enter seems to do nothing

The run drives a real mouse. Before v2.22.2, `Switch-ToEdge`'s synthetic
Alt+Tab could fail to land (Windows ignores it under some policies), and the
next click then went into the **PowerShell console** instead of Edge. Windows
consoles ship with QuickEdit mode on, where a click starts a text selection,
and a console with an active selection **blocks the process's output**. The
run looked frozen with no error, forever.

Two guards now make that impossible:

- the click is made only when an Edge window is in the foreground, and
  `Confirm-GiftMqEdgeForeground` checks that after every switch, asking the
  operator to click Edge when it fails rather than proceeding blind;
- QuickEdit is turned off for the duration of the run and restored on exit.

The page-text poll also prints one line per attempt (`read LIST page (try 1):
4821 chars, not the page yet`), so a silent console now means a real hang, not
a slow poll. **If an older copy of the script does freeze: press `Esc` or
right-click inside the console window — output resumes immediately if a
selection was the cause.**

The fallback that needs no automation at all: copy the page yourself in Edge
(Ctrl+A, Ctrl+C), paste into a UTF-8 text file, and pass it with
`-PageTextFile`. With `-NoDetail` that fills every start time and touches
neither Edge nor the mouse.

## Console encoding

The script does **not** touch `[Console]::OutputEncoding`. On the office PC the
console runs the JP codepage (932); forcing UTF-8 output there turns every
non-ASCII byte printed into mojibake. The first real run showed the output
workbook's own path as `C:\Users\...\<garbage>BIX.xlsx` while having opened the
right file, which reads like a failure and is not one. If a path still looks
garbled, check the file name itself before assuming the tool is wrong.

## Not verified yet (office PC)

Authored without Windows/Excel/Edge. Confirm, in this order:

1. `-DryRun -NoDetail` -- the mapping columns are found, the match table
   is right for a known day.
2. `-NoDetail` -- start times land in the right rows, appended pairs look
   like the hand-typed ones.
3. one day with details -- the first Detail click verifies (watch the
   `nav Find: ...` lines; if Find never verifies, pass `-DetailNav Tab`),
   the Back button returns to the LIST, the end times match what you read
   by hand.
4. `-TeamsTextFile` with one day's pasted chat.
