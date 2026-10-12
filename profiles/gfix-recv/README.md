# profiles/gfix-recv

GIFT -> GFIX migration, **receive side on OPEN** (J4 verification): for each
job the HOST team runs, check that the file GoAnywhere received and handed to
the GFIX receive batch is the same as the GIFT-era file, tell the HOST team
in Teams, then build the evidence workbook. The first profile written for a
new job on the ebi-dance engine (P3) -- read off the operator's own
description, the sample evidence workbooks and the real page texts.

Workflows that use it (`workflows/gfixRecv.*.json`):

| workflow | when | what |
|----------|------|------|
| `gfixRecv.plan` | morning | today's jobs = WBS (受信 / 修正後実施 / 担当 厳, arrows as in OwnerFilter) ∪ mapping.xlsx (GFIX実行日 = today) -> `gfixrecv.csv` |
| `gfixRecv.track` | each HOST job | wait for the Teams "正常終了", refresh GoAnywhere, find the job's rows, Jenkins file, download to DATA\GFIX\<W-job>, pair with DATA\GIFT\<J-job>, compare, DF captures, Teams text + pictures on the clipboard |
| `gfixRecv.logs` | end of day | GFIXReceive.log (+ Unzip) -> log\GFIXReceive\mmdd.log (SJIS), GoAnywhere / Jenkins overview texts, each Receive job's log -> log\GFIX受信ログ |
| `gfixRecv.evidence` | after logs | GFIX受信結果 (Excel snap copy, job log + yellow, GFIXReceive.log block + yellow, Jenkins snap + red box), GIFTデータvsGFIXデータ (B3, DF captures, wave, red boxes), tidy + save |

Files:

- `pages.json` -- GoAnywhere 完了したジョブ (window title, refresh key recipe,
  job-log download recipe, Teams crop geometry), Jenkins data/report (refresh
  recipe, received-file name pattern, box geometry), the downloaded job log.
- `grammar.json` -- the three page/file texts; fixtures under `fixtures/` are
  the operator's real samples (`ebi profile check gfix-recv` is green on them).
- `layout.json` -- the evidence workbook: sheet names, labels, MS Gothic 10,
  highlight patterns (width = ceil(chars / 3) columns, measured on the
  sample), DF window 1133 x 429 (= 16 lines, the sample captures) and its two
  boxes, the wave separator (`assets/wave.png`).
- `paths.json` -- **placeholders only**. The real share folder, WBS path,
  URLs go in `<WorkDir>\ebi.local.json` (see `docs/gfix-recv/RUNBOOK.zh.md`);
  the mask gate refuses UNC paths / personal folders in the repo.
- `worklist.json` -- `gfixrecv.csv`, one row per Excel_NAME (the W-name),
  verdict columns `track` / `logs` / `evidence`.

Measured, not guessed (from the sample workbook QJDSWM39 and the operator's
screenshots): Teams crop (GoAnywhere maximized, 1920 x 1080), DF line-number
strip and status-bar cell, Jenkins row pitch / box width, highlight widths.
Still to confirm on the office PC on the first run: the GoAnywhere refresh
recipe (Ctrl+F -> Esc -> Tab x 5 -> Enter), the job-log download recipe, and
whether a rich (HTML) paste lands in Teams with its pictures.
