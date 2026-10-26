# profiles/jod-kadou

JJODWDBx operation check (稼働確認, J4, 10/26): the verification
environment runs JJODWDBF / JJODWDBH / JJODWDBM on its own every two hours
(8:00-18:00); each sends three TEXT files HOST (JOD) -> GFIX -> OPEN
(IDS IF5050, IGP2 IF5133 / IF5141). This profile collects the OPEN-side
evidence only -- B = GoAnywhere job logs, C = the Jenkins receive folder and
the バッチ処理状況一覧 page; the send side is the HOST team's. Read-only
throughout. Operator guide: `docs/kadou/RUNBOOK.zh.md`.

| workflow | what |
|----------|------|
| `kadou.collect` | per I/F: every job log of today (all rounds) from GoAnywhere; per transfer: verify.job_logs (upload / rc 0 / delete / end, shell + SS), Jenkins listing + Ctrl+F screenshot, the received file (GET), バッチ処理状況一覧 screenshot after the operator's search, one gate |
| `kadou.evidence` | one book per job from `assets/kadou-template.xlsx`: GFIX受信結果 (batch-status picture, ▼GFIXログ + yellow Command line, Jenkins picture) and GFIXデータ (the received file's text) |

Files:

- `worklist.json` -- `kadou.csv`, one row per transfer (相関ID, the
  verification-environment ids ending in S); `targets` = the nine rows with
  their I/F facts and the rows they take in the job book (reference DBH
  layout). Columns B / C_JK / C_BS are the 9 x B / C summary table.
- `pages.json` -- GoAnywhere (the gfix-recv page, other projects), the
  Jenkins receive folder, バッチ処理状況一覧 (its window title is a guess).
- `grammar.json` / `rules.json` -- the GoAnywhere list and the job log;
  fixtures under `fixtures/` (`ebi profile check jod-kadou` is green).
- `layout.json` -- the book name, sheet names, labels, fonts, highlight.
- `paths.json` -- placeholders only; the real URLs go in
  `<WorkDir>/ebi.local.json`.
