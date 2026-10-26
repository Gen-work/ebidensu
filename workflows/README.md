# workflows

Declarative ebi-dance workflow JSON belongs here. Workflows compose registered
steps into operations and follow `docs/ebi-dance/spec/WORKFLOW-SCHEMA.md`.

PowerShell implementations and environment-specific secrets do not belong
here; put capabilities in `modules/` and local values in ignored local config.

## Files

- `spike.capture_window.json` -- the P0-08 end-to-end spike: `human.prepare`
  -> `browser.ensure` (registers `mainWindow`) -> `screen.capture_window`
  (writes `capture/spike/window.png` under the work dir). No profile, no
  `source`; it exists to prove the runner and the Session hand-off on a
  real machine. Run it as `kernel/README.md` shows.
- `gfixRecv.plan.json` / `gfixRecv.track.json` / `gfixRecv.logs.json` /
  `gfixRecv.evidence.json` -- the GFIX receive verification day on profile
  `gfix-recv` (morning plan, one track round per HOST job, end-of-day logs,
  the evidence workbook). Operator guide: `docs/gfix-recv/RUNBOOK.zh.md`.
- `kadou.collect.json` / `kadou.evidence.json` -- the 10/26 JJODWDBx operation
  check on profile `jod-kadou` (HOST JOD -> GFIX -> OPEN IDS / IGP2, nine
  transfers run by the verification environment every two hours): collect
  reads every round's GoAnywhere job logs (B) and the Jenkins receive folder
  + バッチ処理状況一覧 (C), evidence fills one book per job (GFIX受信結果 /
  GFIXデータ). Read-only. Operator guide: `docs/kadou/RUNBOOK.zh.md`.
