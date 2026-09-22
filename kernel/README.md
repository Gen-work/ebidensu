# kernel

Shared ebi-dance execution infrastructure belongs here: context evaluation,
step registration, workflow execution, tracing, gates, and generated docs.

Business- or system-specific automation does not belong here. Put reusable
capabilities in `modules/` and declarative orchestration in `workflows/`.

## What is here today

- `Trace.ps1` -- append-only `run/<runId>/trace.jsonl` (P0-03).
- `Runner.ps1` -- the P0-07 spike of the workflow runner, and the seed of
  P1-03/P1-04. It runs a workflow's `setup` and `teardown` sections, owns
  `$Ctx` (including `$Ctx.Session`) and the resource channel of
  `docs/ebi-dance/spec/STEP-CONTRACT.md` section 3.4 point 7, and refuses
  `source` / `each` / `{{...}}` / `when` / `onError` up front rather than
  half-running them. `ebi.ps1 run|dryrun <workflow> -WorkDir <dir>` wraps it
  (the P1-10 seed); it can also be driven by hand:

  ```powershell
  . .\kernel\Runner.ps1
  Invoke-EbiWorkflow -Path .\workflows\x.json -WorkDir C:\work [-DryRun]
  ```

  The result is a hashtable (`ok`, `runId`, `steps`, `session`); the run's
  trace is at `<WorkDir>\run\<runId>\trace.jsonl`.


- `Context.ps1` -- P1-01, pure `{{...}}` template evaluation per
  `docs/ebi-dance/spec/WORKFLOW-SCHEMA.md` section 4: `New-EbiTemplateScope`,
  `Resolve-EbiPath`, `Expand-EbiTemplate`, plus `Test-EbiTemplateString` /
  `Get-EbiTemplateReferences` for `ebi lint`. Failures are records naming the
  unresolved segment, never exceptions. Not wired into `Runner.ps1` yet (P1-03).
- `Key.ps1` -- key normalization shared by everything that renders an item's
  key (`Context.ps1` today; `table.load` and `table.key` later): full-width
  folding, the `" / "` display form, the `_`-joined file-safe form. The seed
  of P1-27; the one place a key comparison rule may live.
