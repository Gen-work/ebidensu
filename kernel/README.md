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
  half-running them. Until `ebi.ps1` exists (P1-10) it is driven by hand:

  ```powershell
  . .\kernel\Runner.ps1
  Invoke-EbiWorkflow -Path .\workflows\x.json -WorkDir C:\work [-DryRun]
  ```

  The result is a hashtable (`ok`, `runId`, `steps`, `session`); the run's
  trace is at `<WorkDir>\run\<runId>\trace.jsonl`.

