# kernel

Shared ebi-dance execution infrastructure belongs here: context evaluation,
step registration, workflow execution, tracing, gates, and generated docs.

Business- or system-specific automation does not belong here. Put reusable
capabilities in `modules/` and declarative orchestration in `workflows/`.

## Files

| File | Role |
|------|------|
| `Trace.ps1` | Append-only `run/<runId>/trace.jsonl` writer/reader (P0-03). |
| `Runner.ps1` | P0-07 spike runner: `Invoke-EbiWorkflow` runs `setup`/`teardown`, owns `$Ctx.Session` (as / resource / release), traces each step, guarantees teardown on in-process exits. Grows into the real runner in P1-03/P1-04. |
| `Win32.ps1` | Lazily compiled user32 P/Invoke (`Get-EbiWin32`, `Test-EbiWindowHandle`, `Get-EbiWindowRect`, `Set-EbiForegroundWindow`) for window-facing steps. |

All three are dot-source libraries: no `param()`, ASCII source, no `class`.
