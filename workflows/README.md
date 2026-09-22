# workflows

Declarative ebi-dance workflow JSON belongs here. Workflows compose registered
steps into operations and follow `docs/ebi-dance/spec/WORKFLOW-SCHEMA.md`.

PowerShell implementations and environment-specific secrets do not belong
here; put capabilities in `modules/` and local values in ignored local config.

## spike.capture.json (P0-08)

The one workflow the P0 spike runs: `human.prepare` -> `browser.ensure`
(registers the Edge window as `mainWindow`) -> `screen.capture_window`
(saves `capture/spike/window.png` under the WorkDir). It has no `profile`
and no `each`; the P0-07 runner executes `setup`/`teardown` only. It is a
development artifact, not a project workflow -- P2-05 writes the first
real one. Run it on an office PC with

    .\ebi.ps1 run workflows\spike.capture.json -WorkDir C:\path\to\work

and a PNG of the Edge window is the P0-08 acceptance.
