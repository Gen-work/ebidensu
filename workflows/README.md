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

