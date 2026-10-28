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
- `util.wideSnap.json` -- a horizontally long snap (10/28): `file.line_ends`
  reads the data file (line ends, widest line), `screen.stitch_scroll` opens
  it read-only in sakura and captures page by page to the right until that
  line's end mark is on a capture, then the book gets the merged PNG
  (outside-Excel merge) and/or the same pages lined up and grouped
  (inside-Excel merge), both through `excel.place_tiles`; after the save
  `verify.picture_pixels` reads the .xlsx to say whether any picture lost
  pixels. `-Var dataFile=... -Var book=... [-Var mode=both|stitched|tiles]`;
  the book is saved -- try it on a copy first.
- `util.wideSnap.capture.json` -- the same captures, nothing else: the page
  PNGs (`<file>__01.png` ..) and the merged `__wide.png` under
  `capture/wide`, for pasting by hand. Excel is not touched; the window
  keeps the size and place the viewer remembers (`-Var width=1906` sets
  the width only; `-Var close=false` leaves the viewer open after).
