mark_templates/
================

Reference images for Mark.ps1's OPTIONAL image-recognition box placement.

Background
----------
By default, Mark.ps1 draws each red rectangle at a fixed pixel offset from
the picture's top-left corner (Mark.Boxes.<folder>[].OffsetX/OffsetY in
VerifyConfig.psd1). That works only as long as every screenshot in a folder
has the target field in exactly the same place. If the captured page shifts
(window resize, scroll position, a slightly different HM/MQ/Jenkins layout,
the target row landing on row 2 instead of row 1), the fixed-offset box ends
up in the wrong place.

Adding a "Template" key to a Mark.Boxes entry makes Mark.ps1 try to LOCATE
the target on the actual screenshot first (via Locate-ByImage.ps1, LockBits
template matching), and only falls back to the fixed offset when no
template is configured, the template file is missing, or no match is found.
This degrades gracefully -- Mark never blocks on a missing/failed template.

Template works on BOTH box kinds (since v2.10.8): fixed-offset boxes
(OffsetX/OffsetY) fall back to the offset, and cell-range boxes
(CellCols/RowsFromBottom, e.g. the DF folder) fall back to the cell-range
placement. The DF same-content button is the motivating case: df.exe is
captured in 'window' mode, so the button's position moves with however the
operator sized the window -- only a template match can track it. Example:

     Mark.Boxes.DF = @( @{ CellCols = 'AW:BC'; RowsFromBottom = 2; Template = 'DfSame.png' } )

How to create a template
-------------------------
1. Open one representative screenshot from the folder you want to calibrate,
   e.g. <WorkDir>\snap\GIFT_HM\<some correl>.png.
2. Crop out a SMALL, VISUALLY DISTINCTIVE region that reliably identifies the
   target field (e.g. just the status cell's background + a few characters
   of its label -- not the whole row, and not a region that repeats
   elsewhere on the page). Save it as a PNG here, e.g. gift_hm_status.png.
   Keep it small: LockBits matching is a pixel-for-pixel scan, so a smaller,
   more unique template matches faster and more reliably than a large one.
3. Reference it from VerifyConfig.psd1 (or verify_config.json):

     Mark.Boxes.GIFT_HM = @( @{ Template = 'gift_hm_status.png'; Width = 62.2; Height = 16.5 } )

   Template filenames are resolved against Mark.TemplateDir first (if set),
   then this folder (mark_templates\), then as an absolute/relative path.
   Width/Height/OffsetX/OffsetY on the same entry are the FALLBACK box used
   only when the template match fails -- keep them as a safety net.
4. Optional per-box overrides:
     Tolerance         : LockBits color tolerance for this box only (default: Mark.ImageMatch.Tolerance, 15)
     PadX/PadY         : shifts the matched anchor point on each axis (default: 0)
     PadWidth/PadHeight: constant extra size added on top of Width/Height (or,
                         when the box has no Width/Height, on top of the
                         matched crop's own size) (default: 0)

   PadWidth/PadHeight are a FIXED number added to every correl in the folder
   -- they do NOT make the box track each correl's actual on-page content
   length (e.g. a Jenkins file-list entry whose filename is longer some runs
   than others). Use them only to add a constant safety margin. A real
   per-correl auto-size needs a measured source (SnapVerify's M5 pixel
   localisation loc.json, or an OCR-based text-width measurement like
   GfixLog.AutoHighlightWidth uses for the GFIX log highlight) -- not yet
   wired into plain Template boxes.
5. Run `.\VerifyTool.ps1 -Phase MarkGift` (or MarkGfix/MarkDf) and check the
   console: a line tagged [MARK-IMG] means the template matched; [MARK] means
   it fell back to the fixed offset (check the [WARN] line above it for why).

This needs a real Windows + Excel session with real HM/MQ/Jenkins screenshots
to calibrate -- there is no Windows/Excel in the tool's dev environment, so
this folder ships empty. See CLAUDE.md's TODO list for background.

StampImage (image-recognition-only stamp, no fixed-offset fallback)
---------------------------------------------------------------------
Add a "StampImage" key ALONGSIDE "Template" on a Boxes entry to insert a
whole image (native size) at the Template match location instead of drawing
a rectangle. Unlike the plain Template box above, StampImage has NO
OffsetX/OffsetY fallback: when the template does not match, nothing is
inserted at all -- "no match" IS the answer (the target pattern genuinely
isn't there), not a calibration miss to paper over with a fixed guess.

This is the mechanism wired for GIFT_noGfixfile (F4, "no GFIX file expected"
past-data check):

    Mark.Boxes.GIFT_noGfixfile = @( @{ Template = 'NoGfixHit.png'; StampImage = 'already_exists.png' } )

1. NoGfixHit.png: crop a SMALL, visually distinctive region from a real
   GIFT_noGfixfile snap screenshot (<WorkDir>\snap\GIFT_noGfixfile\<correl>.png)
   that ONLY appears when a past-data file was actually found on the Jenkins
   list page (e.g. the file-list row / reference marker itself) -- the same
   "small and unique" rule as any other Template crop. On the normal case
   (no file found -- the expected, OK outcome) this pattern is simply absent
   from the screenshot, so Locate-ByImage naturally returns no match and no
   stamp is drawn; that is the correct behavior, not a failure.
2. already_exists.png: the stamp image inserted (as-is, no scaling) at the
   matched location when NoGfixHit.png IS found. This is the "past-data
   exists" flag the operator sees on the evidence sheet.
3. Both files go directly in this folder (or Mark.TemplateDir if set) --
   same filename resolution as any other Template/StampImage value.
4. Run `.\VerifyTool.ps1 -Phase MarkGift` and check the console:
   [STAMP-IMG] = matched and stamped; [SKIP-STAMP] = no match, nothing drawn
   (expected for every correl where no past-data file exists); [WARN] means
   StampImage itself was not found even though the Template matched.
5. Per-box Tolerance/PadX/PadY overrides work the same as plain Template
   boxes (step 4 above).

This is independent of -- and does not require -- Mark.NoteStamps below: it
runs directly against the source snap PNG via Locate-ByImage, with no
dependency on SnapVerify.Localize being enabled or a .note.json sidecar
existing. Needs a real Windows + Excel session with a real past-data hit
screenshot to calibrate NoGfixHit.png; ships without either PNG here.

NoteStamps (verifyNote annotation stamp images)
------------------------------------------------
This folder also holds stamp images for Mark.NoteStamps -- a SEPARATE opt-in
feature from the Template box placement above. Instead of drawing a red
rectangle, it inserts a whole image (e.g. already_exists.png) next to a
'verifyNote' annotation (currently only the F4/M6 GIFT_noGfixfile past-data
hit). It reuses the pixel rect already carried in that annotation's payload
(from the snap-time <correl>.loc.json / .note.json sidecars) to find the
highlighted Jenkins row -- no separate template match, no re-scanning the
source PNG. Configure it in VerifyConfig.psd1 (or verify_config.json):

    Mark.NoteStamps.GIFT_noGfixfile = @{ Image = 'already_exists.png'; Column = 'AF'; RowOffset = 0 }

Image filenames resolve the same way Template does (Mark.TemplateDir first,
then this folder). RowOffset shifts from the highlighted row (0 = same row).
Drop already_exists.png here to enable it; leaving Mark.NoteStamps empty (or
the key unset) disables the stamp with no other effect on Mark.
