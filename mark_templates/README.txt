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
     Tolerance : LockBits color tolerance for this box only (default: Mark.ImageMatch.Tolerance, 15)
     PadX/PadY : pixels added around the matched region on each axis (default: 0)
5. Run `.\VerifyTool.ps1 -Phase MarkGift` (or MarkGfix/MarkDf) and check the
   console: a line tagged [MARK-IMG] means the template matched; [MARK] means
   it fell back to the fixed offset (check the [WARN] line above it for why).

This needs a real Windows + Excel session with real HM/MQ/Jenkins screenshots
to calibrate -- there is no Windows/Excel in the tool's dev environment, so
this folder ships empty. See CLAUDE.md's TODO list for background.

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
