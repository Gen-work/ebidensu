# ============================================================
# Locate-ByImage.ps1
#
# Template Matching tool.
# Finds a small image (Template) inside a larger image (Source) and
# returns the bounding box of the match, in pixels.
#
# No WinRT/UWP dependency. Uses plain .NET Framework (System.Drawing) and
# LockBits raw pointer access for a fast scan (tens of ms).
#
# Return value (PSCustomObject):
#   X, Y, Width, Height (pixel coordinates in the image)
#   $null when no match is found
#
# Usage (Standalone Test):
#   .\Locate-ByImage.ps1 -SourcePath "work\snap\GIFT_HM\JIDSC48S.png" -TemplatePath "work\anchor_hm.png"
#
# Usage (From Mark.ps1):
#   $box = .\Locate-ByImage.ps1 -SourcePath $png -TemplatePath $tpl -Quiet
#   if ($box) { $ptX = $box.X * 0.75 ... }
# ============================================================

param(
    [Parameter(Mandatory=$true)]
    [string]$SourcePath,

    [Parameter(Mandatory=$true)]
    [string]$TemplatePath,

    # Tolerance for small color differences (e.g. ClearType); 0 = exact match, 20 = practical slack
    [int]$Tolerance = 15,

    [switch]$Quiet
)

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Drawing -ErrorAction SilentlyContinue

# 1. Compile the C# execution class (fast pixel scan via LockBits)
$csharpCode = @"
using System;
using System.Drawing;
using System.Drawing.Imaging;
using System.Runtime.InteropServices;

namespace VerifyTool.ImageUtils
{
    public class Locator
    {
        public static int[] Find(string sourcePath, string templatePath, int tolerance)
        {
            using (Bitmap src = new Bitmap(sourcePath))
            using (Bitmap tpl = new Bitmap(templatePath))
            {
                if (tpl.Width > src.Width || tpl.Height > src.Height) return null;

                // Normalize to 32bppArgb (1 pixel = 4 bytes: B,G,R,A)
                BitmapData srcData = src.LockBits(new Rectangle(0, 0, src.Width, src.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                BitmapData tplData = tpl.LockBits(new Rectangle(0, 0, tpl.Width, tpl.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);

                int srcBytes = Math.Abs(srcData.Stride) * src.Height;
                byte[] srcValues = new byte[srcBytes];
                Marshal.Copy(srcData.Scan0, srcValues, 0, srcBytes);

                int tplBytes = Math.Abs(tplData.Stride) * tpl.Height;
                byte[] tplValues = new byte[tplBytes];
                Marshal.Copy(tplData.Scan0, tplValues, 0, tplBytes);

                src.UnlockBits(srcData);
                tpl.UnlockBits(tplData);

                int srcWidth = src.Width;
                int srcHeight = src.Height;
                int tplWidth = tpl.Width;
                int tplHeight = tpl.Height;

                for (int y = 0; y <= srcHeight - tplHeight; y++)
                {
                    for (int x = 0; x <= srcWidth - tplWidth; x++)
                    {
                        if (CheckMatch(srcValues, srcData.Stride, tplValues, tplData.Stride, x, y, tplWidth, tplHeight, tolerance))
                        {
                            return new int[] { x, y, tplWidth, tplHeight };
                        }
                    }
                }
                return null;
            }
        }

        private static bool CheckMatch(byte[] src, int srcStride, byte[] tpl, int tplStride, int startX, int startY, int tplWidth, int tplHeight, int tol)
        {
            for (int ty = 0; ty < tplHeight; ty++)
            {
                int srcRowOffset = (startY + ty) * srcStride;
                int tplRowOffset = ty * tplStride;

                for (int tx = 0; tx < tplWidth; tx++)
                {
                    int sIdx = srcRowOffset + (startX + tx) * 4;
                    int tIdx = tplRowOffset + tx * 4;

                    // Ignore the Alpha channel (Index + 3); compare only B/G/R
                    if (Math.Abs(src[sIdx] - tpl[tIdx]) > tol ||         // Blue
                        Math.Abs(src[sIdx + 1] - tpl[tIdx + 1]) > tol || // Green
                        Math.Abs(src[sIdx + 2] - tpl[tIdx + 2]) > tol)   // Red
                    {
                        return false;
                    }
                }
            }
            return true;
        }
    }
}
"@

if (-not ([System.Management.Automation.PSTypeName]'VerifyTool.ImageUtils.Locator').Type) {
    if (-not $Quiet) { Write-Host "[INFO] Compiling Image Locator Wrapper..." -ForegroundColor DarkGray }
    Add-Type -TypeDefinition $csharpCode -ReferencedAssemblies @("System.Drawing.dll") -ErrorAction Stop
}

# 2. Resolve paths and run
if (-not (Test-Path -LiteralPath $SourcePath))   { throw "Source image not found: $SourcePath" }
if (-not (Test-Path -LiteralPath $TemplatePath)) { throw "Template image not found: $TemplatePath" }

$srcFull = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
$tplFull = (Resolve-Path -LiteralPath $TemplatePath).ProviderPath

if (-not $Quiet) {
    Write-Host ("Source   : {0}" -f (Split-Path $srcFull -Leaf)) -ForegroundColor Cyan
    Write-Host ("Template : {0}" -f (Split-Path $tplFull -Leaf)) -ForegroundColor Cyan
    Write-Host ("Tolerance: {0}" -f $Tolerance) -ForegroundColor DarkGray
}

$boxResult = [VerifyTool.ImageUtils.Locator]::Find($srcFull, $tplFull, $Tolerance)

if ($null -ne $boxResult) {
    $resultObj = [PSCustomObject]@{
        X      = $boxResult[0]
        Y      = $boxResult[1]
        Width  = $boxResult[2]
        Height = $boxResult[3]
    }

    if (-not $Quiet) {
        Write-Host ("`n[SUCCESS] Found at Pixels: X={0}, Y={1}, W={2}, H={3}" -f `
            $resultObj.X, $resultObj.Y, $resultObj.Width, $resultObj.Height) -ForegroundColor Green
    }
    return $resultObj
} else {
    if (-not $Quiet) {
        Write-Host "`n[WARN] Template not found in source image." -ForegroundColor Yellow
    }
    return $null
}
