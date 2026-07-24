#!/usr/bin/env node
// pixeldiff.mjs -- Phase-0 separability prototype for the ProcessTime
// old-snap 9->3 hand-verification plan (docs/ProcessTime-OldSnap-Verify-Plan.md,
// section 4.1). GO/NO-GO gate for D2's per-digit 3/9 image discrimination.
//
// The question this answers, entirely in the cloud/CI, BEFORE any office-PC
// GDI work:
//
//   Can a pixel metric reliably tell "this digit crop matches its OCR'd digit"
//   from "it does not (a 9 that OCR read as a 3, or vice versa)", with a
//   threshold that separates right-vs-right (accept) from right-vs-wrong
//   (reject) by a safe margin -- tolerant to anti-aliasing and small offsets?
//
// Method (the D2 comparison, prototyped): rasterize each digit glyph, then for
// a candidate crop score it against a '3' template and a '9' template rendered
// at the same size. The metric is scale/offset/AA tolerant:
//   grayscale ink -> binarize -> trim to bounding box -> average-pool to a
//   fixed normalized grid -> normalized cross-correlation (NCC) between grids.
// The winning template is argmax(NCC). A digit is "separable" when its NCC to
// the correct glyph beats its NCC to the confusable one by a margin.
//
// The PURE scoring functions (below, exported) are unit-tested with synthetic
// matrices via `node --test` (tests/pixeldiff.test.mjs) so the metric's
// behavior is pinned without a browser. The RENDER-and-report harness (main)
// uses the pre-installed Chromium to rasterize real font glyphs at a small,
// HM-cell-like size and prints the 3/9 separation margin + a GO/NO-GO verdict.
//
// FONT CAVEAT (same as gen.mjs): MS Gothic -- the exact font whose glyphs
// drive the real ja-OCR 9<->3 misread -- is Windows-only. A Linux/Chromium
// render falls back to another font, so this proves the METHOD separates a 3
// from a 9 (and how much AA/offset jitter it tolerates); the absolute margin
// on the office PC's MS Gothic captures must still be confirmed there. Pass
// --font "MS Gothic" on a Windows box (or point CHROMIUM at msedge) to render
// the production glyphs.
//
// Usage:
//   node pixeldiff.mjs [--font <family>] [--px <n>] [--grid WxH]
//                      [--jitter <n>] [--margin <f>] [--json]
//
//   --font    font-family for glyphs (default: a monospace stack)
//   --px      glyph pixel size        (default: 13, ~ the HM 10pt cell digit)
//   --grid    normalized grid WxH     (default: 16x24)
//   --jitter  sub-pixel offsets tried per glyph, to test AA/offset tolerance
//             (default: 3 -> offsets 0,.33,.66 px on each axis)
//   --margin  GO threshold on the worst 3/9 separation margin (default: 0.04)
//   --json    print the report object as JSON instead of a table
//
// Exit code: 0 on GO, 2 on NO-GO, 1 on error. (A gate MAY fail the build so a
// regression in separability is caught; wire into CI accordingly.)

import { dirname } from 'node:path';
import { fileURLToPath } from 'node:url';
import { existsSync, readdirSync } from 'node:fs';
import { join } from 'node:path';

const HERE = dirname(fileURLToPath(import.meta.url));

// ---------------------------------------------------------------------------
// PURE scoring functions (exported for node --test). An "image" here is a
// plain object { w, h, gray } where gray is a row-major array of 0..255
// luminance (0 = black ink, 255 = white paper), length w*h.
// ---------------------------------------------------------------------------

// Ink coverage in [0,1] per pixel: 1 = full black ink, 0 = white paper.
export function toInk(gray) {
  const ink = new Float64Array(gray.length);
  for (let i = 0; i < gray.length; i++) ink[i] = (255 - gray[i]) / 255;
  return ink;
}

// Bounding box of pixels whose ink exceeds `thr` (default 0.5). Returns null
// when the image is blank. Coordinates are inclusive-min / exclusive-max.
export function inkBBox(ink, w, h, thr = 0.5) {
  let x0 = w, y0 = h, x1 = 0, y1 = 0, any = false;
  for (let y = 0; y < h; y++) {
    for (let x = 0; x < w; x++) {
      if (ink[y * w + x] > thr) {
        any = true;
        if (x < x0) x0 = x;
        if (y < y0) y0 = y;
        if (x + 1 > x1) x1 = x + 1;
        if (y + 1 > y1) y1 = y + 1;
      }
    }
  }
  if (!any) return null;
  return { x0, y0, x1, y1 };
}

// Average-pool the ink inside `box` into a gw x gh normalized grid (row-major
// Float64Array). Average pooling (not nearest-neighbor) keeps anti-aliased
// edge coverage, which is what makes the metric AA-tolerant. A blank image
// (box null) yields an all-zero grid.
export function normalizeGrid(ink, w, h, box, gw = 16, gh = 24) {
  const out = new Float64Array(gw * gh);
  if (!box) return out;
  const bw = box.x1 - box.x0;
  const bh = box.y1 - box.y0;
  if (bw <= 0 || bh <= 0) return out;
  for (let cy = 0; cy < gh; cy++) {
    const sy0 = box.y0 + (cy / gh) * bh;
    const sy1 = box.y0 + ((cy + 1) / gh) * bh;
    for (let cx = 0; cx < gw; cx++) {
      const sx0 = box.x0 + (cx / gw) * bw;
      const sx1 = box.x0 + ((cx + 1) / gw) * bw;
      let sum = 0, n = 0;
      const iy0 = Math.floor(sy0), iy1 = Math.max(iy0 + 1, Math.ceil(sy1));
      const ix0 = Math.floor(sx0), ix1 = Math.max(ix0 + 1, Math.ceil(sx1));
      for (let y = iy0; y < iy1 && y < h; y++) {
        for (let x = ix0; x < ix1 && x < w; x++) {
          sum += ink[y * w + x];
          n++;
        }
      }
      out[cy * gw + cx] = n ? sum / n : 0;
    }
  }
  return out;
}

// Normalized cross-correlation of two equal-length vectors, in [-1, 1]. Mean-
// subtracted then norm-normalized, so it is invariant to overall ink amount
// and brightness -- it compares SHAPE. Two all-equal (zero-variance) vectors
// score 1 (identical); one zero-variance vs a varying one scores 0.
export function ncc(a, b) {
  const n = a.length;
  let ma = 0, mb = 0;
  for (let i = 0; i < n; i++) { ma += a[i]; mb += b[i]; }
  ma /= n; mb /= n;
  let num = 0, da = 0, db = 0;
  for (let i = 0; i < n; i++) {
    const va = a[i] - ma, vb = b[i] - mb;
    num += va * vb; da += va * va; db += vb * vb;
  }
  if (da === 0 && db === 0) return 1;      // both flat -> identical
  if (da === 0 || db === 0) return 0;      // one flat, one not -> unrelated
  return num / Math.sqrt(da * db);
}

// Shape similarity of two images in [-1, 1]: normalize both to the same grid,
// then NCC. Higher = more alike.
export function similarity(imgA, imgB, gw = 16, gh = 24, thr = 0.5) {
  const ia = toInk(imgA.gray), ib = toInk(imgB.gray);
  const ga = normalizeGrid(ia, imgA.w, imgA.h, inkBBox(ia, imgA.w, imgA.h, thr), gw, gh);
  const gb = normalizeGrid(ib, imgB.w, imgB.h, inkBBox(ib, imgB.w, imgB.h, thr), gw, gh);
  return ncc(ga, gb);
}

// Decide which of two templates a candidate matches, and by how much. Returns
// { pick: 'a'|'b', sa, sb, margin } where margin = |sa - sb| (>0 means a clear
// winner; a small margin means the two templates are nearly indistinguishable
// for this candidate -> the caller should refuse to auto-confirm).
export function classify(cand, tmplA, tmplB, gw = 16, gh = 24, thr = 0.5) {
  const sa = similarity(cand, tmplA, gw, gh, thr);
  const sb = similarity(cand, tmplB, gw, gh, thr);
  return { pick: sa >= sb ? 'a' : 'b', sa, sb, margin: Math.abs(sa - sb) };
}

// ---------------------------------------------------------------------------
// Render harness (Chromium). Not exercised by the unit tests.
// ---------------------------------------------------------------------------

function resolveChromium() {
  if (process.env.CHROMIUM_PATH && existsSync(process.env.CHROMIUM_PATH)) return process.env.CHROMIUM_PATH;
  const root = process.env.PLAYWRIGHT_BROWSERS_PATH || '/opt/pw-browsers';
  if (existsSync(root)) {
    const dir = readdirSync(root).filter(d => d.startsWith('chromium-')).sort().pop();
    if (dir) {
      const p = join(root, dir, 'chrome-linux', 'chrome');
      if (existsSync(p)) return p;
    }
  }
  return undefined;
}

function parseArgs(argv) {
  const a = { font: '"MS Gothic","MS Gothic",monospace', px: 13, gw: 16, gh: 24, jitter: 3, margin: 0.04, json: false };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '--font') a.font = argv[++i];
    else if (k === '--px') a.px = Number(argv[++i]);
    else if (k === '--grid') { const [w, h] = argv[++i].split('x').map(Number); a.gw = w; a.gh = h; }
    else if (k === '--jitter') a.jitter = Number(argv[++i]);
    else if (k === '--margin') a.margin = Number(argv[++i]);
    else if (k === '--json') a.json = true;
    else throw new Error(`unknown argument: ${k}`);
  }
  return a;
}

// Rasterize one digit to { w, h, gray } via an in-page canvas (the browser
// gives us pixel access directly, so no PNG decoder is needed in node).
async function renderDigit(page, digit, fontPx, fontFamily, dx, dy) {
  return await page.evaluate(({ digit, fontPx, fontFamily, dx, dy }) => {
    const pad = Math.ceil(fontPx * 0.6);
    const W = fontPx * 2 + pad * 2;
    const H = Math.ceil(fontPx * 1.6) + pad * 2;
    const c = document.createElement('canvas');
    c.width = W; c.height = H;
    const g = c.getContext('2d');
    g.fillStyle = '#fff'; g.fillRect(0, 0, W, H);
    g.fillStyle = '#000';
    g.textBaseline = 'top';
    g.font = `${fontPx}px ${fontFamily}`;
    g.fillText(String(digit), pad + dx, pad + dy);
    const img = g.getImageData(0, 0, W, H).data;
    const gray = new Array(W * H);
    for (let i = 0; i < W * H; i++) {
      // luminance (Rec.601); text is black on white so R==G==B, but average anyway
      gray[i] = Math.round(0.299 * img[i * 4] + 0.587 * img[i * 4 + 1] + 0.114 * img[i * 4 + 2]);
    }
    return { w: W, h: H, gray };
  }, { digit, fontPx, fontFamily, dx, dy });
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const { chromium } = await import('playwright-core');
  const browser = await chromium.launch({ executablePath: resolveChromium() });
  const report = {
    font: args.font, px: args.px, grid: `${args.gw}x${args.gh}`, jitter: args.jitter,
    goMargin: args.margin, samples: 0, pairs: {}, worst39: null, go: false,
  };
  try {
    const page = await browser.newPage({ deviceScaleFactor: 2 });
    await page.setContent('<!doctype html><meta charset=utf-8><body></body>', { waitUntil: 'load' });

    // Build a jittered template + candidate set for every digit.
    const offsets = [];
    for (let j = 0; j < args.jitter; j++) offsets.push((j / args.jitter));
    const digits = ['0','1','2','3','4','5','6','7','8','9'];
    const variants = {}; // digit -> array of images at different sub-pixel offsets
    for (const d of digits) {
      variants[d] = [];
      for (const oy of offsets) for (const ox of offsets) {
        variants[d].push(await renderDigit(page, d, args.px, args.font, ox, oy));
      }
    }

    // For the 3/9 decision that matters, and for every other digit pair as a
    // control, measure the separation margin: a candidate rendering of X, when
    // classified between template X and template Y (Y from a canonical, zero-
    // offset render), must pick X, and by a margin. We report the WORST-case
    // (minimum) correct margin over all jitter -- that is the number the GO
    // threshold guards.
    const canonical = {}; // digit -> the zero-offset template
    for (const d of digits) canonical[d] = variants[d][0];

    function pairStats(x, y) {
      let minMargin = Infinity, wrongPicks = 0, total = 0;
      for (const cand of variants[x]) {
        const r = classify(cand, canonical[x], canonical[y], args.gw, args.gh);
        total++;
        if (r.pick !== 'a') wrongPicks++;             // 'a' == template x
        const signed = r.sa - r.sb;                   // >0 means x preferred (correct)
        if (signed < minMargin) minMargin = signed;
      }
      return { minMargin, wrongPicks, total };
    }

    // 3-vs-9 both directions (the target decision).
    const s39 = pairStats('3', '9');
    const s93 = pairStats('9', '3');
    report.pairs['3vs9'] = s39;
    report.pairs['9vs3'] = s93;
    report.worst39 = Math.min(s39.minMargin, s93.minMargin);
    report.wrong39 = s39.wrongPicks + s93.wrongPicks;

    // Controls: a few other confusable pairs, to sanity-check the metric.
    for (const [x, y] of [['8','9'], ['3','8'], ['0','8'], ['5','6']]) {
      report.pairs[`${x}vs${y}`] = pairStats(x, y);
    }

    report.samples = digits.reduce((s, d) => s + variants[d].length, 0);
    report.go = report.wrong39 === 0 && report.worst39 >= report.goMargin;
  } finally {
    await browser.close();
  }

  if (args.json) {
    console.log(JSON.stringify(report, null, 2));
  } else {
    console.log(`\n=== Phase-0 pixel-diff separability (font=${report.font}, px=${report.px}, grid=${report.grid}, jitter=${report.jitter}) ===`);
    console.log(`samples per digit: ${report.jitter * report.jitter}, GO margin: ${report.goMargin}\n`);
    const rows = Object.entries(report.pairs);
    console.log(['pair','minCorrectMargin','wrongPicks/total'].map(s => s.padEnd(18)).join(' '));
    for (const [k, v] of rows) {
      console.log([k, v.minMargin.toFixed(4), `${v.wrongPicks}/${v.total}`].map(s => String(s).padEnd(18)).join(' '));
    }
    console.log('');
    console.log(`3<->9 worst-case correct margin: ${report.worst39.toFixed(4)}  (wrong picks: ${report.wrong39})`);
    console.log(`VERDICT: ${report.go ? 'GO -- 3/9 separable with margin' : 'NO-GO -- separation below threshold; ship D1 + deterministic checks only'}`);
  }
  return report.go ? 0 : 2;
}

if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  main().then(code => process.exit(code)).catch(e => { console.error('[pixeldiff] FAILED:', e.message); process.exit(1); });
}
