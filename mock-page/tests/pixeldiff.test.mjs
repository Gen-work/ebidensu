// Unit tests for pixeldiff.mjs PURE scoring logic. Run: node --test (from mock-page/)
// No browser: synthetic little "images" pin the metric's behavior so the
// render harness's numbers are interpretable.
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { toInk, inkBBox, normalizeGrid, ncc, similarity, classify } from '../pixeldiff.mjs';

// Build an image object from a row-major 0/1 ink matrix (1 = black).
function img(rows) {
  const h = rows.length, w = rows[0].length;
  const gray = new Array(w * h);
  for (let y = 0; y < h; y++) for (let x = 0; x < w; x++) gray[y * w + x] = rows[y][x] ? 0 : 255;
  return { w, h, gray };
}

test('toInk: black -> 1, white -> 0', () => {
  const ink = toInk([0, 255, 128]);
  assert.equal(ink[0], 1);
  assert.equal(ink[1], 0);
  assert.ok(Math.abs(ink[2] - 0.498) < 0.01);
});

test('inkBBox: tight bounding box of the ink; null when blank', () => {
  const i = img([
    [0, 0, 0, 0],
    [0, 1, 1, 0],
    [0, 1, 1, 0],
    [0, 0, 0, 0],
  ]);
  const box = inkBBox(toInk(i.gray), i.w, i.h);
  assert.deepEqual(box, { x0: 1, y0: 1, x1: 3, y1: 3 });
  const blank = img([[0, 0], [0, 0]]);
  assert.equal(inkBBox(toInk(blank.gray), blank.w, blank.h), null);
});

test('normalizeGrid: a solid box maps to an all-ink grid regardless of size', () => {
  const i = img([
    [0, 0, 0, 0, 0],
    [0, 1, 1, 1, 0],
    [0, 1, 1, 1, 0],
    [0, 0, 0, 0, 0],
  ]);
  const ink = toInk(i.gray);
  const g = normalizeGrid(ink, i.w, i.h, inkBBox(ink, i.w, i.h), 3, 2);
  for (const v of g) assert.ok(v > 0.9, `cell ${v} should be ~full ink`);
});

test('ncc: identical vectors -> 1, opposite pattern -> negative', () => {
  assert.ok(Math.abs(ncc([1, 0, 1, 0], [1, 0, 1, 0]) - 1) < 1e-9);
  assert.ok(ncc([1, 0, 1, 0], [0, 1, 0, 1]) < -0.9);
  // A flat vector vs a varying one is treated as unrelated (0).
  assert.equal(ncc([1, 1, 1, 1], [1, 0, 1, 0]), 0);
  // Two flat vectors are identical (1).
  assert.equal(ncc([1, 1, 1], [1, 1, 1]), 1);
});

test('similarity: same glyph (scaled) ~ 1; a different pattern lower', () => {
  // An "L" shape and a scaled copy of it should be near-identical.
  const L = img([
    [1, 0, 0],
    [1, 0, 0],
    [1, 1, 1],
  ]);
  const Lbig = img([
    [1, 1, 0, 0, 0, 0],
    [1, 1, 0, 0, 0, 0],
    [1, 1, 0, 0, 0, 0],
    [1, 1, 0, 0, 0, 0],
    [1, 1, 1, 1, 1, 1],
    [1, 1, 1, 1, 1, 1],
  ]);
  const T = img([
    [1, 1, 1],
    [0, 1, 0],
    [0, 1, 0],
  ]);
  const sSame = similarity(L, Lbig, 6, 6);
  const sDiff = similarity(L, T, 6, 6);
  assert.ok(sSame > sDiff, `same-glyph similarity ${sSame} should beat cross-glyph ${sDiff}`);
  assert.ok(sSame > 0.8, `scaled same glyph should score high (${sSame})`);
});

test('classify: picks the closer template and reports a positive margin', () => {
  const three = img([
    [1, 1, 1],
    [0, 0, 1],
    [1, 1, 1],
    [0, 0, 1],
    [1, 1, 1],
  ]);
  const nine = img([
    [1, 1, 1],
    [1, 0, 1],
    [1, 1, 1],
    [0, 0, 1],
    [1, 1, 1],
  ]);
  // A slightly noisy copy of `three` must classify as `three` (template a).
  const threeNoisy = img([
    [1, 1, 1],
    [0, 0, 1],
    [1, 1, 1],
    [0, 0, 1],
    [1, 1, 0],
  ]);
  const r = classify(threeNoisy, three, nine, 5, 6);
  assert.equal(r.pick, 'a', `expected the 3 template; sa=${r.sa} sb=${r.sb}`);
  assert.ok(r.margin > 0, `margin ${r.margin} should be positive`);
});

test('classify: a blank candidate does not crash and yields a finite margin', () => {
  const blank = img([[0, 0, 0], [0, 0, 0], [0, 0, 0]]);
  const a = img([[1, 1, 1], [0, 0, 1], [1, 1, 1]]);
  const b = img([[1, 1, 1], [1, 0, 1], [1, 1, 1]]);
  const r = classify(blank, a, b, 4, 4);
  assert.ok(Number.isFinite(r.margin));
});
