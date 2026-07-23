// Unit tests for compare.mjs pure logic. Run: node --test  (from mock-page/)
import { test } from 'node:test';
import assert from 'node:assert/strict';
import { compareDigits, scoreBenchmark } from '../compare.mjs';

test('compareDigits: exact match, all digits counted', () => {
  const c = compareDigits('03:59:07', '03:59:07');
  assert.equal(c.exact, true);
  assert.equal(c.digitTotal, 6);
  assert.equal(c.digitMatched, 6);
  assert.deepEqual(c.swaps, {});
  assert.equal(c.structural, false);
});

test('compareDigits: a single 9->3 misread', () => {
  const c = compareDigits('03:59:07', '03:53:07');
  assert.equal(c.exact, false);
  assert.equal(c.digitMatched, 5);
  assert.equal(c.swaps['9->3'], 1);
});

test('compareDigits: reverse 3->9', () => {
  const c = compareDigits('136,045', '196,045');
  assert.equal(c.swaps['3->9'], 1);
  assert.equal(c.swaps['9->3'], undefined);
});

test('compareDigits: multiple swaps in one field', () => {
  const c = compareDigits('20260723035907', '20260723035307');
  assert.equal(c.swaps['9->3'], 1);
  const c2 = compareDigits('999', '333');
  assert.equal(c2.swaps['9->3'], 3);
});

test('compareDigits: non-digit chars are ignored, only digits tallied', () => {
  const c = compareDigits('2026/07/23', '2026/07/23');
  assert.equal(c.digitTotal, 8); // 2026 07 23
  assert.equal(c.digitMatched, 8);
});

test('compareDigits: length mismatch flagged structural, no misaligned swaps', () => {
  const c = compareDigits('12:00:00', '12:0:00'); // OCR dropped a char
  assert.equal(c.structural, true);
  // must not invent a big swap tally from misalignment
  const swapTotal = Object.values(c.swaps).reduce((a, n) => a + n, 0);
  assert.ok(swapTotal <= 2, `unexpected swap tally ${swapTotal}`);
});

test('compareDigits: blank expected and blank actual is exact, zero digits', () => {
  const c = compareDigits('', '');
  assert.equal(c.exact, true);
  assert.equal(c.digitTotal, 0);
});

test('scoreBenchmark: aggregates the confusion matrix across rows and fields', () => {
  const truth = [
    { correlId: 'A', start: '03:59:07', end: '03:59:52', duration: '00:00:45', datestamp: '20260723035907', count: '59,476' },
    { correlId: 'B', start: '03:54:00', end: '03:55:34', duration: '00:01:34', datestamp: '20260723035400', count: '136,045' },
  ];
  const read = [
    { correlId: 'A', start: '03:53:07', end: '03:53:52', duration: '00:00:45', datestamp: '20260723035307', count: '53,476' }, // 4x 9->3
    { correlId: 'B', start: '03:54:00', end: '03:55:34', duration: '00:01:34', datestamp: '20260723035400', count: '196,045' }, // 1x 3->9
  ];
  const { overall } = scoreBenchmark(truth, read);
  assert.equal(overall.matchedRows, 2);
  assert.equal(overall.confusion93, 4);
  assert.equal(overall.confusion39, 1);
  assert.equal(overall.otherSwaps, 0);
});

test('scoreBenchmark: reports missing and extra reads by correlId', () => {
  const truth = [{ correlId: 'A', start: '1', end: '1', duration: '1', datestamp: '', count: '0' }];
  const read = [{ correlId: 'Z', start: '1', end: '1', duration: '1', datestamp: '', count: '0' }];
  const { overall } = scoreBenchmark(truth, read);
  assert.deepEqual(overall.missingReads, ['A']);
  assert.deepEqual(overall.extraReads, ['Z']);
  assert.equal(overall.matchedRows, 0);
});

test('scoreBenchmark: perfect read = 100% accuracy, empty confusion', () => {
  const rows = [{ correlId: 'A', start: '03:59:07', end: '03:59:52', duration: '00:00:45', datestamp: '20260723035907', count: '59,476' }];
  const { overall } = scoreBenchmark(rows, rows.map(r => ({ ...r })));
  assert.equal(overall.digitAccuracy, 1);
  assert.equal(overall.confusion93, 0);
  assert.equal(overall.confusion39, 0);
});
