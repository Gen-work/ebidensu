#!/usr/bin/env node
// compare.mjs -- OCR benchmark comparator.
//
// Scores an OCR read against the ground-truth manifest that gen.mjs emitted,
// field by field, and reports a 3<->9 confusion matrix plus per-field digit
// accuracy. Pure logic (string compare + counting) -- no OCR, no browser.
//
// The manifest is the correct answer (we generated the page). The "read" is
// what OCR returned; on an office PC that comes from Windows.Media.Ocr, here it
// is a hand-made fixture (samples/sample-ocr-read.json). Rows are matched by
// correlId.
//
// Usage:
//   node compare.mjs [--manifest <file>] [--read <file>] [--json]
//
//   --manifest  ground-truth manifest (default: out/sample-truth.manifest.json)
//   --read      OCR-read rows          (default: samples/sample-ocr-read.json)
//   --json      print the score object as JSON instead of a table
//
// Exit code 0 always (a benchmark reports, it does not fail the build); parse
// errors exit 1.

import { readFileSync } from 'node:fs';
import { dirname, join } from 'node:path';
import { fileURLToPath } from 'node:url';

const HERE = dirname(fileURLToPath(import.meta.url));

// Fields whose digits the OCR must read back (the 3/9-prone ones).
const FIELDS = ['start', 'end', 'duration', 'datestamp', 'count'];

function parseArgs(argv) {
  const a = {
    manifest: join(HERE, 'out', 'sample-truth.manifest.json'),
    read: join(HERE, 'samples', 'sample-ocr-read.json'),
    json: false,
  };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '--manifest') a.manifest = argv[++i];
    else if (k === '--read') a.read = argv[++i];
    else if (k === '--json') a.json = true;
    else throw new Error(`unknown argument: ${k}`);
  }
  return a;
}

const isDigit = (c) => c >= '0' && c <= '9';

// Pure digit comparison of one expected vs actual field string.
// Positional over the shared length; a length difference is flagged structural
// (its extra expected digits count as unmatched, but no misaligned swap tally).
// Returns { exact, digitTotal, digitMatched, swaps:{'9->3':n,...}, structural }.
export function compareDigits(expected, actual) {
  const exp = String(expected ?? '');
  const act = String(actual ?? '');
  const res = { exact: exp === act, digitTotal: 0, digitMatched: 0, swaps: {}, structural: exp.length !== act.length };
  for (let i = 0; i < exp.length; i++) {
    if (!isDigit(exp[i])) continue;
    res.digitTotal++;
    if (i < act.length && act[i] === exp[i]) {
      res.digitMatched++;
    } else if (i < act.length && isDigit(act[i])) {
      const key = `${exp[i]}->${act[i]}`;
      res.swaps[key] = (res.swaps[key] || 0) + 1;
    }
    // else: dropped / non-digit in that position -> counted as unmatched only
  }
  return res;
}

function blankScore() {
  return { rows: 0, exactRows: 0, digitTotal: 0, digitMatched: 0, swaps: {} };
}
function addSwaps(into, from) {
  for (const k of Object.keys(from)) into[k] = (into[k] || 0) + from[k];
}

export function scoreBenchmark(truthRows, readRows) {
  const readByCorrel = new Map();
  for (const r of readRows) if (r.correlId != null) readByCorrel.set(r.correlId, r);

  const perField = {};
  for (const f of FIELDS) perField[f] = blankScore();
  const overall = { rows: truthRows.length, matchedRows: 0, missingReads: [], extraReads: [], swaps: {} };

  const seen = new Set();
  for (const t of truthRows) {
    const r = readByCorrel.get(t.correlId);
    if (!r) { overall.missingReads.push(t.correlId); continue; }
    seen.add(t.correlId);
    overall.matchedRows++;
    for (const f of FIELDS) {
      const c = compareDigits(t[f], r[f]);
      const ps = perField[f];
      ps.rows++;
      if (c.exact) ps.exactRows++;
      ps.digitTotal += c.digitTotal;
      ps.digitMatched += c.digitMatched;
      addSwaps(ps.swaps, c.swaps);
      addSwaps(overall.swaps, c.swaps);
    }
  }
  for (const r of readRows) if (r.correlId != null && !seen.has(r.correlId)) overall.extraReads.push(r.correlId);

  // totals
  let dt = 0, dm = 0;
  for (const f of FIELDS) { dt += perField[f].digitTotal; dm += perField[f].digitMatched; }
  overall.digitTotal = dt;
  overall.digitMatched = dm;
  overall.digitAccuracy = dt ? dm / dt : 1;
  overall.confusion93 = overall.swaps['9->3'] || 0;
  overall.confusion39 = overall.swaps['3->9'] || 0;
  overall.otherSwaps = Object.entries(overall.swaps)
    .filter(([k]) => k !== '9->3' && k !== '3->9')
    .reduce((s, [, n]) => s + n, 0);

  return { perField, overall };
}

function pct(n, d) { return d ? (100 * n / d).toFixed(1) + '%' : '--'; }

function printReport(score, names) {
  const { perField, overall } = score;
  console.log(`\n=== OCR benchmark: ${names} ===`);
  console.log(`rows: ${overall.rows} truth, matched ${overall.matchedRows}` +
    (overall.missingReads.length ? `, missing read: ${overall.missingReads.join(',')}` : '') +
    (overall.extraReads.length ? `, extra read: ${overall.extraReads.join(',')}` : ''));
  console.log('');
  const hdr = ['field', 'exactRows', 'digitAcc', '9->3', '3->9', 'other'];
  const widths = [10, 10, 9, 5, 5, 6];
  const pad = (s, w) => String(s).padEnd(w);
  console.log(hdr.map((h, i) => pad(h, widths[i])).join(' '));
  for (const f of FIELDS) {
    const s = perField[f];
    const other = Object.entries(s.swaps).filter(([k]) => k !== '9->3' && k !== '3->9').reduce((a, [, n]) => a + n, 0);
    console.log([
      pad(f, widths[0]),
      pad(`${s.exactRows}/${s.rows}`, widths[1]),
      pad(pct(s.digitMatched, s.digitTotal), widths[2]),
      pad(s.swaps['9->3'] || 0, widths[3]),
      pad(s.swaps['3->9'] || 0, widths[4]),
      pad(other, widths[5]),
    ].join(' '));
  }
  console.log('');
  console.log(`OVERALL digit accuracy: ${overall.digitMatched}/${overall.digitTotal} (${pct(overall.digitMatched, overall.digitTotal)})`);
  console.log(`3<->9 confusion: 9->3 = ${overall.confusion93}, 3->9 = ${overall.confusion39}, other digit swaps = ${overall.otherSwaps}`);
}

function main() {
  const args = parseArgs(process.argv.slice(2));
  const manifest = JSON.parse(readFileSync(args.manifest, 'utf8'));
  const read = JSON.parse(readFileSync(args.read, 'utf8'));
  const score = scoreBenchmark(manifest.rows || [], read.rows || []);
  if (args.json) {
    console.log(JSON.stringify(score, null, 2));
  } else {
    printReport(score, `${manifest.png || args.manifest}`);
  }
}

// run only as a CLI (allow importing the pure functions for tests)
if (process.argv[1] && fileURLToPath(import.meta.url) === process.argv[1]) {
  try { main(); } catch (e) { console.error('[compare] FAILED:', e.message); process.exit(1); }
}
