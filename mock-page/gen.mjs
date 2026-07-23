#!/usr/bin/env node
// gen.mjs -- mock HM batch-status page generator.
//
// Renders the HM バッチ処理状況一覧 template (templates/hm-batch-status.html)
// filled with a ground-truth row set, screenshots it to a PNG with Chromium
// (Playwright), and writes a manifest.json pairing the PNG with the exact
// values the OCR must read back. Runs fully on Linux/CI -- no office PC.
//
// Usage:
//   node gen.mjs [--truth <file>] [--out <dir>] [--name <stem>] [--scale <n>] [--full]
//
//   --truth  ground-truth JSON (default: samples/sample-truth.json)
//   --out    output directory   (default: out)
//   --name   output file stem   (default: truth file basename)
//   --scale  device scale factor, i.e. render DPI (default: 2)
//   --full   screenshot the whole page instead of a tight crop of the table
//
// Output: <out>/<name>.png  +  <out>/<name>.manifest.json

import { readFileSync, writeFileSync, mkdirSync, existsSync, readdirSync } from 'node:fs';
import { dirname, join, basename } from 'node:path';
import { fileURLToPath } from 'node:url';
import { chromium } from 'playwright-core';

const HERE = dirname(fileURLToPath(import.meta.url));

function parseArgs(argv) {
  const a = { truth: join(HERE, 'samples', 'sample-truth.json'), out: join(HERE, 'out'),
              name: null, scale: 2, full: false };
  for (let i = 0; i < argv.length; i++) {
    const k = argv[i];
    if (k === '--truth') a.truth = argv[++i];
    else if (k === '--out') a.out = argv[++i];
    else if (k === '--name') a.name = argv[++i];
    else if (k === '--scale') a.scale = Number(argv[++i]);
    else if (k === '--full') a.full = true;
    else throw new Error(`unknown argument: ${k}`);
  }
  return a;
}

// Resolve the pre-installed Chromium (no download; the env blocks the CDN).
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
  return undefined; // let Playwright try its own resolution
}

function esc(s) {
  return String(s ?? '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

const STATUS_TEXT = { normal: '正常終了', abend: '異常終了' };

function buildRowsHtml(rows) {
  return rows.map((r, i) => {
    const statusKey = (r.status || 'normal').toLowerCase();
    const statusText = STATUS_TEXT[statusKey] || esc(r.status);
    const statusCls = statusKey === 'abend' ? 'status-ng' : 'status-ok';
    return [
      '        <tr>',
      `          <td class="num">${i + 1}</td>`,
      `          <td>${esc(r.correlId)}</td>`,
      `          <td>${esc(r.system)}</td>`,
      `          <td>${esc(r.jobName)}</td>`,
      `          <td class="dt">${esc(r.start)}</td>`,
      `          <td class="dt">${esc(r.end)}</td>`,
      `          <td class="${statusCls}">${statusText}</td>`,
      `          <td class="dt">${esc(r.datestamp)}</td>`,
      `          <td class="num">${esc(r.count)}</td>`,
      `          <td class="num">${esc(r.rc)}</td>`,
      '        </tr>',
    ].join('\n');
  }).join('\n');
}

async function main() {
  const args = parseArgs(process.argv.slice(2));
  const truth = JSON.parse(readFileSync(args.truth, 'utf8'));
  const rows = truth.rows || [];
  const stem = args.name || basename(args.truth).replace(/\.json$/i, '');

  const template = readFileSync(join(HERE, 'templates', 'hm-batch-status.html'), 'utf8');
  // replaceAll: the template's header comment also mentions these tokens, and
  // a single .replace() would consume that first mention instead of the body.
  const html = template
    .replaceAll('{{TITLE}}', esc(truth.title || 'バッチ処理状況一覧'))
    .replaceAll('{{META}}', esc(truth.meta || ''))
    .replaceAll('{{ROWS}}', buildRowsHtml(rows));

  if (!existsSync(args.out)) mkdirSync(args.out, { recursive: true });

  const browser = await chromium.launch({ executablePath: resolveChromium() });
  try {
    const page = await browser.newPage({ deviceScaleFactor: args.scale });
    await page.setContent(html, { waitUntil: 'networkidle' });
    const pngPath = join(args.out, `${stem}.png`);
    if (args.full) {
      await page.screenshot({ path: pngPath, fullPage: true });
    } else {
      const el = await page.$('table.batch');
      await (el || page).screenshot({ path: pngPath });
    }

    const manifest = {
      png: `${stem}.png`,
      source: basename(args.truth),
      title: truth.title || 'バッチ処理状況一覧',
      scale: args.scale,
      generatedAt: new Date().toISOString(),
      rows: rows.map((r, i) => ({
        rowIndex: i + 1,
        correlId: r.correlId,
        side: r.side || null,
        start: r.start,
        end: r.end,
        status: STATUS_TEXT[(r.status || 'normal').toLowerCase()] || r.status,
        datestamp: r.datestamp,
        count: r.count,
      })),
    };
    writeFileSync(join(args.out, `${stem}.manifest.json`), JSON.stringify(manifest, null, 2) + '\n', 'utf8');

    console.log(`[mock-page] rendered ${rows.length} row(s) -> ${pngPath}`);
    console.log(`[mock-page] manifest       -> ${join(args.out, `${stem}.manifest.json`)}`);
  } finally {
    await browser.close();
  }
}

main().catch(e => { console.error('[mock-page] FAILED:', e.message); process.exit(1); });
