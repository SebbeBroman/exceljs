/* oxlint-disable no-console */
/**
 * Browser experiment: excel-ts (esbuild browser bundle) vs exceljs (official browser dist).
 *
 * Measures:
 *   1) Ship size (minify + gzip) for write-only and write+load graphs
 *   2) Chrome headless write / read / round-trip timings (median of N runs)
 *
 * Why this exists: Node benches favor exceljs's mature Node/stream path.
 * Browser is where excel-ts (ESM, fflate, fewer polyfills) should be compared.
 *
 * Usage:
 *   pnpm build && pnpm bench:browser
 *   pnpm bench:browser -- --rows 2000 --runs 7
 *
 * Requires Google Chrome (or Chromium) on the machine. Override with CHROME_PATH.
 */
import * as esbuild from 'esbuild';
import {writeFileSync, mkdirSync, readFileSync, copyFileSync, existsSync} from 'node:fs';
import {join, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {gzipSync} from 'node:zlib';
import {createServer} from 'node:http';
import {extname} from 'node:path';
import puppeteer from 'puppeteer-core';

const __dirname = dirname(fileURLToPath(import.meta.url));
const root = join(__dirname, '../..');
const outDir = join(root, 'build', 'browser-bench');
mkdirSync(outDir, {recursive: true});

// --- CLI ---
const argv = process.argv.slice(2);
function flag(name, fallback) {
  const i = argv.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const v = argv[i + 1];
  if (v == null || v.startsWith('--')) return true;
  return v;
}
const ROWS = Number(flag('rows', 2000));
const COLS = Number(flag('cols', 8));
const RUNS = Number(flag('runs', 7));
const WARMUP = Number(flag('warmup', 2));

const fsAlias = join(root, 'lib/shims/fs-browser.ts');
const nodeModuleAlias = join(root, 'lib/shims/node-module-browser.ts');

const sharedEsbuild = {
  bundle: true,
  format: 'esm',
  platform: 'browser',
  mainFields: ['browser', 'module', 'main'],
  conditions: ['browser', 'import', 'default'],
  resolveExtensions: ['.ts', '.js', '.mjs', '.json'],
  define: {global: 'globalThis'},
  alias: {
    fs: fsAlias,
    'node:module': nodeModuleAlias,
    module: nodeModuleAlias,
  },
  minify: true,
  logLevel: 'warning',
};

function sizeReport(label, filePath) {
  const raw = readFileSync(filePath);
  const gz = gzipSync(raw);
  return {
    label,
    path: filePath,
    raw: raw.byteLength,
    gzip: gz.byteLength,
    rawKB: +(raw.byteLength / 1024).toFixed(1),
    gzipKB: +(gz.byteLength / 1024).toFixed(1),
  };
}

// --- Build excel-ts browser modules ---
const excelTsWriteEntry = join(outDir, 'entry-excel-ts-write.js');
writeFileSync(
  excelTsWriteEntry,
  `
import { workbook, writeBuffer } from '../../excel.ts';

export async function writeDense(rows, cols) {
  const grid = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    grid[r] = row;
  }
  const buf = await writeBuffer(
    workbook().sheet('data').rows(grid),
    { useSharedStrings: true, useStyles: false },
  );
  return buf.byteLength || buf.length;
}

export async function writeAndReturn(rows, cols) {
  const grid = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    grid[r] = row;
  }
  return writeBuffer(
    workbook().sheet('data').rows(grid),
    { useSharedStrings: true, useStyles: false },
  );
}
`,
);

const excelTsRoundTripEntry = join(outDir, 'entry-excel-ts-roundtrip.js');
writeFileSync(
  excelTsRoundTripEntry,
  `
import { workbook, writeBuffer, load } from '../../excel.ts';

export async function writeDense(rows, cols) {
  const grid = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    grid[r] = row;
  }
  const buf = await writeBuffer(
    workbook().sheet('data').rows(grid),
    { useSharedStrings: true, useStyles: false },
  );
  return buf.byteLength || buf.length;
}

export async function writeAndReturn(rows, cols) {
  const grid = new Array(rows);
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    grid[r] = row;
  }
  return writeBuffer(
    workbook().sheet('data').rows(grid),
    { useSharedStrings: true, useStyles: false },
  );
}

export async function loadBuf(bytes) {
  const u8 = bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes);
  const wb = await load(u8);
  return wb.sheets?.length ?? 0;
}

export async function roundTrip(rows, cols) {
  const buf = await writeAndReturn(rows, cols);
  const n = await loadBuf(buf);
  if (!n) throw new Error('no sheets');
  return buf.byteLength || buf.length;
}
`,
);

const excelTsWriteOut = join(outDir, 'excel-ts-write.min.mjs');
const excelTsRoundOut = join(outDir, 'excel-ts-roundtrip.min.mjs');

await esbuild.build({
  ...sharedEsbuild,
  entryPoints: [excelTsWriteEntry],
  outfile: excelTsWriteOut,
});
await esbuild.build({
  ...sharedEsbuild,
  entryPoints: [excelTsRoundTripEntry],
  outfile: excelTsRoundOut,
});

// --- exceljs official browser builds (not re-bundled) ---
const exceljsMin = join(root, 'node_modules/exceljs/dist/exceljs.min.js');
const exceljsBareMin = join(root, 'node_modules/exceljs/dist/exceljs.bare.min.js');
if (!existsSync(exceljsMin)) {
  throw new Error('exceljs browser dist missing — is exceljs a devDependency?');
}
copyFileSync(exceljsMin, join(outDir, 'exceljs.min.js'));
if (existsSync(exceljsBareMin)) {
  copyFileSync(exceljsBareMin, join(outDir, 'exceljs.bare.min.js'));
}

// Small ESM wrapper around global ExcelJS (loaded via classic script in the page)
const exceljsApi = join(outDir, 'exceljs-api.mjs');
writeFileSync(
  exceljsApi,
  `
/* global ExcelJS */
export async function writeDense(rows, cols) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('data');
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    ws.addRow(row);
  }
  const buf = await wb.xlsx.writeBuffer({ useSharedStrings: true, useStyles: false });
  return buf.byteLength || buf.length;
}

export async function writeAndReturn(rows, cols) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('data');
  for (let r = 0; r < rows; r++) {
    const row = new Array(cols);
    for (let c = 0; c < cols; c++) row[c] = c === 0 ? 'r' + r : r * cols + c;
    ws.addRow(row);
  }
  return wb.xlsx.writeBuffer({ useSharedStrings: true, useStyles: false });
}

export async function loadBuf(bytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb.worksheets?.length ?? 0;
}

export async function roundTrip(rows, cols) {
  const buf = await writeAndReturn(rows, cols);
  const n = await loadBuf(buf);
  if (!n) throw new Error('no sheets');
  return buf.byteLength || buf.length;
}
`,
);

const sizes = [
  sizeReport('excel-ts write-only (esbuild browser)', excelTsWriteOut),
  sizeReport('excel-ts write+load (esbuild browser)', excelTsRoundOut),
  sizeReport('exceljs dist/exceljs.min.js (official browser)', join(outDir, 'exceljs.min.js')),
];
if (existsSync(join(outDir, 'exceljs.bare.min.js'))) {
  sizes.push(sizeReport('exceljs dist/exceljs.bare.min.js', join(outDir, 'exceljs.bare.min.js')));
}

console.log(JSON.stringify({phase: 'sizes', rows: ROWS, cols: COLS, sizes}, null, 2));

// --- Static server for Chrome (ESM + classic script) ---
const mime = {
  '.html': 'text/html; charset=utf-8',
  '.js': 'text/javascript; charset=utf-8',
  '.mjs': 'text/javascript; charset=utf-8',
};

function startServer() {
  return new Promise(resolve => {
    const server = createServer((req, res) => {
      const urlPath = (req.url || '/').split('?')[0];
      const file = urlPath === '/' ? '/harness.html' : urlPath;
      const abs = join(outDir, file.replace(/^\//, ''));
      if (!abs.startsWith(outDir) || !existsSync(abs)) {
        res.writeHead(404);
        res.end('not found');
        return;
      }
      const body = readFileSync(abs);
      res.writeHead(200, {
        'Content-Type': mime[extname(abs)] || 'application/octet-stream',
        'Cache-Control': 'no-store',
      });
      res.end(body);
    });
    server.listen(0, '127.0.0.1', () => {
      const {port} = server.address();
      resolve({server, port});
    });
  });
}

const harnessHtml = `<!doctype html>
<html>
<head>
  <meta charset="utf-8" />
  <title>excel browser bench</title>
</head>
<body>
  <pre id="out">running…</pre>
  <!-- Official exceljs UMD browser build -->
  <script src="./exceljs.min.js"></script>
  <script type="module">
    import * as excelTsWrite from './excel-ts-write.min.mjs';
    import * as excelTsRt from './excel-ts-roundtrip.min.mjs';
    import * as exceljsApi from './exceljs-api.mjs';

    const params = new URLSearchParams(location.search);
    const ROWS = Number(params.get('rows') || ${ROWS});
    const COLS = Number(params.get('cols') || ${COLS});
    const RUNS = Number(params.get('runs') || ${RUNS});
    const WARMUP = Number(params.get('warmup') || ${WARMUP});

    function median(xs) {
      const a = [...xs].sort((x, y) => x - y);
      return a[Math.floor(a.length / 2)];
    }

    async function time(fn, runs, warmup) {
      for (let i = 0; i < warmup; i++) await fn();
      const times = [];
      for (let i = 0; i < runs; i++) {
        const t0 = performance.now();
        await fn();
        times.push(performance.now() - t0);
      }
      return {
        medianMs: Math.round(median(times) * 100) / 100,
        minMs: Math.round(Math.min(...times) * 100) / 100,
        maxMs: Math.round(Math.max(...times) * 100) / 100,
        timesMs: times.map(t => Math.round(t * 100) / 100),
      };
    }

    // Cache one buffer per library for pure read benches
    const tsBuf = await excelTsRt.writeAndReturn(ROWS, COLS);
    const xjBuf = await exceljsApi.writeAndReturn(ROWS, COLS);

    const results = {
      env: {
        userAgent: navigator.userAgent,
        rows: ROWS,
        cols: COLS,
        runs: RUNS,
        warmup: WARMUP,
        fixtureBytes: {
          excelTs: tsBuf.byteLength || tsBuf.length,
          exceljs: xjBuf.byteLength || xjBuf.length,
        },
      },
      benches: {},
    };

    results.benches['write: excel-ts builder'] = await time(
      () => excelTsWrite.writeDense(ROWS, COLS),
      RUNS,
      WARMUP,
    );
    results.benches['write: exceljs@4 browser'] = await time(
      () => exceljsApi.writeDense(ROWS, COLS),
      RUNS,
      WARMUP,
    );

    results.benches['read: excel-ts load'] = await time(
      () => excelTsRt.loadBuf(tsBuf),
      RUNS,
      WARMUP,
    );
    results.benches['read: exceljs load'] = await time(
      () => exceljsApi.loadBuf(xjBuf),
      RUNS,
      WARMUP,
    );

    // Cross-read (format interop + cost)
    results.benches['read: excel-ts load(exceljs bytes)'] = await time(
      () => excelTsRt.loadBuf(xjBuf instanceof Uint8Array ? xjBuf : new Uint8Array(xjBuf)),
      RUNS,
      WARMUP,
    );
    results.benches['read: exceljs load(excel-ts bytes)'] = await time(
      () => exceljsApi.loadBuf(tsBuf),
      RUNS,
      WARMUP,
    );

    results.benches['roundtrip: excel-ts'] = await time(
      () => excelTsRt.roundTrip(ROWS, COLS),
      RUNS,
      WARMUP,
    );
    results.benches['roundtrip: exceljs'] = await time(
      () => exceljsApi.roundTrip(ROWS, COLS),
      RUNS,
      WARMUP,
    );

    document.getElementById('out').textContent = JSON.stringify(results, null, 2);
    window.__BENCH_RESULTS__ = results;
  </script>
</body>
</html>
`;
writeFileSync(join(outDir, 'harness.html'), harnessHtml);

function findChrome() {
  if (process.env.CHROME_PATH && existsSync(process.env.CHROME_PATH)) {
    return process.env.CHROME_PATH;
  }
  const candidates = [
    '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    '/Applications/Chromium.app/Contents/MacOS/Chromium',
    '/usr/bin/google-chrome',
    '/usr/bin/chromium',
    '/usr/bin/chromium-browser',
  ];
  for (const c of candidates) {
    if (existsSync(c)) return c;
  }
  return null;
}

const chromePath = findChrome();
if (!chromePath) {
  console.error('No Chrome/Chromium found. Set CHROME_PATH or install Chrome.');
  console.error('Size comparison above still valid.');
  writeFileSync(join(outDir, 'sizes.json'), JSON.stringify({sizes}, null, 2));
  process.exit(2);
}

const {server, port} = await startServer();
const url = `http://127.0.0.1:${port}/harness.html?rows=${ROWS}&cols=${COLS}&runs=${RUNS}&warmup=${WARMUP}`;

console.log(JSON.stringify({phase: 'chrome', chromePath, url}));

let browser;
try {
  browser = await puppeteer.launch({
    executablePath: chromePath,
    headless: true,
    args: ['--no-sandbox', '--disable-gpu'],
  });
  const page = await browser.newPage();
  page.on('pageerror', err => console.error('pageerror', err));
  page.on('console', msg => {
    if (msg.type() === 'error') console.error('console', msg.text());
  });

  await page.goto(url, {waitUntil: 'networkidle0', timeout: 120_000});
  await page.waitForFunction(() => window.__BENCH_RESULTS__, {timeout: 300_000});
  const perf = await page.evaluate(() => window.__BENCH_RESULTS__);

  const report = {
    suite: 'browser excel-ts vs exceljs',
    sizes,
    perf,
  };

  writeFileSync(join(outDir, 'results.json'), JSON.stringify(report, null, 2));

  // Pretty table
  console.log('\n=== Bundle size (browser) ===');
  for (const s of sizes) {
    console.log(
      `  ${s.label.padEnd(48)} raw ${String(s.rawKB).padStart(7)} KB   gzip ${String(s.gzipKB).padStart(6)} KB`,
    );
  }

  console.log(`\n=== Chrome timings (${ROWS}×${COLS}, median of ${RUNS}, warmup ${WARMUP}) ===`);
  console.log(`  UA: ${perf.env.userAgent}`);
  console.log(
    `  fixture bytes: excel-ts=${perf.env.fixtureBytes.excelTs}  exceljs=${perf.env.fixtureBytes.exceljs}`,
  );
  for (const [name, stats] of Object.entries(perf.benches)) {
    console.log(
      `  ${name.padEnd(40)} median ${String(stats.medianMs).padStart(8)} ms   (min ${stats.minMs} … max ${stats.maxMs})`,
    );
  }

  // Speed ratios for headline pairs
  const b = perf.benches;
  const ratio = (a, bms) => (bms / a).toFixed(2);
  if (b['write: excel-ts builder'] && b['write: exceljs@4 browser']) {
    const ts = b['write: excel-ts builder'].medianMs;
    const xj = b['write: exceljs@4 browser'].medianMs;
    console.log(
      `\n  write: excel-ts is ${ratio(ts, xj)}× relative to exceljs  (excel-ts ${ts}ms, exceljs ${xj}ms)`,
    );
  }
  if (b['read: excel-ts load'] && b['read: exceljs load']) {
    const ts = b['read: excel-ts load'].medianMs;
    const xj = b['read: exceljs load'].medianMs;
    console.log(
      `  read:  excel-ts is ${ratio(ts, xj)}× relative to exceljs  (excel-ts ${ts}ms, exceljs ${xj}ms)`,
    );
  }
  if (b['roundtrip: excel-ts'] && b['roundtrip: exceljs']) {
    const ts = b['roundtrip: excel-ts'].medianMs;
    const xj = b['roundtrip: exceljs'].medianMs;
    console.log(
      `  round: excel-ts is ${ratio(ts, xj)}× relative to exceljs  (excel-ts ${ts}ms, exceljs ${xj}ms)`,
    );
  }

  console.log(`\nWrote ${join(outDir, 'results.json')}`);
} finally {
  if (browser) await browser.close();
  server.close();
}
