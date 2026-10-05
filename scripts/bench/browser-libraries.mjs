/* oxlint-disable no-console */
/** Browser size and runtime comparison with local fast-csv wired in.
 * Run: node scripts/bench/browser-libraries.mjs
 * SheetJS sources are downloaded to ignored build/browser-libraries only.
 */
import {build, version as esbuildVersion} from 'esbuild';
import {existsSync, mkdirSync, readFileSync, writeFileSync, copyFileSync} from 'node:fs';
import {resolve, dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createServer} from 'node:http';
import {gzipSync, brotliCompressSync} from 'node:zlib';
import puppeteer from 'puppeteer-core';
const rowFlag = process.argv.indexOf('--rows');
const ROW_COUNT = rowFlag < 0 ? 2000 : Number(process.argv[rowFlag + 1]);
if (!Number.isSafeInteger(ROW_COUNT) || ROW_COUNT < 1) throw new Error('Invalid row count');
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/browser-libraries');
mkdirSync(out, {recursive: true});
for (const [file, remote] of [
  ['sheetjs.mjs', 'xlsx.mjs'],
  ['sheetjs.full.min.js', 'dist/xlsx.full.min.js'],
  ['sheetjs.mini.min.js', 'dist/xlsx.mini.min.js'],
]) {
  if (!existsSync(join(out, file))) {
    const response = await fetch(`https://cdn.sheetjs.com/xlsx-0.20.3/package/${remote}`);
    if (!response.ok) throw new Error(`SheetJS download: ${response.status}`);
    writeFileSync(join(out, file), Buffer.from(await response.arrayBuffer()));
  }
}
const shared = {
  bundle: true,
  minify: true,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  metafile: true,
  logLevel: 'warning',
  define: {global: 'globalThis'},
  alias: {
    fs: join(root, 'lib/shims/fs-browser.ts'),
    module: join(root, 'lib/shims/node-module-browser.ts'),
    'node:module': join(root, 'lib/shims/node-module-browser.ts'),
  },
};
const local = JSON.stringify(join(root, 'dist/excel.js'));
const sheetjs = JSON.stringify(join(out, 'sheetjs.mjs'));
const entries = {
  'local-write': `import {workbook,writeBuffer} from ${local}; export {workbook,writeBuffer};`,
  'local-read-write': `import {workbook,writeBuffer,load} from ${local}; export {workbook,writeBuffer,load};`,
  'local-view': `import {viewWorkbook} from ${local}; export {viewWorkbook};`,
  'local-csv': `import {csv} from ${local}; export {csv};`,
  'sheetjs-write': `import {write,utils} from ${sheetjs}; export {write,utils};`,
  'sheetjs-read-write': `import {read,write,utils,version} from ${sheetjs}; export {read,write,utils,version};`,
  'sheetjs-read': `import {read,utils} from ${sheetjs}; export {read,utils};`,
};
const sizes = [];
function size(name, file, extra = {}) {
  const bytes = readFileSync(file);
  sizes.push({
    name,
    raw: bytes.length,
    gzip: gzipSync(bytes).length,
    brotli: brotliCompressSync(bytes).length,
    ...extra,
  });
}
for (const [name, source] of Object.entries(entries)) {
  const file = join(out, `${name}.mjs`);
  const result = await build({
    ...shared,
    stdin: {contents: source, resolveDir: root},
    outfile: file,
  });
  writeFileSync(`${file}.meta.json`, JSON.stringify(result.metafile, null, 2));
  size(name, file, {modules: Object.keys(result.metafile.inputs).length});
}
for (const name of ['exceljs.min.js', 'exceljs.bare.min.js']) {
  copyFileSync(join(root, 'node_modules/exceljs/dist', name), join(out, name));
  size(name, join(out, name));
}
size('sheetjs-full-standalone', join(out, 'sheetjs.full.min.js'));
size('sheetjs-mini-standalone', join(out, 'sheetjs.mini.min.js'));
console.table(
  sizes.map(x => ({
    name: x.name,
    rawKiB: (x.raw / 1024).toFixed(1),
    gzipKiB: (x.gzip / 1024).toFixed(1),
  })),
);
writeFileSync(
  join(out, 'harness.html'),
  '<!doctype html><title>Browser library comparison</title><script src="./exceljs.min.js"></script>',
);
const server = createServer((req, res) => {
  const name = (req.url || '/').split('?')[0].slice(1) || 'harness.html';
  if (!/^[\w.-]+$/.test(name) || !existsSync(join(out, name))) {
    res.writeHead(404);
    res.end();
    return;
  }
  res.setHeader('Content-Type', name.endsWith('.html') ? 'text/html' : 'text/javascript');
  res.end(readFileSync(join(out, name)));
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
let browser;
try {
  browser = await puppeteer.launch({
    executablePath:
      process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true,
  });
  const page = await browser.newPage();
  page.on('pageerror', e => console.error(e.message));
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const perf = await page.evaluate(async rowCount => {
    const local = {
      ...(await import('./local-read-write.mjs')),
      ...(await import('./local-view.mjs')),
    };
    const sheetjs = await import('./sheetjs-read-write.mjs');
    const grid = Array.from({length: rowCount}, (_, r) =>
      Array.from({length: 8}, (_, c) => (c === 0 ? `r${r}` : r * 8 + c)),
    );
    const writeOptions = {useSharedStrings: true, useStyles: false};
    const APIs = {
      local: {
        write: () => local.writeBuffer(local.workbook().sheet('data').rows(grid), writeOptions),
        read: b => local.load(b),
        values: w =>
          w.sheets[0].rows.map(r => Array.from({length: 8}, (_, c) => r.cells[c + 1]?.value ?? '')),
      },
      exceljs: {
        write: async () => {
          const wb = new window.ExcelJS.Workbook();
          wb.addWorksheet('data').addRows(grid);
          return wb.xlsx.writeBuffer(writeOptions);
        },
        read: async b => {
          const wb = new window.ExcelJS.Workbook();
          await wb.xlsx.load(b);
          return wb;
        },
        values: w =>
          w.worksheets[0]
            .getSheetValues()
            .slice(1)
            .map(r => r.slice(1)),
      },
      sheetjs: {
        write: () => {
          const wb = sheetjs.utils.book_new();
          sheetjs.utils.book_append_sheet(wb, sheetjs.utils.aoa_to_sheet(grid), 'data');
          return new Uint8Array(
            sheetjs.write(wb, {type: 'array', bookType: 'xlsx', compression: true, bookSST: true}),
          );
        },
        read: b => sheetjs.read(b, {type: 'array', dense: true}),
        values: w =>
          sheetjs.utils.sheet_to_json(w.Sheets[w.SheetNames[0]], {
            header: 1,
            raw: true,
            defval: '',
          }),
      },
    };
    const buffers = {};
    for (const [name, api] of Object.entries(APIs))
      buffers[name] = new Uint8Array(await api.write());
    const expected = JSON.stringify(grid);
    for (const [writer, bytes] of Object.entries(buffers))
      for (const [reader, api] of Object.entries(APIs)) {
        const rows = api.values(await api.read(bytes));
        if (JSON.stringify(rows) !== expected)
          throw new Error(`Cross-read mismatch: ${writer} → ${reader}`);
      }
    const view = await local.viewWorkbook(buffers.local, {format: 'xlsx'});
    if (JSON.stringify(view.sheet(0).rows({values: 'cell'})) !== expected)
      throw new Error('View values mismatch');
    const jobs = {
      write: Object.fromEntries(Object.entries(APIs).map(([n, a]) => [n, () => a.write()])),
      'read shared XLSX': Object.fromEntries(
        Object.entries(APIs).map(([n, a]) => [n, () => a.read(buffers.local)]),
      ),
      'read own XLSX': Object.fromEntries(
        Object.entries(APIs).map(([n, a]) => [n, () => a.read(buffers[n])]),
      ),
      roundtrip: Object.fromEntries(
        Object.entries(APIs).map(([n, a]) => [
          n,
          async () => a.read(new Uint8Array(await a.write())),
        ]),
      ),
      'read values shared XLSX': {
        'local-view': async () => {
          const w = await local.viewWorkbook(buffers.local, {format: 'xlsx'});
          return w.sheet(0).rows({values: 'cell'});
        },
        sheetjs: () => APIs.sheetjs.values(APIs.sheetjs.read(buffers.local)),
        exceljs: async () => APIs.exceljs.values(await APIs.exceljs.read(buffers.local)),
      },
    };
    const timings = {};
    for (const [job, libs] of Object.entries(jobs)) {
      const names = Object.keys(libs);
      const samples = Object.fromEntries(names.map(n => [n, []]));
      for (let i = 0; i < 9; i++)
        for (let k = 0; k < names.length; k++) {
          const n = names[(k + i) % names.length];
          const start = performance.now();
          await libs[n]();
          if (i >= 2) samples[n].push(performance.now() - start);
        }
      timings[job] = Object.fromEntries(
        Object.entries(samples).map(([n, x]) => [
          n,
          {medianMs: [...x].sort((a, b) => a - b)[3], samplesMs: x},
        ]),
      );
    }
    return {
      userAgent: navigator.userAgent,
      versions: {sheetjs: sheetjs.version, exceljs: '4.4.0', local: '0.1.0'},
      rows: rowCount,
      cols: 8,
      runs: 7,
      warmup: 2,
      outputBytes: Object.fromEntries(Object.entries(buffers).map(([n, b]) => [n, b.length])),
      validation: `Nine writer/reader combinations and local values view reproduce all ${rowCount * 8} cells`,
      timings,
    };
  }, ROW_COUNT);
  const report = {
    date: new Date().toISOString(),
    node: process.version,
    esbuild: esbuildVersion,
    sizes,
    perf,
  };
  const reportName = ROW_COUNT === 2000 ? 'results.json' : `results-${ROW_COUNT}.json`;
  writeFileSync(join(out, reportName), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(perf, null, 2));
} finally {
  if (browser) await browser.close();
  await new Promise(r => server.close(r));
}
