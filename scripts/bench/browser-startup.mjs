/* oxlint-disable no-console */
// Measure static startup and actual chunks fetched by the first plain XLSX export.
import {build} from 'esbuild';
import {mkdirSync, readFileSync, writeFileSync} from 'node:fs';
import {resolve, dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createServer} from 'node:http';
import {gzipSync} from 'node:zlib';
import puppeteer from 'puppeteer-core';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/browser-startup');
mkdirSync(out, {recursive: true});
const report = {};
for (const variant of ['static-reader', 'lazy-reader']) {
  const dest = join(out, variant);
  const plugins =
    variant === 'static-reader'
      ? [
          {
            name: 'baseline-reader',
            setup(b) {
              b.onLoad({filter: /\/dist\/lib\/xlsx\/xlsx\.js$/}, ({path}) => {
                let text = readFileSync(path, 'utf8');
                text =
                  "import {canUseFastSheetData,isFastSheetDataEnabled,parseFastSheetData,splitSheetData} from './xform/sheet/fast-sheet-data.js';\n" +
                  text;
                text = text
                  .replace('const { parseFastSheetData } = await loadFastSheetData();', '')
                  .replace(
                    'const { isFastSheetDataEnabled, splitSheetData, canUseFastSheetData } = await loadFastSheetData();',
                    '',
                  );
                if (text.includes('await loadFastSheetData();'))
                  throw new Error('Baseline replacement failed');
                return {contents: text, loader: 'js', resolveDir: dirname(path)};
              });
            },
          },
        ]
      : [];
  const result = await build({
    bundle: true,
    minify: true,
    splitting: true,
    metafile: true,
    platform: 'browser',
    format: 'esm',
    target: 'es2022',
    stdin: {
      contents: `import {workbook,writeBuffer} from ${JSON.stringify(join(root, 'dist/excel.js'))};export {workbook,writeBuffer};`,
      resolveDir: root,
    },
    outdir: dest,
    define: {global: 'globalThis'},
    alias: {
      fs: join(root, 'lib/shims/fs-browser.ts'),
      module: join(root, 'lib/shims/node-module-browser.ts'),
      'node:module': join(root, 'lib/shims/node-module-browser.ts'),
    },
    plugins,
  });
  const outputs = result.metafile.outputs;
  const entry = Object.keys(outputs).find(p => outputs[p].entryPoint === '<stdin>');
  const initial = new Set();
  function visit(p) {
    if (initial.has(p)) return;
    initial.add(p);
    for (const d of outputs[p].imports)
      if (d.kind !== 'dynamic-import' && outputs[d.path]) visit(d.path);
  }
  visit(entry);
  const files = Object.keys(outputs).map(p => ({
    path: p,
    url: '/' + resolve(root, p).slice(out.length + 1),
    raw: readFileSync(resolve(root, p)).length,
    gzip: gzipSync(readFileSync(resolve(root, p))).length,
    initial: initial.has(p),
  }));
  report[variant] = {
    entry: '/' + resolve(root, entry).slice(out.length + 1),
    files,
    initialGzip: files.filter(f => f.initial).reduce((s, f) => s + f.gzip, 0),
    totalGzip: files.reduce((s, f) => s + f.gzip, 0),
  };
}
const server = createServer((req, res) => {
  if (req.url === '/') {
    res.end('<!doctype html><title>Startup measurement</title>');
    return;
  }
  try {
    const path = resolve(out, '.' + req.url);
    if (!path.startsWith(out + '/')) throw new Error('Invalid path');
    res.setHeader('Content-Type', 'text/javascript');
    res.end(readFileSync(path));
  } catch {
    res.writeHead(404);
    res.end();
  }
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
let browser;
try {
  browser = await puppeteer.launch({
    executablePath:
      process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true,
  });
  for (const [variant, r] of Object.entries(report)) {
    const page = await browser.newPage();
    const fetched = new Set();
    page.on('request', q => fetched.add(new URL(q.url()).pathname));
    await page.goto(`http://127.0.0.1:${server.address().port}/`);
    r.firstExport = await page.evaluate(async entry => {
      const m = await import(entry);
      const grid = Array.from({length: 2000}, (_, r) => [String(r), r]);
      const bytes = await m.writeBuffer(m.workbook().sheet('data').rows(grid), {useStyles: false});
      return {bytes: bytes.length};
    }, r.entry);
    r.firstExportGzip = r.files.filter(f => fetched.has(f.url)).reduce((s, f) => s + f.gzip, 0);
    r.firstExportFiles = r.files.filter(f => fetched.has(f.url)).map(f => f.url);
    await page.close();
  }
} finally {
  await browser?.close();
  await new Promise(r => server.close(r));
}
writeFileSync(join(out, 'results.json'), JSON.stringify(report, null, 2));
console.table(
  Object.entries(report).map(([variant, r]) => ({
    variant,
    initialGzipKiB: r.initialGzip / 1024,
    firstExportGzipKiB: r.firstExportGzip / 1024,
    totalGzipKiB: r.totalGzip / 1024,
    exportBytes: r.firstExport.bytes,
  })),
);
