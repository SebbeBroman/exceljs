/* oxlint-disable no-console */
/** Compare a saved baseline dist with the current build in one browser session.
 * Before editing: pnpm build; mkdir -p build/doc-bridge-before; cp -R dist build/doc-bridge-before/dist
 * After editing: pnpm build; node scripts/bench/doc-bridge.mjs
 */
import {build} from 'esbuild';
import {mkdirSync, readFileSync, writeFileSync, existsSync} from 'node:fs';
import {resolve, join, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createServer} from 'node:http';
import {gzipSync} from 'node:zlib';
import puppeteer from 'puppeteer-core';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/doc-bridge-comparison');
mkdirSync(out, {recursive: true});
const sizes = {};
const baseline = process.argv[2]
  ? resolve(process.argv[2])
  : join(root, 'build/doc-bridge-before/dist');
if (!existsSync(join(baseline, 'excel.js')))
  throw new Error('Save the baseline dist before editing');
for (const [name, dist] of Object.entries({before: baseline, after: join(root, 'dist')})) {
  sizes[name] = {};
  for (const mode of ['write', 'read-write']) {
    const exports = mode === 'write' ? 'workbook,writeBuffer' : 'workbook,writeBuffer,load';
    const outfile = join(out, `${name}-${mode}.mjs`);
    const result = await build({
      stdin: {
        contents: `export {${exports}} from ${JSON.stringify(join(dist, 'excel.js'))}`,
        resolveDir: root,
      },
      outfile,
      bundle: true,
      minify: true,
      platform: 'browser',
      format: 'esm',
      target: 'es2022',
      metafile: true,
      define: {global: 'globalThis'},
      alias: {
        fs: join(root, 'lib/shims/fs-browser.ts'),
        module: join(root, 'lib/shims/node-module-browser.ts'),
        'node:module': join(root, 'lib/shims/node-module-browser.ts'),
      },
    });
    const bytes = readFileSync(outfile);
    sizes[name][mode] = {rawBytes: bytes.length, gzipBytes: gzipSync(bytes).length};
    writeFileSync(`${outfile}.meta.json`, JSON.stringify(result.metafile, null, 2));
  }
}
writeFileSync(
  join(out, 'index.html'),
  '<!doctype html><title>DocWorkbook bridge comparison</title>',
);
const server = createServer((req, res) => {
  const file = (req.url || '/').slice(1) || 'index.html';
  if (!/^[\w.-]+$/.test(file) || !existsSync(join(out, file))) {
    res.writeHead(404);
    res.end();
    return;
  }
  res.setHeader('Content-Type', file.endsWith('.html') ? 'text/html' : 'text/javascript');
  res.end(readFileSync(join(out, file)));
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
let browser;
try {
  browser = await puppeteer.launch({
    executablePath:
      process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true,
    args: ['--js-flags=--expose-gc'],
  });
  const page = await browser.newPage();
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const results = await page.evaluate(async () => {
    const apis = {
      before: await import('./before-read-write.mjs'),
      after: await import('./after-read-write.mjs'),
    };
    const results = [];
    const meta = {
      creator: 'comparison',
      created: new Date('2020-01-01'),
      modified: new Date('2020-01-01'),
    };
    const options = {useSharedStrings: true, useStyles: false};
    for (const rowCount of [2000, 20000]) {
      const grid = Array.from({length: rowCount}, (_, r) =>
        Array.from({length: 8}, (_, c) => (c ? r * 8 + c : `r${r}`)),
      );
      const make = api => api.workbook(meta).sheet('data').rows(grid);
      const buffers = {};
      for (const [name, api] of Object.entries(apis))
        buffers[name] = await api.writeBuffer(make(api), options);
      for (const [writer, bytes] of Object.entries(buffers))
        for (const [reader, api] of Object.entries(apis)) {
          const plain = await api.load(bytes);
          const actual = plain.sheets[0].rows.map(row =>
            Array.from({length: 8}, (_, c) => row.cells[c + 1]?.value),
          );
          if (JSON.stringify(actual) !== JSON.stringify(grid))
            throw new Error(`${writer} → ${reader}: mismatched values`);
        }
      const snapshot = await apis.before.load(buffers.before);
      const jobs = {
        'builder write': api => api.writeBuffer(make(api), options),
        'snapshot write': api => api.writeBuffer(snapshot, options),
        'full read': api => api.load(buffers.before),
        roundtrip: async api => api.load(await api.writeBuffer(make(api), options)),
      };
      const timings = {};
      for (const [job, run] of Object.entries(jobs)) {
        const samples = {before: [], after: []};
        for (let iteration = 0; iteration < 14; iteration++) {
          const order = iteration % 2 ? ['after', 'before'] : ['before', 'after'];
          for (const name of order) {
            globalThis.gc?.();
            const start = performance.now();
            await run(apis[name]);
            const elapsed = performance.now() - start;
            if (iteration >= 3) samples[name].push(elapsed);
          }
        }
        timings[job] = Object.fromEntries(
          Object.entries(samples).map(([name, values]) => [
            name,
            {medianMs: [...values].sort((a, b) => a - b)[5], samplesMs: values},
          ]),
        );
      }
      results.push({
        rowCount,
        cols: 8,
        warmups: 3,
        runs: 11,
        outputBytes: Object.fromEntries(Object.entries(buffers).map(([n, b]) => [n, b.length])),
        timings,
      });
    }
    return {userAgent: navigator.userAgent, results};
  });
  const report = {
    date: new Date().toISOString(),
    node: process.version,
    baseline,
    sizes,
    ...results,
  };
  writeFileSync(join(out, 'results.json'), JSON.stringify(report, null, 2));
  for (const result of report.results) {
    console.log(`${result.rowCount} × ${result.cols} cells`);
    console.table(
      Object.entries(result.timings).map(([job, t]) => ({
        job,
        beforeMs: t.before.medianMs.toFixed(1),
        afterMs: t.after.medianMs.toFixed(1),
        improvementPercent: (100 * (1 - t.after.medianMs / t.before.medianMs)).toFixed(1),
      })),
    );
  }
  console.log(JSON.stringify(sizes, null, 2));
} finally {
  await browser?.close();
  await new Promise(r => server.close(r));
}
