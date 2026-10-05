/** Local browser comparison. Run: node scripts/bench/local-fast-csv.mjs
 * Baseline is ExcelJS HEAD + registry fast-csv 5.0.7 + real browser polyfills.
 * Candidate is this checkout + linked fast-csv/browser, without CSV stubs.
 * --sizes-only updates sizes without Chrome; --microtask-only also checks an
 * optimized registry nextTick scheduler and saves performanceMicrotask.
 * Install baseline-only deps in build/local-fast-csv/polyfills with npm:
 * npm install --prefix build/local-fast-csv/polyfills --no-audit --no-fund --ignore-scripts fast-csv@5.0.7 stream-browserify util buffer process events string_decoder
 */
import {build, version as esbuildVersion} from 'esbuild';
import {mkdirSync, readFileSync, writeFileSync, existsSync} from 'node:fs';
import {join, dirname, resolve} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {execFileSync} from 'node:child_process';
import {gzipSync, brotliCompressSync} from 'node:zlib';
import {createServer} from 'node:http';
import puppeteer from 'puppeteer-core';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/local-fast-csv');
mkdirSync(out, {recursive: true});
const baseline = join(out, 'baseline');
if (!existsSync(join(baseline, 'excel.ts'))) {
  mkdirSync(baseline, {recursive: true});
  const archive = execFileSync('git', ['archive', 'HEAD', 'excel.ts', 'lib'], {cwd: root});
  execFileSync('tar', ['-x', '-C', baseline], {input: archive});
}
const requirePolyfill = createRequire(join(out, 'polyfills/package.json'));
const inject = join(out, 'inject.js');
writeFileSync(
  inject,
  `export {default as process} from ${JSON.stringify(requirePolyfill.resolve('process/browser.js'))};\nexport {Buffer} from ${JSON.stringify(requirePolyfill.resolve('buffer/'))};\n`,
);
writeFileSync(
  inject,
  readFileSync(inject, 'utf8') +
    `
const channel = new MessageChannel();
const tasks = new Map(); let next = 0;
channel.port1.onmessage = e => { const fn = tasks.get(e.data); tasks.delete(e.data); if (fn) fn(); };
export function setImmediate(fn, ...args) { const id = ++next; tasks.set(id, () => fn(...args)); channel.port2.postMessage(id); return id; }
export function clearImmediate(id) { tasks.delete(id); }
`,
);
const baseAliases = {
  fs: join(root, 'lib/shims/fs-browser.ts'),
  module: join(root, 'lib/shims/node-module-browser.ts'),
  'node:module': join(root, 'lib/shims/node-module-browser.ts'),
};
const polyfills = {
  stream: requirePolyfill.resolve('stream-browserify'),
  util: requirePolyfill.resolve('util/'),
  buffer: requirePolyfill.resolve('buffer/'),
  events: requirePolyfill.resolve('events/'),
  string_decoder: requirePolyfill.resolve('string_decoder/'),
  process: requirePolyfill.resolve('process/browser.js'),
};
const shared = {
  bundle: true,
  minify: true,
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  metafile: true,
  logLevel: 'silent',
  define: {global: 'globalThis'},
};
const cases = {
  'csv-parse': `import {parseCsv} from ENTRY; export const parse = parseCsv;`,
  'csv-stringify': `import {stringifyCsv} from ENTRY; export const stringify = stringifyCsv;`,
  'csv-both': `import {csv} from ENTRY; export {csv};`,
  'csv-view': `import {viewWorkbook} from ENTRY; export {viewWorkbook};`,
  'xlsx-write': `import {workbook, writeBuffer} from ENTRY; export {workbook, writeBuffer};`,
  'xlsx-read-write': `import {workbook, writeBuffer, load} from ENTRY; export {workbook, writeBuffer, load};`,
};
const reports = [];
const unpolyfilled = [];
for (const [variant, entry] of [
  ['registry', join(baseline, 'excel.ts')],
  ['local', join(root, 'excel.ts')],
]) {
  for (const [name, source] of Object.entries(cases)) {
    const options = {
      ...shared,
      stdin: {
        contents: source.replace('ENTRY', JSON.stringify(entry)),
        resolveDir: root,
        sourcefile: `${name}.js`,
      },
      alias: {
        ...baseAliases,
        ...(variant === 'registry' ? {'fast-csv': requirePolyfill.resolve('fast-csv')} : {}),
      },
    };
    if (variant === 'registry') {
      try {
        await build({...options, write: false});
        unpolyfilled.push({name, ok: true});
      } catch (e) {
        unpolyfilled.push({name, ok: false, errors: [...new Set(e.errors.map(x => x.text))]});
      }
      options.alias = {...options.alias, ...polyfills};
      options.inject = [inject];
    }
    const file = join(out, `${variant}-${name}.mjs`);
    const result = await build({...options, outfile: file});
    const bytes = readFileSync(file);
    writeFileSync(`${file}.meta.json`, JSON.stringify(result.metafile, null, 2));
    const inputs = Object.keys(result.metafile.inputs);
    const nodePolyfillInputs = inputs.filter(x =>
      /node_modules\/(stream-browserify|readable-stream|buffer|process|util|events|string_decoder)\//.test(
        x,
      ),
    );
    if (variant === 'local' && nodePolyfillInputs.length)
      throw new Error(`Unexpected polyfill: ${nodePolyfillInputs}`);
    reports.push({
      variant,
      name,
      raw: bytes.length,
      gzip: gzipSync(bytes).length,
      brotli: brotliCompressSync(bytes).length,
      modules: inputs.length,
      nodePolyfillModules: nodePolyfillInputs.length,
    });
    const split = await build({
      ...options,
      splitting: true,
      write: false,
      outdir: join(out, `split-${variant}-${name}`),
    });
    reports.at(-1).split = split.outputFiles.map(f => ({
      name: f.path.split('/').pop(),
      raw: f.contents.length,
      gzip: gzipSync(f.contents).length,
    }));
  }
}
console.table(
  reports.map(({variant, name, raw, gzip, brotli, nodePolyfillModules}) => ({
    variant,
    name,
    raw,
    gzip,
    brotli,
    nodePolyfillModules,
  })),
);
const microtaskOnly = process.argv.includes('--microtask-only');
if (microtaskOnly) {
  const optimizedInject = join(out, 'inject-microtask.js');
  const processPath = JSON.stringify(requirePolyfill.resolve('process/browser.js'));
  const originalInject = readFileSync(inject, 'utf8');
  writeFileSync(
    optimizedInject,
    originalInject.replace(
      `export {default as process} from ${processPath};`,
      `import process from ${processPath}; process.nextTick = (fn, ...args) => queueMicrotask(() => fn(...args)); export {process};`,
    ),
  );
  await build({
    ...shared,
    stdin: {
      contents: `import {csv} from ${JSON.stringify(join(baseline, 'excel.ts'))}; export {csv};`,
      resolveDir: root,
    },
    alias: {...baseAliases, ...polyfills, 'fast-csv': requirePolyfill.resolve('fast-csv')},
    inject: [optimizedInject],
    outfile: join(out, 'registry-microtask-csv.mjs'),
  });
}
if (process.argv.includes('--sizes-only')) {
  const path = join(out, 'results.json');
  const previous = existsSync(path) ? JSON.parse(readFileSync(path, 'utf8')) : {};
  writeFileSync(path, JSON.stringify({...previous, reports, unpolyfilled}, null, 2));
  console.log(`Updated sizes: ${path}`);
  process.exit(0);
}
writeFileSync(join(out, 'harness.html'), '<!doctype html><title>Local fast-csv benchmark</title>');
const server = createServer((req, res) => {
  const name = (req.url || '/').split('?')[0].replace(/^\//, '') || 'harness.html';
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
  page.on('console', msg => console.log('browser:', msg.text()));
  page.on('pageerror', err => console.error('browser error:', err.message));
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const performance = await page.evaluate(async microtask => {
    const assert = (ok, msg) => {
      if (!ok) throw new Error(msg);
    };
    const equal = (a, b) =>
      assert(
        JSON.stringify(a) === JSON.stringify(b),
        `Mismatch: ${JSON.stringify(a)} / ${JSON.stringify(b)}`,
      );
    assert(
      typeof globalThis.process === 'undefined' && typeof globalThis.Buffer === 'undefined',
      'Node globals present',
    );
    const modules = {};
    for (const variant of ['registry', 'local']) {
      modules[variant] = {
        csv: (
          await import(
            variant === 'registry' && microtask
              ? './registry-microtask-csv.mjs'
              : `./${variant}-csv-both.mjs`
          )
        ).csv,
        view: (await import(`./${variant}-csv-view.mjs`)).viewWorkbook,
        xlsx: await import(`./${variant}-xlsx-read-write.mjs`),
      };
    }
    const fixtures = [
      ['a,b\r\n"hello, world","a""b"\r\n"two\nlines",42', {}],
      ['name;value\nalpha;1\nbeta;2', {parserOptions: {delimiter: ';', headers: true}}],
      ['a,b\n1,2\n\n3,4', {parserOptions: {ignoreEmpty: true, trim: true, maxRows: 2}}],
      ['# comment\na,b\n1,2', {parserOptions: {comment: '#'}}],
      ['a,b\ntrue,#N/A\n2020-01-15,', {}],
    ];
    console.log('CSV fixtures');
    for (const [text, opts] of fixtures)
      equal(
        await modules.registry.csv.parse(text, opts),
        await modules.local.csv.parse(text, opts),
      );
    for (const variant of ['registry', 'local']) {
      let rejected = false;
      try {
        await modules[variant].csv.parse('a,"unterminated');
      } catch {
        rejected = true;
      }
      assert(rejected, 'Malformed quote did not reject');
    }
    const rows = Array.from({length: 2000}, (_, r) => [
      `row ${r}, quoted`,
      ...Array.from({length: 7}, (_, c) => String(r * 8 + c)),
    ]);
    const text = rows.map(row => '"' + row[0] + '",' + row.slice(1).join(',')).join('\n');
    const models = {};
    for (const variant of ['registry', 'local'])
      models[variant] = modules[variant].xlsx.workbook().sheet('Data').rows(rows).build();
    console.log('formatter parity');
    for (const formatterOptions of [
      {},
      {
        delimiter: ';',
        rowDelimiter: '\r\n',
        quoteColumns: true,
        writeBOM: true,
        includeEndRowDelimiter: true,
      },
      {headers: ['a', 'b'], alwaysWriteHeaders: true},
    ]) {
      equal(
        await modules.registry.csv.stringify(models.registry, {formatterOptions}),
        await modules.local.csv.stringify(models.local, {formatterOptions}),
      );
    }
    console.log('view parity');
    equal(
      (await modules.registry.view(text, {format: 'csv'})).sheet(0).rows(),
      (await modules.local.view(text, {format: 'csv'})).sheet(0).rows(),
    );
    console.log('XLSX parity');
    const buffers = {};
    for (const variant of ['registry', 'local'])
      buffers[variant] = await modules[variant].xlsx.writeBuffer(models[variant]);
    for (const variant of ['registry', 'local']) {
      const loaded = await modules[variant].xlsx.load(buffers[variant]);
      equal(loaded.sheets[0].rows[0].cells[1].value, rows[0][0]);
      assert(loaded.sheets[0].rows.length === rows.length, 'XLSX rows');
    }
    const jobs = {
      'csv parse (map identity)': v => modules[v].csv.parse(text, {map: x => x}),
      'csv stringify': v => modules[v].csv.stringify(models[v]),
      'csv view': v => modules[v].view(text, {format: 'csv'}),
      'xlsx write': v => modules[v].xlsx.writeBuffer(models[v]),
      'xlsx load': v => modules[v].xlsx.load(buffers[v]),
    };
    const timings = {};
    for (const [name, job] of Object.entries(jobs)) {
      console.log('timing', name);
      const samples = {registry: [], local: []};
      for (let i = 0; i < 9; i++) {
        for (const variant of i % 2 ? ['local', 'registry'] : ['registry', 'local']) {
          const start = performance.now();
          await job(variant);
          if (i >= 2) samples[variant].push(performance.now() - start);
        }
      }
      timings[name] = Object.fromEntries(
        Object.entries(samples).map(([v, xs]) => [
          v,
          {medianMs: [...xs].sort((a, b) => a - b)[3], samplesMs: xs},
        ]),
      );
    }
    return {
      userAgent: navigator.userAgent,
      rows: 2000,
      cols: 8,
      runs: 7,
      warmup: 2,
      fixturesPassed: fixtures.length,
      checks:
        'CSV parity, malformed input, formatter options, CSV view, XLSX read/write; local bundles run without Node globals',
      timings,
    };
  }, microtaskOnly);
  const report = {
    date: new Date().toISOString(),
    node: process.version,
    esbuild: esbuildVersion,
    baseline:
      'ExcelJS pre-change source + registry fast-csv@5.0.7 + stream-browserify/util/buffer/process/events/string_decoder',
    candidate: 'Local ExcelJS + linked fast-csv/browser; real CSV, no stubs or Node polyfills',
    reports,
    unpolyfilled,
    performance,
  };
  const reportPath = join(out, 'results.json');
  if (microtaskOnly && existsSync(reportPath)) {
    const previous = JSON.parse(readFileSync(reportPath, 'utf8'));
    previous.performanceMicrotask = performance;
    writeFileSync(reportPath, JSON.stringify(previous, null, 2));
  } else {
    writeFileSync(reportPath, JSON.stringify(report, null, 2));
  }
  console.log(JSON.stringify(performance, null, 2));
  console.log(`Results: ${join(out, 'results.json')}`);
} finally {
  if (browser) await browser.close();
  await new Promise(r => server.close(r));
}
