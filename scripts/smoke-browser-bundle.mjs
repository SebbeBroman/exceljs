/**
 * Bundle for browser WITHOUT process polyfills.
 * Asserts the client path works and does not depend on a live `process` object.
 * Reports minified single-file size and code-split entry chunk sizes (write-only vs round-trip).
 */
import * as esbuild from 'esbuild';
import {writeFileSync, mkdirSync, readFileSync, readdirSync, statSync} from 'node:fs';
import {join, dirname} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {gzipSync} from 'node:zlib';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'build', 'browser-smoke');
mkdirSync(outDir, {recursive: true});

const fsAlias = join(root, 'lib/shims/fs-browser.js');
const cryptoAlias = join(root, 'lib/shims/crypto-browser.js');

const sharedBuild = {
  bundle: true,
  format: 'esm',
  platform: 'browser',
  mainFields: ['browser', 'module', 'main'],
  conditions: ['browser', 'import', 'default'],
  define: {
    global: 'globalThis',
  },
  alias: {
    fs: fsAlias,
    crypto: cryptoAlias,
    'node:module': join(root, 'lib/shims/node-module-browser.js'),
    module: join(root, 'lib/shims/node-module-browser.js'),
  },
  logLevel: 'warning',
};

const entry = join(outDir, 'entry.js');
writeFileSync(
  entry,
  `
import { Workbook } from '../../excel.js';

export async function run() {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Browser');
  ws.getCell('A1').value = 'hello browser';
  ws.getCell('B1').value = 99;
  const buf = await wb.xlsx.writeBuffer();
  const wb2 = new Workbook();
  await wb2.xlsx.load(buf);
  const v = wb2.getWorksheet('Browser').getCell('A1').value;
  if (v !== 'hello browser') throw new Error('round-trip failed: ' + v);
  return { ok: true, bytes: buf.length || buf.byteLength };
}
`,
);

const writeOnlyEntry = join(outDir, 'entry-write-only.js');
writeFileSync(
  writeOnlyEntry,
  `
import { Workbook } from '../../excel.js';

export async function run() {
  const wb = new Workbook();
  const ws = wb.addWorksheet('Browser');
  ws.getCell('A1').value = 'hello browser';
  ws.getCell('B1').value = 99;
  const buf = await wb.xlsx.writeBuffer();
  return { ok: true, bytes: buf.length || buf.byteLength };
}
`,
);

// --- Single-file minified (everything that dynamic-import can still pull in is included) ---
const outfile = join(outDir, 'bundle.mjs');
await esbuild.build({
  ...sharedBuild,
  entryPoints: [entry],
  outfile,
  minify: true,
});

const code = readFileSync(outfile, 'utf8');
if (code.includes('from "readable-stream"') || code.includes("from 'readable-stream'")) {
  throw new Error('bundle still imports readable-stream');
}
if (
  code.includes('node_modules/buffer/') ||
  code.includes('from "buffer"') ||
  code.includes("from 'buffer'")
) {
  throw new Error('bundle still imports npm buffer package');
}
if (
  code.includes('node_modules/events/') ||
  code.includes('from "events"') ||
  code.includes("from 'events'") ||
  code.includes('from "node:events"') ||
  code.includes("from 'node:events'")
) {
  throw new Error('bundle still imports npm events / node:events package');
}

const mod = await import(pathToFileURL(outfile).href + `?t=${Date.now()}`);
const result = await mod.run();
const size = readFileSync(outfile).byteLength;
const gzip = gzipSync(readFileSync(outfile)).byteLength;
console.log('browser bundle smoke ok', result);
console.log(
  `  single-file minify: ${(size / 1024).toFixed(1)}KB  gzip ${(gzip / 1024).toFixed(1)}KB`,
);
console.log('  (no process polyfill, no npm buffer/events, no readable-stream, fflate zip)');

// --- Code-split: entry chunk is what apps pay for initially ---
async function measureSplit(label, entryPoint) {
  const splitDir = join(outDir, `split-${label}`);
  mkdirSync(splitDir, {recursive: true});
  // clear previous chunks
  for (const f of readdirSync(splitDir)) {
    try {
      // only files
      if (statSync(join(splitDir, f)).isFile()) {
        // leave dir, rebuild overwrites
      }
    } catch {
      /* ignore */
    }
  }

  await esbuild.build({
    ...sharedBuild,
    entryPoints: [entryPoint],
    outdir: splitDir,
    entryNames: 'entry',
    chunkNames: 'chunk-[hash]',
    splitting: true,
    minify: true,
  });

  const files = readdirSync(splitDir).filter(f => f.endsWith('.js'));
  let entryBytes = 0;
  let totalBytes = 0;
  const chunks = [];
  for (const f of files) {
    const b = statSync(join(splitDir, f)).size;
    totalBytes += b;
    chunks.push({f, b});
    if (f.startsWith('entry')) entryBytes = b;
  }
  chunks.sort((a, b) => b.b - a.b);
  console.log(
    `  split/${label}: entry ${(entryBytes / 1024).toFixed(1)}KB  total ${(totalBytes / 1024).toFixed(1)}KB  (${files.length} files)`,
  );
  return {entryBytes, totalBytes, chunks};
}

const writeSplit = await measureSplit('write-only', writeOnlyEntry);
const roundTripSplit = await measureSplit('round-trip', entry);

// Sanity: write-only entry should not retain saxes / CF if lazy load works
const writeEntryCode = readFileSync(join(outDir, 'split-write-only', 'entry.js'), 'utf8');
const saxesInEntry = writeEntryCode.includes('SaxesParser');
const cfInEntry =
  writeEntryCode.includes('conditionalFormattings') && writeEntryCode.includes('cf-rule');
// table.js exclusive string (styles also mention TableStyleMedium2)
const tableDocInEntry = writeEntryCode.includes('Invalid Totals Row Function');
console.log(
  `  write-only: saxes in entry=${saxesInEntry} cf-ish=${cfInEntry} table-doc=${tableDocInEntry}  optional chunks total~${((writeSplit.totalBytes - writeSplit.entryBytes) / 1024).toFixed(1)}KB`,
);
console.log(
  `  round-trip entry vs write-only entry delta: ${((roundTripSplit.entryBytes - writeSplit.entryBytes) / 1024).toFixed(1)}KB`,
);
