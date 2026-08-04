/**
 * Bundle for browser WITHOUT process polyfills.
 * Asserts the client write path works and does not depend on a live `process` object.
 * Reports minified single-file size and code-split entry chunk sizes (write-only).
 *
 * CSV is optional (dynamic import on builder `.csv()`). For write-only size we
 * stub `lib/csv/public` so the measurement matches a bundler that splits/drops
 * unused CSV — and so browser builds do not need Node builtins from fast-csv.
 */
import * as esbuild from 'esbuild';
import {writeFileSync, mkdirSync, readFileSync, readdirSync, statSync} from 'node:fs';
import {join, dirname} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {gzipSync} from 'node:zlib';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'build', 'browser-smoke');
mkdirSync(outDir, {recursive: true});

const fsAlias = join(root, 'lib/shims/fs-browser.ts');

/** Drop optional CSV module from write-only browser measurements. */
const writeOnlyCsvStubPlugin = {
  name: 'write-only-csv-stub',
  setup(build) {
    build.onResolve({filter: /csv\/public(\.ts|\.js)?$/}, args => ({
      path: args.path,
      namespace: 'csv-stub',
    }));
    build.onLoad({filter: /.*/, namespace: 'csv-stub'}, () => ({
      contents: `
        export async function parseCsv() {
          throw new Error('csv omitted from write-only browser smoke');
        }
        export async function stringifyCsv() {
          throw new Error('csv omitted from write-only browser smoke');
        }
        export const csv = { parse: parseCsv, stringify: stringifyCsv };
      `,
      loader: 'js',
    }));
  },
};

const sharedBuild = {
  bundle: true,
  format: 'esm',
  platform: 'browser',
  mainFields: ['browser', 'module', 'main'],
  conditions: ['browser', 'import', 'default'],
  resolveExtensions: ['.ts', '.js', '.mjs', '.json'],
  define: {
    global: 'globalThis',
  },
  alias: {
    fs: fsAlias,
    'node:module': join(root, 'lib/shims/node-module-browser.ts'),
    module: join(root, 'lib/shims/node-module-browser.ts'),
  },
  plugins: [writeOnlyCsvStubPlugin],
  logLevel: 'warning',
};

const writeOnlyEntry = join(outDir, 'entry-write-only.js');
writeFileSync(
  writeOnlyEntry,
  `
import { workbook } from '../../excel.ts';

export async function run() {
  const buf = await workbook()
    .sheet('Browser')
    .cell('A1', 'hello browser')
    .cell('B1', 99)
    .writeBuffer();
  return { ok: true, bytes: buf.length || buf.byteLength };
}
`,
);

// --- Single-file minified ---
const outfile = join(outDir, 'bundle.mjs');
await esbuild.build({
  ...sharedBuild,
  entryPoints: [writeOnlyEntry],
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
if (code.includes('fast-csv') || code.includes('@fast-csv')) {
  throw new Error('write-only bundle still pulls fast-csv');
}

const mod = await import(pathToFileURL(outfile).href + `?t=${Date.now()}`);
const result = await mod.run();
const size = readFileSync(outfile).byteLength;
const gzip = gzipSync(readFileSync(outfile)).byteLength;
const singleLine = `single-file minify: ${(size / 1024).toFixed(1)}KB  gzip ${(gzip / 1024).toFixed(1)}KB`;
console.log('browser bundle smoke ok', result);
console.log(`  ${singleLine}`);

// --- Code-split ---
async function measureSplit(label, entryPoint) {
  const splitDir = join(outDir, `split-${label}`);
  mkdirSync(splitDir, {recursive: true});

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
  for (const f of files) {
    const b = statSync(join(splitDir, f)).size;
    totalBytes += b;
    if (f.startsWith('entry')) entryBytes = b;
  }
  const line = `split/${label}: entry ${(entryBytes / 1024).toFixed(1)}KB  total ${(totalBytes / 1024).toFixed(1)}KB  (${files.length} files)`;
  console.log(`  ${line}`);
  return {entryBytes, totalBytes, line};
}

const split = await measureSplit('write-only', writeOnlyEntry);
console.log('  (builder write-only path; csv stubbed; node entry not imported)');

// Persist sizes for docs
const sizesPath = join(outDir, 'sizes.txt');
writeFileSync(
  sizesPath,
  [
    singleLine,
    split.line,
    `bytes: single=${size} gzip=${gzip} splitEntry=${split.entryBytes} splitTotal=${split.totalBytes}`,
  ].join('\n') + '\n',
);
console.log(`  wrote ${sizesPath}`);
