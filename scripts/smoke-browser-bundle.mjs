/**
 * Bundle for browser WITHOUT process polyfills.
 * Asserts the client path works and does not depend on a live `process` object.
 */
import * as esbuild from 'esbuild';
import {writeFileSync, mkdirSync, readFileSync} from 'node:fs';
import {join, dirname} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const outDir = join(root, 'build', 'browser-smoke');
mkdirSync(outDir, {recursive: true});

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
`
);

const outfile = join(outDir, 'bundle.mjs');

await esbuild.build({
  entryPoints: [entry],
  bundle: true,
  outfile,
  format: 'esm',
  platform: 'browser',
  mainFields: ['browser', 'module', 'main'],
  conditions: ['browser', 'import', 'default'],
  define: {
    global: 'globalThis',
  },
  alias: {
    fs: join(root, 'lib/shims/fs-browser.js'),
    crypto: join(root, 'lib/shims/crypto-browser.js'),
  },
  // No process/buffer inject — Buffer comes from the `buffer` package imports in source.
  logLevel: 'warning',
});

const code = readFileSync(outfile, 'utf8');
// Our own code must not call process.nextTick; jszip may keep a dead-branch reference.
if (code.includes('from "readable-stream"') || code.includes("from 'readable-stream'")) {
  throw new Error('bundle still imports readable-stream');
}

const mod = await import(pathToFileURL(outfile).href + `?t=${Date.now()}`);
const result = await mod.run();
const size = readFileSync(outfile).byteLength;
console.log('browser bundle smoke ok', result, `bundle=${(size / 1024).toFixed(1)}KB`);
console.log('  (no process polyfill, no readable-stream, fflate zip)');
