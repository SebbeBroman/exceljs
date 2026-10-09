/** Guard the separation between OOXML parsing and rendering in browser bundles. */
import {build} from 'esbuild';
import {resolve} from 'node:path';
const root = resolve(import.meta.dirname, '..');
for (const [name, contents] of Object.entries({
  load: "import {load} from '@sebbebroman/exceljs'; export const run = bytes => load(bytes);",
  write:
    "import {workbook} from '@sebbebroman/exceljs'; export const run = () => workbook().sheet('S').row([1]).writeBuffer();",
  readRows:
    "import {readRows} from '@sebbebroman/exceljs'; export const run = bytes => readRows(bytes, {format: 'xlsx'});",
})) {
  const result = await build({
    stdin: {contents, resolveDir: root},
    bundle: true,
    write: false,
    minify: true,
    metafile: true,
    platform: 'browser',
    format: 'esm',
    conditions: ['browser', 'import', 'default'],
    mainFields: ['browser', 'module', 'main'],
    alias: {fs: resolve(root, 'dist/lib/shims/fs-browser.js')},
    define: {global: 'globalThis'},
    logLevel: 'error',
  });
  const included = Object.entries(Object.values(result.metafile.outputs)[0].inputs)
    .filter(([, info]) => info.bytesInOutput > 0)
    .map(([path]) => path);
  const forbidden = included.filter(path =>
    name === 'write'
      ? /lib\/xlsx\/(?:parser\/|base-parser\.js|lazy-parsers\.js)/.test(path)
      : name === 'load'
        ? /lib\/(?:xlsx\/xform\/(?!xform-state\.js)|utils\/xml-stream\.js|utils\/buffer-zip\.js)/.test(
            path,
          )
        : /lib\/xlsx\/(?:parser\/|xform\/|xlsx-reader\.js|xlsx-writer\.js)/.test(path),
  );
  if (forbidden.length)
    throw new Error(`${name} retains the other transform direction:\n${forbidden.join('\n')}`);
  console.log(`${name}: directional bundle guard passed`);
}
