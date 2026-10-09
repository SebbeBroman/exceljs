/** Browser dependency and size regression tests, using fresh in-memory builds. */
import {build} from 'esbuild';
import {resolve} from 'node:path';
import {readFileSync, mkdirSync, writeFileSync} from 'node:fs';
import {gzipSync} from 'node:zlib';
import assert from 'node:assert/strict';
const root = resolve(import.meta.dirname, '..');
const budgets = JSON.parse(
  readFileSync(resolve(root, 'scripts/browser-bundle-budgets.json'), 'utf8'),
);
const measurements = {};
for (const [name, contents] of Object.entries({
  load: "import {load} from '@sebbebroman/exceljs'; export const run = bytes => load(bytes);",
  write:
    "import {workbook} from '@sebbebroman/exceljs'; export const run = () => workbook().sheet('S').row([1]).writeBuffer();",
  viewWorkbook:
    "import {viewWorkbook} from '@sebbebroman/exceljs'; export const run = bytes => viewWorkbook(bytes);",
  readRows:
    "import {readRows} from '@sebbebroman/exceljs'; export const run = bytes => readRows(bytes, {format: 'xlsx'});",
})) {
  const options = {
    stdin: {contents, resolveDir: root, sourcefile: `${name}.js`},
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
  };
  const result = await build(options);
  const included = Object.entries(Object.values(result.metafile.outputs)[0].inputs)
    .filter(([, info]) => info.bytesInOutput > 0)
    .map(([path]) => path);
  const forbidden = included.filter(
    path =>
      /(?:@noble|utils\/encryptor|protection\/sheet-protection|fast-csv|dayjs|lib\/csv\/)/.test(
        path,
      ) ||
      (name === 'write'
        ? /lib\/xlsx\/(?:parser\/|base-parser\.js|lazy-parsers\.js)/.test(path)
        : name === 'load'
          ? /lib\/(?:xlsx\/xform\/(?!xform-state\.js)|utils\/xml-stream\.js|utils\/buffer-zip\.js)/.test(
              path,
            )
          : /lib\/xlsx\/(?:parser\/|xform\/|xlsx-reader\.js|xlsx-writer\.js)/.test(path)),
  );
  if (forbidden.length)
    throw new Error(`${name} retains forbidden dependencies:\n${forbidden.join('\n')}`);
  const split = await build({
    ...options,
    splitting: true,
    outdir: resolve(root, 'build/bundle-regression', name),
  });
  const outputs = split.metafile.outputs;
  const entry = Object.keys(outputs).find(path => outputs[path].entryPoint === `${name}.js`);
  assert.ok(entry, `${name}: missing split entry`);
  function closure(dynamic) {
    const seen = new Set();
    function visit(path) {
      if (seen.has(path)) return;
      seen.add(path);
      for (const dep of outputs[path].imports) {
        if (!dep.external && (dynamic || dep.kind !== 'dynamic-import')) visit(dep.path);
      }
    }
    visit(entry);
    return [...seen];
  }
  function sum(paths, gzip) {
    return paths.reduce((total, path) => {
      const file = split.outputFiles.find(file => file.path === resolve(root, path));
      assert.ok(file, `${name}: missing chunk ${path}`);
      return total + (gzip ? gzipSync(file.contents).length : file.contents.length);
    }, 0);
  }
  const single = result.outputFiles[0].contents;
  const initial = closure(false);
  const reachable = closure(true);
  const sizes = {
    minified: single.length,
    gzip: gzipSync(single).length,
    initialMinified: sum(initial, false),
    initialGzip: sum(initial, true),
    reachableMinified: sum(reachable, false),
    reachableGzip: sum(reachable, true),
  };
  assert.deepEqual(
    Object.keys(budgets[name]).sort(),
    Object.keys(sizes).sort(),
    `${name}: incomplete budgets`,
  );
  for (const [metric, size] of Object.entries(sizes)) {
    assert.ok(
      size <= budgets[name][metric],
      `${name}.${metric}: ${size} bytes exceeds ${budgets[name][metric]} byte budget; review growth before updating scripts/browser-bundle-budgets.json`,
    );
  }
  measurements[name] = sizes;
  console.log(
    `${name}: direction and size budgets passed (${sizes.minified} bytes / ${sizes.gzip} gzip)`,
  );
}

assert.deepEqual(
  Object.keys(measurements).sort(),
  Object.keys(budgets).sort(),
  'Unmeasured bundle budget',
);
mkdirSync(resolve(root, 'build/bundle-regression'), {recursive: true});
writeFileSync(
  resolve(root, 'build/bundle-regression/sizes.json'),
  JSON.stringify(measurements, null, 2) + '\n',
);
