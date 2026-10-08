/** Record comparable bundles and representative workloads for each independent rewrite. */
import * as esbuild from 'esbuild';
import {gzipSync} from 'node:zlib';
import {performance} from 'node:perf_hooks';
import {resolve, dirname, join} from 'node:path';
import {pathToFileURL, fileURLToPath} from 'node:url';
import {writeFileSync} from 'node:fs';
import {execFileSync} from 'node:child_process';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const stage = process.argv[2] || 'baseline';
const directory = resolve(process.argv[3] || join(root, 'dist'));
const output = join(root, 'build/five-rewrites', stage + '.json');
const cases = {
  write: `import {workbook} from ENTRY;export const run=()=>workbook().sheet('Data').rows([['Name','Value'],['A',1]]).writeBuffer();`,
  load: `import {load} from ENTRY;export const run=bytes=>load(bytes);`,
  readRows: `import {readRows} from ENTRY;export const run=bytes=>readRows(bytes,{format:'xlsx'});`,
};
const bundles = {};
for (const [name, contents] of Object.entries(cases)) {
  const common = {
    stdin: {
      contents: contents.replace('ENTRY', JSON.stringify(join(directory, 'excel.js'))),
      resolveDir: root,
      sourcefile: `${name}.js`,
    },
    bundle: true,
    write: false,
    minify: true,
    format: 'esm',
    platform: 'browser',
    metafile: true,
    mainFields: ['browser', 'module', 'main'],
    conditions: ['browser', 'import', 'default'],
    define: {global: 'globalThis'},
    alias: {fs: join(directory, 'lib/shims/fs-browser.js')},
    logLevel: 'error',
  };
  const single = await esbuild.build(common);
  const split = await esbuild.build({
    ...common,
    splitting: true,
    outdir: join(root, 'build/five-rewrites', stage, name),
  });
  const outputs = split.metafile.outputs;
  const entry = Object.keys(outputs).find(path => outputs[path].entryPoint === `${name}.js`);
  if (!entry) throw new Error('Missing bundle entry');
  const initial = new Set();
  const visit = path => {
    if (initial.has(path)) return;
    initial.add(path);
    for (const dep of outputs[path].imports)
      if (!dep.external && dep.kind !== 'dynamic-import') visit(dep.path);
  };
  visit(entry);
  const sum = (paths, gzip) =>
    paths.reduce((total, path) => {
      const bytes = split.outputFiles.find(file => file.path === resolve(root, path)).contents;
      return total + (gzip ? gzipSync(bytes).length : bytes.length);
    }, 0);
  const bytes = single.outputFiles[0].contents;
  bundles[name] = {
    minified: bytes.length,
    gzip: gzipSync(bytes).length,
    initialMinified: sum([...initial], false),
    initialGzip: sum([...initial], true),
    splitTotal: sum(Object.keys(outputs), false),
  };
}
const {workbook, load, readRows} = await import(pathToFileURL(join(directory, 'excel.js')));
const median = values => [...values].sort((a, b) => a - b)[Math.floor(values.length / 2)];
async function measure(fn) {
  for (let i = 0; i < 3; i++) await fn();
  const values = [];
  for (let i = 0; i < 11; i++) {
    const start = performance.now();
    await fn();
    values.push(performance.now() - start);
  }
  return {medianMs: median(values), minMs: Math.min(...values), maxMs: Math.max(...values)};
}
const grid = Array.from({length: 1000}, (_, r) =>
  Array.from({length: 8}, (_, c) => (c === 0 ? `row ${r}` : r * 8 + c)),
);
const build = () => workbook().sheet('Data').rows(grid);
const bytes = await build().writeBuffer();
const runtime = {
  write: await measure(() => build().writeBuffer()),
  load: await measure(() => load(bytes)),
  readRows: await measure(() => readRows(bytes, {format: 'xlsx'})),
};
const tableBytes = await workbook()
  .sheet('S')
  .table({
    name: 'People',
    ref: 'B3',
    columns: [{name: 'Name'}, {name: 'Score'}],
    rows: [
      ['Ada', 1],
      ['Bob', 2],
    ],
  })
  .writeBuffer();
try {
  const model = await load(tableBytes);
  await workbook(model).writeBuffer();
  runtime.tableEdit = {
    works: true,
    ...(await measure(async () =>
      workbook(await load(tableBytes))
        .sheet('S')
        .cell('C4', 3)
        .writeBuffer(),
    )),
  };
} catch (error) {
  runtime.tableEdit = {works: false, error: error.message};
}
// Fresh process: avoids measurements being masked by earlier column-cache warmup.
const addresses = JSON.parse(
  execFileSync(
    process.execPath,
    ['--expose-gc', join(root, 'scripts/bench/address-cache.mjs'), directory],
    {encoding: 'utf8'},
  ),
);
const result = {stage, bundles, runtime, addresses};
writeFileSync(output, JSON.stringify(result, null, 2) + '\n');
console.log(JSON.stringify(result, null, 2));
