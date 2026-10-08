/* oxlint-disable no-console */
import {spawnSync} from 'node:child_process';
import {readFileSync, writeFileSync, mkdirSync} from 'node:fs';
import {resolve, join, dirname} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {gzipSync} from 'node:zlib';
import * as esbuild from 'esbuild';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const dir = join(root, 'build/internal-cleanup');
const median = a => a.sort((a, b) => a - b)[Math.floor(a.length / 2)];
if (process.argv[2] === '--worker') {
  const [stage, mode, count] = process.argv.slice(3);
  const base = join(dir, stage, 'dist');
  const entry = mode.startsWith('names')
    ? 'lib/model/defined-names.js'
    : 'lib/xlsx/xform/sheet/data-validations-xform.js';
  const {default: Constructor} = await import(pathToFileURL(join(base, entry)).href);
  const ref = count === 'grid' ? 'A1:XFD1048576' : `A1:CV${Number(count) / 100}`;
  const {default: colCache} = await import(
    pathToFileURL(join(base, 'lib/utils/col-cache.js')).href
  );
  const cells = mode.endsWith('-cells')
    ? Array.from({length: Number(count)}, (_, n) =>
        colCache.encodeAddress(Math.floor(n / 100) + 1, (n % 100) + 1),
      )
    : [];
  const times = [];
  let entries;
  for (let sample = 0; sample < 5; sample++) {
    if (global.gc) global.gc();
    const start = performance.now();
    const object = new Constructor();
    if (mode.startsWith('names')) {
      if (cells.length) for (const address of cells) object.add(`Data!${address}`, 'DataRange');
      else object.add(`Data!${ref}`, 'DataRange');
      entries = object.model[0].ranges.length;
    } else if (cells.length) {
      const rule = {type: 'whole', formulae: [1, 10]};
      object.model = Object.fromEntries(cells.map(address => [address, rule]));
      const xml = object.toXml(object.model);
      entries = Number(xml.match(/count="(\d+)"/)[1]);
      if (entries !== 1) throw new Error('Dense validation rectangle did not coalesce');
    } else {
      object.parseOpen({name: 'dataValidations', attributes: {}});
      object.parseOpen({name: 'dataValidation', attributes: {type: 'whole', sqref: ref}});
      object.parseClose('dataValidation');
      entries = Object.keys(object.model).length;
    }
    times.push(performance.now() - start);
  }
  console.log(
    JSON.stringify({ms: median(times), peakRssMiB: process.resourceUsage().maxRSS / 1024, entries}),
  );
} else if (process.argv[2] === 'ranges') {
  const stages = process.argv.slice(3);
  const results = [];
  for (const count of ['10000', '100000', 'grid'])
    for (const mode of count === 'grid'
      ? ['validations', 'names']
      : ['validations', 'names', 'validations-cells', 'names-cells'])
      for (let pass = 0; pass < 3; pass++)
        for (const stage of pass % 2 ? [...stages].reverse() : stages) {
          const p = spawnSync(
            process.execPath,
            [
              '--expose-gc',
              '--max-old-space-size=128',
              fileURLToPath(import.meta.url),
              '--worker',
              stage,
              mode,
              count,
            ],
            {encoding: 'utf8', timeout: count === 'grid' ? 1000 : 10000, maxBuffer: 2000},
          );
          const sample = {
            stage,
            mode,
            count,
            pass,
            ...(p.status === 0
              ? JSON.parse(p.stdout)
              : {
                  status:
                    p.error?.code === 'ETIMEDOUT'
                      ? 'exceeded 1s budget'
                      : 'failed under 128MiB heap budget',
                }),
          };
          results.push(sample);
          console.log(JSON.stringify(sample));
        }
  writeFileSync(join(dir, 'ranges.json'), JSON.stringify(results, null, 2) + '\n');
} else if (process.argv[2] === 'bundle') {
  for (const stage of process.argv.slice(3)) {
    const out = join(dir, stage);
    mkdirSync(out, {recursive: true});
    const result = await esbuild.build({
      stdin: {
        contents: `import {workbook} from ${JSON.stringify(join(out, 'dist/excel.js'))};export const run=()=>workbook().sheet('S').row([1,'x']).writeBuffer();`,
        resolveDir: root,
        sourcefile: 'entry.js',
      },
      bundle: true,
      write: false,
      platform: 'browser',
      format: 'esm',
      minify: true,
      metafile: true,
      alias: {
        fs: join(root, 'lib/shims/fs-browser.ts'),
        'node:module': join(root, 'lib/shims/node-module-browser.ts'),
        module: join(root, 'lib/shims/node-module-browser.ts'),
      },
      define: {global: 'globalThis'},
      conditions: ['browser', 'import', 'default'],
      mainFields: ['browser', 'module', 'main'],
      logLevel: 'error',
    });
    const bytes = result.outputFiles[0].contents;
    const sample = {
      bytes: bytes.length,
      gzipBytes: gzipSync(bytes).length,
      modules: Object.keys(result.metafile.inputs).length,
    };
    console.log(JSON.stringify({stage, ...sample}));
    writeFileSync(join(out, 'bundle.json'), JSON.stringify(sample, null, 2) + '\n');
    writeFileSync(join(out, 'bundle-meta.json'), JSON.stringify(result.metafile, null, 2));
  }
}
