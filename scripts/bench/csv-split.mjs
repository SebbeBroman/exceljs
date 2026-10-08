/** Compare browser bundles before/after isolating CSV (pass the saved baseline dist directory). */
import * as esbuild from 'esbuild';
import {gzipSync} from 'node:zlib';
import {resolve, dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';
import {writeFileSync} from 'node:fs';

const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const before = resolve(process.argv[2] || join(root, 'build/optional-features-before/dist'));
const cases = {
  write: `import {workbook} from ENTRY; export const run=()=>workbook().sheet('Data').rows([['Name','Value'],['A',1]]).writeBuffer();`,
  readRows: `import {readRows} from ENTRY; export const run=bytes=>readRows(bytes,{format:'xlsx'});`,
  load: `import {load} from ENTRY; export const run=bytes=>load(bytes);`,
};
const results = {};
for (const [name, code] of Object.entries(cases)) {
  const sample = {};
  for (const [label, directory] of [
    ['before', before],
    ['after', join(root, 'dist')],
  ]) {
    const bundle = await esbuild.build({
      stdin: {
        contents: code.replace('ENTRY', JSON.stringify(join(directory, 'excel.js'))),
        resolveDir: root,
      },
      bundle: true,
      write: false,
      minify: true,
      platform: 'browser',
      format: 'esm',
      metafile: true,
      conditions: ['browser', 'import', 'default'],
      mainFields: ['browser', 'module', 'main'],
      define: {global: 'globalThis'},
      alias: {
        fs: join(root, 'dist/lib/shims/fs-browser.js'),
        'node:module': join(root, 'dist/lib/shims/node-module-browser.js'),
        module: join(root, 'dist/lib/shims/node-module-browser.js'),
      },
      logLevel: 'error',
    });
    const bytes = bundle.outputFiles[0].contents;
    const csv = Object.values(bundle.metafile.outputs).some(output =>
      Object.entries(output.inputs).some(
        ([path, info]) => info.bytesInOutput > 0 && /fast-csv|dayjs|lib\/csv\//.test(path),
      ),
    );
    if (label === 'after' && csv) throw new Error(`Core ${name} bundle includes CSV`);
    sample[label] = {minified: bytes.length, gzip: gzipSync(bytes).length, csv};
  }
  sample.saved = {
    minified: sample.before.minified - sample.after.minified,
    gzip: sample.before.gzip - sample.after.gzip,
  };
  results[name] = sample;
}
writeFileSync(
  join(root, 'scripts/bench/csv-split-results.json'),
  JSON.stringify(results, null, 2) + '\n',
);
console.log(JSON.stringify(results, null, 2));
