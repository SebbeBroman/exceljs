/* oxlint-disable no-console */
/** Isolated Node process peak RSS for the same saved XLSX (read) or dense rows (write). */
import {readFileSync, writeFileSync, mkdirSync} from 'node:fs';
import {resolve, dirname, join} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {execFileSync} from 'node:child_process';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/doc-bridge-comparison');
const rows = 100000;
if (process.argv[2] === '--worker') {
  const [, , , build, operation] = process.argv;
  const api = await import(pathToFileURL(join(root, build, 'excel.js')).href);
  globalThis.gc();
  const start = performance.now();
  let result;
  if (operation === 'read') result = await api.load(readFileSync(join(out, 'memory-fixture.xlsx')));
  else {
    const grid = Array.from({length: rows}, (_, r) =>
      Array.from({length: 8}, (_, c) => (c ? r * 8 + c : `r${r}`)),
    );
    result = await api.writeBuffer(api.workbook().sheet('data').rows(grid), {
      useStyles: false,
      useSharedStrings: true,
    });
  }
  const milliseconds = performance.now() - start;
  if (operation === 'read' && result.sheets[0].rows.length !== rows)
    throw new Error('Wrong row count');
  console.log(
    JSON.stringify({
      milliseconds,
      peakRssKiB: process.resourceUsage().maxRSS,
      outputSize: operation === 'read' ? result.sheets[0].rows.length : result.length,
    }),
  );
} else {
  mkdirSync(out, {recursive: true});
  const baseline = 'build/doc-bridge-before/dist';
  const api = await import(pathToFileURL(join(root, baseline, 'excel.js')).href);
  const grid = Array.from({length: rows}, (_, r) =>
    Array.from({length: 8}, (_, c) => (c ? r * 8 + c : `r${r}`)),
  );
  writeFileSync(
    join(out, 'memory-fixture.xlsx'),
    await api.writeBuffer(api.workbook().sheet('data').rows(grid), {
      useStyles: false,
      useSharedStrings: true,
    }),
  );
  const samples = {read: {before: [], after: []}, write: {before: [], after: []}};
  for (const operation of ['read', 'write'])
    for (let iteration = 0; iteration < 3; iteration++) {
      for (const name of iteration % 2 ? ['after', 'before'] : ['before', 'after']) {
        const build = name === 'before' ? baseline : 'dist';
        samples[operation][name].push(
          JSON.parse(
            execFileSync(
              process.execPath,
              ['--expose-gc', fileURLToPath(import.meta.url), '--worker', build, operation],
              {encoding: 'utf8'},
            ),
          ),
        );
      }
    }
  const report = {
    date: new Date().toISOString(),
    node: process.version,
    rows,
    cols: 8,
    metric:
      'Peak process RSS; includes runtime, input and output. Independent processes, three runs each.',
    samples,
  };
  writeFileSync(join(out, 'memory-results.json'), JSON.stringify(report, null, 2));
  console.log(JSON.stringify(report, null, 2));
}
