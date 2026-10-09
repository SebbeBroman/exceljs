/** Alternate baseline/candidate order to compare directional transform runtime. */
import {performance} from 'node:perf_hooks';
import {resolve} from 'node:path';
import {pathToFileURL} from 'node:url';
// Run from the repository root, passing the baseline and candidate dist directories.
const before = await import(pathToFileURL(resolve(process.argv[2], 'excel.js')));
const after = await import(pathToFileURL(resolve(process.argv[3] || 'dist', 'excel.js')));
const grid = Array.from({length: 1000}, (_, r) =>
  Array.from({length: 8}, (_, c) => (c ? r * 8 + c : `row ${r}`)),
);
const bytes = await before.workbook().sheet('Data').rows(grid).writeBuffer();
const result = {};
for (const name of ['write', 'load', 'readRows']) {
  const run = mod =>
    name === 'write'
      ? mod.workbook().sheet('Data').rows(grid).writeBuffer()
      : name === 'load'
        ? mod.load(bytes)
        : mod.readRows(bytes, {format: 'xlsx'});
  for (let i = 0; i < 5; i++) {
    await run(before);
    await run(after);
  }
  const samples = {before: [], after: []};
  for (let i = 0; i < 31; i++) {
    for (const [label, mod] of i % 2
      ? [
          ['after', after],
          ['before', before],
        ]
      : [
          ['before', before],
          ['after', after],
        ]) {
      const start = performance.now();
      await run(mod);
      samples[label].push(performance.now() - start);
    }
  }
  result[name] = Object.fromEntries(
    Object.entries(samples).map(([key, values]) => [
      key,
      {
        medianMs: [...values].sort((a, b) => a - b)[15],
        minMs: Math.min(...values),
        maxMs: Math.max(...values),
        samples: values,
      },
    ]),
  );
}
console.log(JSON.stringify(result, null, 2));
