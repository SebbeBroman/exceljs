/* oxlint-disable no-console */
/** Isolated streaming timing/RSS comparison against the saved pre-removal build. */
import {spawnSync} from 'node:child_process';
import {mkdtempSync, rmSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join, resolve, dirname} from 'node:path';
import {pathToFileURL, fileURLToPath} from 'node:url';
import {setImmediate} from 'node:timers/promises';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const worker = process.argv[2];
if (worker === '--worker') {
  const [entry, mode, filename, count] = process.argv.slice(3);
  const {streamWrite, streamRead} = await import(pathToFileURL(entry).href);
  const n = Number(count);
  if (global.gc) global.gc();
  const started = performance.now();
  let rows = 0;
  if (mode === 'write') {
    await streamWrite(filename, {
      useSharedStrings: false,
      sheets: [
        {
          name: 'Data',
          rows: (async function* () {
            for (let r = 0; r < n; r++) {
              yield [`row-${r}`, r, r + 1, r + 2, r + 3, r + 4, r + 5, r + 6];
              if (r % 1000 === 0) await setImmediate();
            }
          })(),
        },
      ],
    });
    rows = n;
  } else {
    for await (const row of streamRead(filename)) {
      if (row.values[2] !== rows) throw new Error(`Incorrect row ${rows}`);
      rows++;
    }
    if (rows !== n) throw new Error('Incorrect row count');
  }
  console.log(
    JSON.stringify({
      ms: performance.now() - started,
      peakRssMiB: process.resourceUsage().maxRSS / 1024,
      rows,
    }),
  );
} else {
  const dir = mkdtempSync(join(tmpdir(), 'legacy-bench-'));
  try {
    const results = [];
    for (const n of [50000, 100000])
      for (let pass = 0; pass < 3; pass++) {
        for (const version of pass % 2 ? ['after', 'before'] : ['before', 'after']) {
          const entry = join(
            root,
            version === 'before' ? 'build/legacy-removal-before/dist/node.js' : 'dist/node.js',
          );
          const file = join(dir, `${version}.xlsx`);
          for (const mode of ['write', 'read']) {
            const p = spawnSync(
              process.execPath,
              [
                '--expose-gc',
                fileURLToPath(import.meta.url),
                '--worker',
                entry,
                mode,
                file,
                String(n),
              ],
              {encoding: 'utf8'},
            );
            if (p.status !== 0) throw new Error(p.stderr || p.stdout);
            const result = {n, pass, version, mode, ...JSON.parse(p.stdout.trim())};
            results.push(result);
            console.log(JSON.stringify(result));
          }
        }
      }
    writeFileSync(
      join(root, 'build/legacy-removal-before/stream-results.json'),
      JSON.stringify(results, null, 2),
    );
  } finally {
    rmSync(dir, {recursive: true, force: true});
  }
}
