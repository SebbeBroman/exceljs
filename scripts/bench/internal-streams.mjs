/* oxlint-disable no-console */
import {spawnSync} from 'node:child_process';
import {Writable} from 'node:stream';
import {createWriteStream, mkdtempSync, rmSync, writeFileSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join, resolve} from 'node:path';
import {pathToFileURL, fileURLToPath} from 'node:url';
const root = resolve(import.meta.dirname, '../..');
if (process.argv[2] === '--worker') {
  const [stage, mode, count, filename] = process.argv.slice(3);
  const {streamWrite} = await import(
    pathToFileURL(join(root, 'build/internal-cleanup', stage, 'dist/node.js')).href
  );
  const n = Number(count);
  let peakQueue = 0,
    bytes = 0;
  const output =
    mode === 'slow'
      ? new Writable({
          highWaterMark: 16384,
          write(chunk, _encoding, callback) {
            bytes += chunk.length;
            peakQueue = Math.max(peakQueue, output.writableLength);
            setTimeout(callback, 1);
          },
        })
      : createWriteStream(filename);
  global.gc?.();
  const started = performance.now();
  await streamWrite(output, {
    useSharedStrings: false,
    created: new Date('2020-01-01T00:00:00Z'),
    modified: new Date('2020-01-01T00:00:00Z'),
    sheets: [
      {
        name: 'Data',
        rows: (function* () {
          for (let r = 0; r < n; r++)
            yield [`row-${r}`, r, r + 1, r + 2, r + 3, r + 4, r + 5, r + 6];
        })(),
      },
    ],
  });
  console.log(
    JSON.stringify({
      ms: performance.now() - started,
      peakRssMiB: process.resourceUsage().maxRSS / 1024,
      peakQueue,
      bytes,
    }),
  );
} else {
  const stages = process.argv.slice(2);
  const dir = mkdtempSync(join(tmpdir(), 'internal-stream-'));
  const results = [];
  try {
    for (const mode of ['file', 'slow'])
      for (const n of [50000, 100000])
        for (let pass = 0; pass < 3; pass++)
          for (const stage of pass % 2 ? [...stages].reverse() : stages) {
            const result = spawnSync(
              process.execPath,
              [
                '--expose-gc',
                fileURLToPath(import.meta.url),
                '--worker',
                stage,
                mode,
                String(n),
                join(dir, 'data.xlsx'),
              ],
              {encoding: 'utf8', timeout: 60000},
            );
            if (result.status !== 0)
              throw new Error(result.stderr || result.stdout || String(result.error));
            const sample = {stage, mode, n, pass, ...JSON.parse(result.stdout.trim())};
            results.push(sample);
            console.log(JSON.stringify(sample));
          }
    writeFileSync(
      join(root, 'build/internal-cleanup/streams.json'),
      JSON.stringify(results, null, 2) + '\n',
    );
  } finally {
    rmSync(dir, {recursive: true, force: true});
  }
}
