import {readFileSync, writeFileSync, existsSync} from 'node:fs';
import {resolve} from 'node:path';
const root = resolve(import.meta.dirname, '../..');
const output = resolve(root, 'scripts/bench/five-rewrites-results.json');
const results = existsSync(output)
  ? JSON.parse(readFileSync(output, 'utf8'))
  : {
      method:
        'Browser ESM, esbuild minification + gzip; initial includes static imports, excludes lazy chunks. Runtime: median of 11 after 3 warmups, 1000 x 8 cells; timings are indicative, not regression thresholds.',
      stages: {},
    };
for (const stage of process.argv.slice(2))
  results.stages[stage] = JSON.parse(
    readFileSync(resolve(root, `build/five-rewrites/${stage}.json`), 'utf8'),
  );
writeFileSync(output, JSON.stringify(results, null, 2) + '\n');
