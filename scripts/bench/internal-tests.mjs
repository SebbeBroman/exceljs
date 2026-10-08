/* oxlint-disable no-console */
import {spawnSync} from 'node:child_process';
import {readFileSync, writeFileSync, mkdirSync} from 'node:fs';
import {join, resolve} from 'node:path';
const root = resolve(import.meta.dirname, '../..');
const results = [];
for (let pass = 0; pass < 3; pass++) {
  for (const stage of pass % 2 ? ['stage4', 'stage3'] : ['stage3', 'stage4']) {
    const cwd = stage === 'stage3' ? join(root, 'build/internal-cleanup/stage3/source') : root;
    const out = join(root, 'build/internal-cleanup', stage);
    mkdirSync(out, {recursive: true});
    const report = join(out, `tests-${pass}.json`);
    const start = performance.now();
    const p = spawnSync(
      join(root, 'node_modules/.bin/vp'),
      ['test', 'run', '--reporter=default', '--reporter=json', `--outputFile=${report}`],
      {cwd, encoding: 'utf8', timeout: 60000, maxBuffer: 1000000},
    );
    const wallMs = performance.now() - start;
    writeFileSync(join(out, `tests-${pass}.txt`), p.stdout + p.stderr);
    if (p.status !== 0) throw new Error((p.stdout + p.stderr).slice(-4000));
    const j = JSON.parse(readFileSync(report, 'utf8'));
    const text = p.stdout.replace(/\x1b\[[0-9;]*m/g, '');
    const duration = text.match(
      /Duration\s+([\d.]+)s \(transform ([\d.]+)(m?s), setup ([\d.]+)(m?s), import ([\d.]+)(m?s)/,
    );
    const sample = {stage, pass, wallMs, tests: j.numPassedTests, files: j.testResults.length};
    if (duration) {
      sample.suiteMs = Number(duration[1]) * 1000;
      sample.setupMs = Number(duration[4]) * (duration[5] === 's' ? 1000 : 1);
      sample.importMs = Number(duration[6]) * (duration[7] === 's' ? 1000 : 1);
    }
    results.push(sample);
    console.log(JSON.stringify(sample));
  }
}
writeFileSync(
  join(root, 'build/internal-cleanup/tests.json'),
  JSON.stringify(results, null, 2) + '\n',
);
