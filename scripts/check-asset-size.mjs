/**
 * Asset size gate for the browser bundle.
 *
 * Replaces chrysanthos/simple-asset-size-reporter@1.0.2, which is unmaintained
 * (last code change April 2020) and crashes on modern runners: it loads the
 * `esm` shim, which patches CommonJS module loading and does not work on the
 * Node 24 that GitHub forces Actions onto. That action has failed every run in
 * this repo's history, including on master.
 *
 * This runs after `pnpm run test:browser-bundle`, reads the byte counts that
 * script already persists to build/browser-smoke/sizes.txt, and fails when any
 * of them regresses past its budget.
 *
 * Budgets are absolute ceilings, not a diff against the base branch. That keeps
 * the gate deterministic and lets it catch an accidental dependency landing in
 * the browser path, which is the regression this repo cares about. Raise a
 * budget deliberately in the same commit that changes bundle contents.
 *
 * Usage:
 *   node scripts/check-asset-size.mjs             # check against budgets
 *   node scripts/check-asset-size.mjs --report    # print sizes, always exit 0
 */
import {readFileSync} from 'node:fs';
import {join, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';

const root = join(dirname(fileURLToPath(import.meta.url)), '..');
const sizesPath = join(root, 'build', 'browser-smoke', 'sizes.txt');

/**
 * Ceilings in bytes. Single-file is the number consumers feel: what a bundler
 * ships when they import the package. gzip is what crosses the wire. The split
 * totals are the same code split into chunks, so they should track single.
 */
const BUDGETS = {
  single: 400_000,
  gzip: 120_000,
  splitEntry: 250_000,
  splitTotal: 400_000,
};

const reportOnly = process.argv.includes('--report');

let raw;
try {
  raw = readFileSync(sizesPath, 'utf8');
} catch {
  console.error(`Asset size check: cannot read ${sizesPath}.`);
  console.error(
    'Run `pnpm run test:browser-bundle` first — it builds the bundle and writes sizes.',
  );
  process.exit(2);
}

const bytesLine = raw.split('\n').find(l => l.startsWith('bytes:'));
if (!bytesLine) {
  console.error(`Asset size check: no "bytes:" line in ${sizesPath}.`);
  process.exit(2);
}

const sizes = Object.fromEntries(
  bytesLine
    .slice('bytes:'.length)
    .trim()
    .split(/\s+/)
    .map(pair => {
      const [key, value] = pair.split('=');
      return [key, Number(value)];
    }),
);

if (Object.values(sizes).some(n => !Number.isFinite(n))) {
  console.error(`Asset size check: could not parse sizes from "${bytesLine}".`);
  process.exit(2);
}

const kb = n => `${(n / 1024).toFixed(1)}KB`;
const overBudget = [];

console.log('Asset size (browser single-file bundle):');
for (const line of raw.split('\n').filter(l => l && !l.startsWith('bytes:'))) {
  console.log(`  ${line}`);
}
console.log('');
for (const [key, limit] of Object.entries(BUDGETS)) {
  const actual = sizes[key];
  if (actual === undefined) {
    console.log(`  ${key.padEnd(11)} not measured — skipped`);
    continue;
  }
  const over = actual > limit;
  const pct = ((actual / limit) * 100).toFixed(0);
  const status = over ? 'OVER' : 'ok  ';
  console.log(
    `  ${status} ${key.padEnd(11)} ${kb(actual).padStart(9)} / ${kb(limit)} budget (${pct}%)`,
  );
  if (over) {
    overBudget.push(`${key}: ${kb(actual)} exceeds ${kb(limit)}`);
  }
}

if (overBudget.length > 0) {
  console.log('');
  if (reportOnly) {
    console.log('Asset size over budget (report-only, not failing):');
    for (const line of overBudget) console.log(`  - ${line}`);
    process.exit(0);
  }
  console.error('Asset size budget exceeded:');
  for (const line of overBudget) console.error(`  - ${line}`);
  console.error('');
  console.error('If this growth is intended, raise the budget in scripts/check-asset-size.mjs');
  console.error('and explain why in the same commit.');
  process.exit(1);
}

console.log('');
console.log('Asset size within budget.');
