/* oxlint-disable no-console */
/** Compare public projections against the saved baseline.
 * The large fixture is checked separately; whole-grid validation expansion in
 * test-issue-1842 stalls the existing parser before either projection runs.
 * Ignore the private worksheet backpointer on image anchors.
 */
import {readFileSync, readdirSync, writeFileSync} from 'node:fs';
import {join, resolve, dirname} from 'node:path';
import {fileURLToPath, pathToFileURL} from 'node:url';
import {isDeepStrictEqual} from 'node:util';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const before = await import(
  pathToFileURL(join(root, 'build/doc-bridge-before/dist/excel.js')).href
);
const after = await import(pathToFileURL(join(root, 'dist/excel.js')).href);
function normalize(x) {
  return JSON.parse(JSON.stringify(x, (key, value) => (key === 'worksheet' ? undefined : value)));
}
function difference(a, b, path = '') {
  if (isDeepStrictEqual(a, b)) return null;
  if (!a || !b || typeof a !== 'object' || typeof b !== 'object')
    return {path, before: a, after: b};
  for (const k of new Set([...Object.keys(a), ...Object.keys(b)])) {
    const d = difference(a[k], b[k], path + '.' + k);
    if (d) return d;
  }
  return {path, before: a, after: b};
}
const results = [];
for (const f of readdirSync(join(root, 'spec/integration/data'), {recursive: true}).filter(
  f => f.endsWith('.xlsx') && !['huge.xlsx', 'test-issue-1842.xlsx'].includes(f),
)) {
  const data = readFileSync(join(root, 'spec/integration/data', f));
  let a, b, ae, be;
  try {
    a = normalize(await before.load(data));
  } catch (e) {
    ae = e.message;
  }
  try {
    b = normalize(await after.load(data));
  } catch (e) {
    be = e.message;
  }
  const d = ae || be ? (ae === be ? null : {beforeError: ae, afterError: be}) : difference(a, b);
  console.log(f, d ? JSON.stringify(d).slice(0, 900) : ae ? 'same rejection: ' + ae : 'match');
  results.push({file: f, match: !d, baselineError: ae, difference: d});
}
writeFileSync(
  join(root, 'build/doc-bridge-comparison/fixture-results.json'),
  JSON.stringify(results, null, 2),
);

if (results.some(result => !result.match)) process.exitCode = 1;
