/** Create and inspect a local release tarball after release:check. Does not publish. */
import {execSync} from 'node:child_process';
import {mkdirSync, readFileSync, writeFileSync} from 'node:fs';
import {resolve} from 'node:path';
import assert from 'node:assert/strict';

const root = resolve(import.meta.dirname, '..');
const destination = resolve(root, 'build/release');
mkdirSync(destination, {recursive: true});
const pkg = JSON.parse(readFileSync(resolve(root, 'package.json'), 'utf8'));
// Fixed command, with no interpolated shell input; npm also works on Windows.
const [packed] = JSON.parse(
  execSync('npm pack --ignore-scripts --json --pack-destination build/release', {
    cwd: root,
    encoding: 'utf8',
  }),
);
assert.equal(packed.name, pkg.name);
assert.equal(packed.version, pkg.version);
const files = new Set(packed.files.map(file => file.path));
for (const entry of Object.values(pkg.exports)) {
  const targets = typeof entry === 'string' ? [entry] : Object.values(entry);
  for (const target of targets)
    assert.ok(files.has(target.replace(/^\.\//, '')), `Missing export: ${target}`);
}
for (const file of ['README.md', 'CHANGELOG.md', 'MIGRATION.md', 'ARCHITECTURE.md', 'LICENSE']) {
  assert.ok(files.has(file), `Missing release documentation: ${file}`);
}
const leaked = [...files].filter(path =>
  /^(?:spec|build|node_modules)\/|^dist\/lib\/doc\/|^index\.d\.ts$/.test(path),
);
assert.deepEqual(leaked, [], 'Legacy or development files leaked into the release');
writeFileSync(resolve(destination, 'pack-metadata.json'), JSON.stringify(packed, null, 2) + '\n');
console.log(`Checked ${pkg.name}@${pkg.version}: ${files.size} files, ${packed.size} packed bytes`);
console.log(resolve(destination, packed.filename));
