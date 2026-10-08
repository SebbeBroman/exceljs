/** Check package declarations without adjacent source files masking them. */
import {cpSync, mkdirSync, rmSync, copyFileSync, writeFileSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {createRequire} from 'node:module';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
const consumer = join(root, 'build', 'public-types');
const pkg = join(consumer, 'node_modules', '@sebbebroman', 'exceljs');
rmSync(consumer, {recursive: true, force: true});
mkdirSync(pkg, {recursive: true});
for (const file of ['package.json', 'excel.d.ts', 'node.d.ts', 'protection.d.ts']) {
  copyFileSync(join(root, file), join(pkg, file));
}
cpSync(join(root, 'spec', 'typescript', 'public'), consumer, {recursive: true});
copyFileSync(join(root, 'tsconfig.public.json'), join(consumer, 'tsconfig.json'));
writeFileSync(join(consumer, 'package.json'), '{"private":true,"type":"module"}\n');
const require = createRequire(import.meta.url);
const compiler = join(dirname(require.resolve('typescript/package.json')), 'bin', 'tsc');
execFileSync(process.execPath, [compiler, '-p', join(consumer, 'tsconfig.json')], {
  cwd: consumer,
  stdio: 'inherit',
});
console.log('Public package declaration checks passed (main + /node + /protection).');
