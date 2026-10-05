/** Build from a clean directory so removed modules cannot leak into releases. */
import {rmSync} from 'node:fs';
import {execFileSync} from 'node:child_process';
import {createRequire} from 'node:module';
import {dirname, join} from 'node:path';
import {fileURLToPath} from 'node:url';

const root = dirname(dirname(fileURLToPath(import.meta.url)));
rmSync(join(root, 'dist'), {recursive: true, force: true});
const require = createRequire(import.meta.url);
const compiler = join(dirname(require.resolve('typescript/package.json')), 'bin', 'tsc');
execFileSync(process.execPath, [compiler, '-p', join(root, 'tsconfig.build.json')], {
  cwd: root,
  stdio: 'inherit',
});
