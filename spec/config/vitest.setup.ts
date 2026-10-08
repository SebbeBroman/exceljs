import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

// Shared output directory for transform fixtures.
const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');
fs.mkdirSync(path.join(root, 'spec/out'), {recursive: true});
