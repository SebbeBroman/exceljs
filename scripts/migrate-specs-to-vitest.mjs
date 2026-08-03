#!/usr/bin/env node
/**
 * Convert spec files from CJS to ESM for Vitest.
 * Leaves verquire()/expect()/describe as globals (injected by vitest.setup.js).
 */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '..');
const specRoot = path.join(root, 'spec');

const SKIP_DIRS = new Set(['config', 'browser', 'dist', 'manual', 'typescript', 'out']);

function walk(dir, files = []) {
  for (const ent of fs.readdirSync(dir, {withFileTypes: true})) {
    if (SKIP_DIRS.has(ent.name)) continue;
    const full = path.join(dir, ent.name);
    if (ent.isDirectory()) walk(full, files);
    else if (ent.name.endsWith('.js')) files.push(full);
  }
  return files;
}

function ensureJsExt(spec, fromFile) {
  if (!spec.startsWith('.')) return spec;
  if (spec.endsWith('.js') || spec.endsWith('.json')) return spec;
  const abs = path.resolve(path.dirname(fromFile), spec);
  if (fs.existsSync(`${abs}.js`)) return `${spec}.js`;
  if (fs.existsSync(abs) && fs.statSync(abs).isDirectory()) {
    if (fs.existsSync(path.join(abs, 'index.js'))) return `${spec}/index.js`;
  }
  if (fs.existsSync(`${abs}.json`)) return `${spec}.json`;
  // default assume .js
  return `${spec}.js`;
}

function coreMod(spec) {
  const cores = new Set([
    'fs', 'path', 'os', 'util', 'stream', 'url', 'crypto', 'buffer', 'events',
    'assert', 'child_process', 'http', 'https', 'zlib', 'module',
  ]);
  if (cores.has(spec)) return `node:${spec}`;
  return spec;
}

function convertFile(filePath) {
  let src = fs.readFileSync(filePath, 'utf8');
  if (src.includes('@vitest-migrated')) return false;

  // Skip pure binary-looking base64 data files that only module.exports a string
  const isDataExportOnly =
    /module\.exports\s*=\s*['"`]/.test(src) && !/require\s*\(/.test(src);

  const imports = [];
  const seen = new Set();
  const add = line => {
    if (!seen.has(line)) {
      seen.add(line);
      imports.push(line);
    }
  };

  // Remove chai imports (global expect)
  src = src.replace(/^const\s+\{\s*expect\s*\}\s*=\s*require\(['"]chai['"]\);\s*\n?/gm, '');
  src = src.replace(/^const\s+chai\s*=\s*require\(['"]chai['"]\);\s*\n?/gm, '');
  src = src.replace(/^const\s+\{\s*expect\s*\}\s*=\s*chai;\s*\n?/gm, '');
  src = src.replace(/^const\s+verquire\s*=\s*require\(['"][^'"]*verquire['"]\);\s*\n?/gm, '');

  // const name = require('...')
  src = src.replace(
    /^(const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*require\((['"])([^'"]+)\3\);?[ \t]*$/gm,
    (_m, _kw, name, _q, spec) => {
      const resolved = ensureJsExt(spec, filePath);
      const mod = coreMod(resolved);
      if (mod.endsWith('.json') || resolved.endsWith('.json')) {
        add(`import ${name} from '${resolved.endsWith('.json') ? resolved : mod}';`);
      } else {
        add(`import ${name} from '${mod}';`);
      }
      return '';
    }
  );

  // const { a, b: c } = require('...')
  src = src.replace(
    /^(const|let|var)\s+(\{[^}]+\})\s*=\s*require\((['"])([^'"]+)\3\);?[ \t]*$/gm,
    (_m, _kw, binding, _q, spec) => {
      const resolved = ensureJsExt(spec, filePath);
      const mod = coreMod(resolved);
      add(`import ${binding} from '${mod}';`);
      return '';
    }
  );

  // tools.fix(require('./data/foo.json')) mid-line → needs import
  let fixIdx = 0;
  src = src.replace(/tools\.fix\(require\((['"])(\.\/data\/[^'"]+\.json)\1\)\)/g, (_m, q, jsonPath) => {
    const id = `__json_${fixIdx++}`;
    add(`import ${id} from '${jsonPath}';`);
    return `tools.fix(${id})`;
  });

  // require('./data/foo.json') mid-expression (not tools.fix)
  src = src.replace(/(?<![\w.])require\((['"])(\.\/data\/[^'"]+\.json)\1\)/g, (_m, q, jsonPath) => {
    const id = `__json_${fixIdx++}`;
    add(`import ${id} from '${jsonPath}';`);
    return id;
  });

  // module.exports = ...
  if (/^module\.exports\s*=/m.test(src)) {
    src = src.replace(/^module\.exports\s*=\s*/m, 'export default ');
  }

  // exports.NAME =
  if (/^exports\./m.test(src)) {
    // rare - leave note
  }

  // __dirname / __filename
  if (/\b__dirname\b/.test(src) || /\b__filename\b/.test(src)) {
    add(`import path from 'node:path';`);
    add(`import {fileURLToPath} from 'node:url';`);
    src = `const __filename = fileURLToPath(import.meta.url);\nconst __dirname = path.dirname(__filename);\n${src}`;
  }

  if (imports.length) {
    src = `${imports.join('\n')}\n\n${src}`;
  }

  // cleanup blank lines
  src = src.replace(/\n{3,}/g, '\n\n');
  src = `// @vitest-migrated\n${src}`;

  // For data-only export files with no imports needed
  if (isDataExportOnly && !imports.length) {
    // already converted module.exports
  }

  fs.writeFileSync(filePath, src);
  return true;
}

const files = walk(specRoot);
let n = 0;
for (const f of files) {
  // skip already rewritten utils we hand-wrote without marker... re-run ok
  try {
    if (convertFile(f)) {
      n++;
      console.log('ok', path.relative(root, f));
    }
  } catch (e) {
    console.error('FAIL', path.relative(root, f), e.message);
  }
}
console.log(`Converted ${n}/${files.length}`);
