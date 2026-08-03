#!/usr/bin/env node
/**
 * One-shot CJS → ESM converter for exceljs lib/.
 * Targeted at this codebase's require/module.exports patterns (not a general codemod).
 */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');
const libRoot = path.join(root, 'lib');

function walk(dir, files = []) {
  for (const entry of fs.readdirSync(dir, {withFileTypes: true})) {
    const full = path.join(dir, entry.name);
    if (entry.isDirectory()) walk(full, files);
    else if (entry.name.endsWith('.js')) files.push(full);
  }
  return files;
}

function resolveSpec(spec, fromFile) {
  if (!spec.startsWith('.') && !spec.startsWith('/')) return spec; // package
  // Ensure relative ESM imports keep explicit .js (Node + Vite both happy)
  const dir = path.dirname(fromFile);
  const abs = path.resolve(dir, spec);
  let resolved = abs;
  if (fs.existsSync(abs) && fs.statSync(abs).isDirectory()) {
    resolved = path.join(abs, 'index.js');
  } else if (!abs.endsWith('.js') && fs.existsSync(`${abs}.js`)) {
    resolved = `${abs}.js`;
  } else if (!abs.endsWith('.js') && fs.existsSync(abs)) {
    resolved = abs;
  } else if (!spec.endsWith('.js') && !spec.endsWith('.json')) {
    // default: add .js for relative JS modules
    return spec.endsWith('/') ? `${spec}index.js` : `${spec}.js`;
  }
  let rel = path.relative(dir, resolved).replace(/\\/g, '/');
  if (!rel.startsWith('.')) rel = `./${rel}`;
  return rel;
}

function convertFile(filePath, source) {
  let code = source.replace(/\r\n/g, '\n');
  const imports = [];
  let n = 0;
  const freshId = () => `__esm_${n++}`;

  const pushDefault = (spec, name) => {
    imports.push(`import ${name} from '${resolveSpec(spec, filePath)}';`);
  };
  const pushNamed = (spec, namedClause) => {
    imports.push(`import { ${namedClause} } from '${resolveSpec(spec, filePath)}';`);
  };
  const pushSideEffect = spec => {
    imports.push(`import '${resolveSpec(spec, filePath)}';`);
  };
  const pushNamespace = (spec, name) => {
    imports.push(`import * as ${name} from '${resolveSpec(spec, filePath)}';`);
  };

  // 1) Side-effect requires on their own line
  code = code.replace(/^require\((['"])([^'"]+)\1\);?\s*$/gm, (_, _q, spec) => {
    pushSideEffect(spec);
    return '';
  });

  // 2) const x = require('pkg').chain...
  code = code.replace(
    /^(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*require\((['"])([^'"]+)\2\)((?:\.\w+(?:\([^;]*\))?)+);?\s*$/gm,
    (_, name, _q, spec, chain) => {
      const id = freshId();
      pushDefault(spec, id);
      return `const ${name} = ${id}${chain};`;
    }
  );

  // 3) const { nested } = require(...)  — including nested destructuring
  code = code.replace(
    /^(?:const|let|var)\s+(\{[\s\S]*?\})\s*=\s*require\((['"])([^'"]+)\2\);?\s*$/gm,
    (match, destr, _q, spec) => {
      // Flatten only shallow { a, b: c } into import { a, b as c }
      const inner = destr.trim().slice(1, -1).trim();
      if (!inner.includes('{') && !inner.includes('}')) {
        const parts = inner
          .split(',')
          .map(s => s.trim())
          .filter(Boolean)
          .map(part => {
            if (part.includes(':')) {
              const [a, b] = part.split(':').map(s => s.trim());
              return `${a} as ${b}`;
            }
            return part;
          });
        pushNamed(spec, parts.join(', '));
        return '';
      }
      // Nested: import default/namespace then destructure
      const id = freshId();
      pushDefault(spec, id);
      return `const ${destr} = ${id};`;
    }
  );

  // 4) const x = require(...)
  code = code.replace(
    /^(?:const|let|var)\s+([A-Za-z_$][\w$]*)\s*=\s*require\((['"])([^'"]+)\2\);?\s*$/gm,
    (_, name, _q, spec) => {
      pushDefault(spec, name);
      return '';
    }
  );

  // 5) Any remaining require(...) expressions (assignments, Object.assign, object literals)
  code = code.replace(/require\((['"])([^'"]+)\1\)/g, (_m, _q, spec) => {
    const id = freshId();
    pushDefault(spec, id);
    return id;
  });

  // 6) exports.foo = bar  → collect named exports
  const exportNames = [];
  code = code.replace(/^exports\.([A-Za-z_$][\w$]*)\s*=\s*([A-Za-z_$][\w$]*);?\s*$/gm, (_, exp, local) => {
    if (exp === local) exportNames.push(exp);
    else exportNames.push(`${local} as ${exp}`);
    return '';
  });

  // 7) module.exports = ...
  // Prefer: export default X; export { X } when X is a simple identifier
  if (/module\.exports\s*=/.test(code)) {
    // module.exports = { name } shorthand-only (single level)
    code = code.replace(/module\.exports\s*=\s*\{([^}]*)\};?\s*$/m, (match, body) => {
      const names = body
        .split(',')
        .map(s => s.trim())
        .filter(Boolean);
      if (names.length && names.every(n => /^[A-Za-z_$][\w$]*$/.test(n))) {
        exportNames.push(...names);
        return `export default { ${names.join(', ')} };`;
      }
      // large object literal (enums, defaultnumformats, rel-type, sp-pr)
      return `export default {${body}};`;
    });

    // module.exports = async function* name? or anonymous
    code = code.replace(
      /module\.exports\s*=\s*(async\s+function\s*\*)/,
      'export default $1'
    );

    // module.exports = Identifier or expression
    code = code.replace(/module\.exports\s*=\s*([^;]+);?\s*$/m, (match, expr) => {
      expr = expr.trim();
      if (/^[A-Za-z_$][\w$]*$/.test(expr)) {
        exportNames.push(expr);
        return `export default ${expr};`;
      }
      return `export default ${expr};`;
    });

    // leftover module.exports = (multiline objects that didn't match $)
    code = code.replace(/module\.exports\s*=\s*/, 'export default ');
  }

  if (exportNames.length) {
    // unique
    const unique = [...new Set(exportNames)];
    code = `${code.trimEnd()}\nexport { ${unique.join(', ')} };\n`;
  }

  // Clean excessive blank lines at top after removing requires
  code = code.replace(/^\s*\n/, '');
  const importBlock = imports.length ? `${imports.join('\n')}\n\n` : '';
  code = importBlock + code.replace(/^\uFEFF?/, '');

  // Drop redundant 'use strict' in ESM
  code = code.replace(/^['"]use strict['"];\s*\n/gm, '');

  if (/\brequire\s*\(/.test(code) || /\bmodule\.exports\b/.test(code) || /\bexports\./.test(code)) {
    throw new Error(`Unconverted CJS remnants in ${path.relative(root, filePath)}`);
  }

  return code;
}

const files = walk(libRoot);
let ok = 0;
const errors = [];

for (const file of files) {
  const src = fs.readFileSync(file, 'utf8');
  // Skip if already looks like ESM-only and has no require
  if (!src.includes('require(') && !src.includes('module.exports') && !src.includes('exports.')) {
    continue;
  }
  try {
    const out = convertFile(file, src);
    fs.writeFileSync(file, out);
    ok++;
  } catch (err) {
    errors.push(String(err.message || err));
  }
}

console.log(`Converted ${ok} files`);
if (errors.length) {
  console.error('Errors:');
  for (const e of errors) console.error(' -', e);
  process.exit(1);
}
