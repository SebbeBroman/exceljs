/**
 * ESM replacement for the old CJS verquire helper.
 * Prefer direct imports; this remains for gradual migration and dynamic paths.
 */
import {createRequire} from 'node:module';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import ExcelJS from '../../lib/exceljs.nodejs.js';
import {enableCsv} from '../../lib/csv-entry.js';

// Ensure CSV API is available on the same Workbook class tests use
enableCsv(ExcelJS.Workbook);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const libRoot = path.resolve(__dirname, '../../lib');
const require = createRequire(import.meta.url);

const cache = Object.create(null);
cache.exceljs = ExcelJS;

function resolveTarget(modPath) {
  if (modPath === 'exceljs') {
    return path.join(libRoot, 'exceljs.nodejs.js');
  }
  const base = path.join(libRoot, modPath);
  if (base.endsWith('.js')) return base;
  return `${base}.js`;
}

/**
 * Synchronously load a lib module.
 * Unwraps default export to match historical CJS module.exports behavior.
 */
export default function verquire(modPath) {
  if (cache[modPath]) return cache[modPath];
  if (modPath === 'exceljs') return ExcelJS;

  const target = resolveTarget(modPath);
  let mod;
  try {
    mod = require(target);
  } catch (err) {
    throw new Error(`verquire failed to load ${modPath} from ${target}: ${err.message}`);
  }
  const value = mod && mod.__esModule && 'default' in mod ? mod.default : mod;
  cache[modPath] = value;
  return value;
}

export {verquire};
