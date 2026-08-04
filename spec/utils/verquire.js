/**
 * ESM replacement for the old CJS verquire helper.
 * Prefer direct imports; this remains for gradual migration and dynamic paths.
 *
 * Loads TypeScript sources under lib/ via Vite's import.meta.glob (Vitest).
 */
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import ExcelJS from '../../lib/exceljs.nodejs.ts';
import {enableCsv} from '../../lib/csv-entry.ts';

// Ensure CSV API is available on the same Workbook class tests use
enableCsv(ExcelJS.Workbook);

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const libRoot = path.resolve(__dirname, '../../lib');

// Eager map of lib modules — Vitest/Vite resolves .ts from these keys
const modules = import.meta.glob('../../lib/**/*.{ts,js}', {eager: true});

const cache = Object.create(null);
cache.exceljs = ExcelJS;

function normalizeKey(modPath) {
  if (modPath === 'exceljs') {
    return '../../lib/exceljs.nodejs.ts';
  }
  const withoutExt = modPath.replace(/\.js$/, '');
  const tsKey = `../../lib/${withoutExt}.ts`;
  const jsKey = `../../lib/${withoutExt}.js`;
  if (modules[tsKey]) return tsKey;
  if (modules[jsKey]) return jsKey;
  return tsKey;
}

/**
 * Synchronously load a lib module.
 * Unwraps default export to match historical CJS module.exports behavior.
 */
export default function verquire(modPath) {
  if (cache[modPath]) return cache[modPath];
  if (modPath === 'exceljs') return ExcelJS;

  const key = normalizeKey(modPath);
  const mod = modules[key];
  if (!mod) {
    const available = Object.keys(modules)
      .filter(k => k.includes(modPath.replace(/\.js$/, '')))
      .slice(0, 5);
    throw new Error(
      `verquire failed to load ${modPath} (key=${key}). Similar: ${available.join(', ') || 'none'}`,
    );
  }
  const value = mod && 'default' in mod ? mod.default : mod;
  cache[modPath] = value;
  return value;
}

export {verquire, libRoot};
