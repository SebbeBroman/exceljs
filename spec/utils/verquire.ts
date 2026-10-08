/**
 * ESM replacement for the old CJS verquire helper.
 * Prefer direct imports; this remains for gradual migration and dynamic paths.
 *
 * Loads TypeScript sources under lib/ via Vite's import.meta.glob (Vitest).
 */
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const libRoot = path.resolve(__dirname, '../../lib');

// Eager map of lib modules — Vitest/Vite resolves .ts from these keys
const modules = import.meta.glob('../../lib/**/*.{ts,js}', {eager: true});

// eslint-disable-next-line @typescript-eslint/no-explicit-any
const cache: Record<string, any> = Object.create(null);

function normalizeKey(modPath: string): string {
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
// eslint-disable-next-line @typescript-eslint/no-explicit-any
export default function verquire(modPath: string): any {
  if (cache[modPath]) return cache[modPath];

  const key = normalizeKey(modPath);
  const mod = modules[key] as {default?: unknown} | undefined;
  if (!mod) {
    const available = Object.keys(modules)
      .filter(k => k.includes(modPath.replace(/\.js$/, '')))
      .slice(0, 5);
    throw new Error(
      `verquire failed to load ${modPath} (key=${key}). Similar: ${available.join(', ') || 'none'}`,
    );
  }
  const value = 'default' in mod ? mod.default : mod;
  cache[modPath] = value;
  return value;
}

export {verquire, libRoot};
