/**
 * Vitest setup: chai-compatible expect + mocha aliases + shared fixtures.
 */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {expect as viExpect, beforeAll, afterAll, beforeEach, afterEach} from 'vite-plus/test';
import {createChaiExpect} from './chai-expect.ts';
import verquire from '../utils/verquire.ts';
// Enable optional CSV API on Workbook (exceljs/csv entry)
import '../../lib/csv-entry.ts';
// Preload lazy doc features so sync APIs (addTable/addImage/…) work under Vitest.
// createRequire cannot load .ts sources; dynamic import via Vite can.
import {ensureDocFeatures} from '../../lib/doc/doc-features.ts';

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), '../..');

// Ensure integration tests can write under spec/out
fs.mkdirSync(path.join(root, 'spec/out'), {recursive: true});

// Chai-style expect used throughout the suite
const expect = createChaiExpect(viExpect);
globalThis.expect = expect;

// Historical helper: load lib modules by path (sync)
globalThis.verquire = verquire;

// Mocha aliases
globalThis.context = globalThis.describe;
globalThis.before = beforeAll;
globalThis.after = afterAll;
// beforeEach/afterEach already global via vitest globals; ensure present
globalThis.beforeEach = beforeEach;
globalThis.afterEach = afterEach;

// Top-level await: Vitest setup files may be async ESM
await ensureDocFeatures();
