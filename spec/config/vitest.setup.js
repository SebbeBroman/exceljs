/**
 * Vitest setup: chai-compatible expect + mocha aliases + shared fixtures.
 */
import fs from 'node:fs';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {expect as viExpect, beforeAll, afterAll, beforeEach, afterEach} from 'vite-plus/test';
import {createChaiExpect} from './chai-expect.js';
import verquire from '../utils/verquire.js';
// Enable optional CSV API on Workbook (exceljs/csv entry)
import '../../lib/csv-entry.js';

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
