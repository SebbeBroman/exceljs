/**
 * Smoke test: native ESM import + basic xlsx write/read round-trip.
 */
import {writeFileSync, readFileSync, unlinkSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import ExcelJS, {Workbook, ValueType} from '../excel.js';

const assert = (cond, msg) => {
  if (!cond) throw new Error(msg || 'assertion failed');
};

assert(typeof Workbook === 'function', 'named Workbook export missing');
assert(ExcelJS.Workbook === Workbook, 'default namespace Workbook mismatch');
assert(ValueType.String === 3, 'ValueType enum export missing');

const wb = new Workbook();
wb.creator = 'esm-smoke';
const ws = wb.addWorksheet('Sheet1');
ws.getCell('A1').value = 'hello esm';
ws.getCell('B1').value = 42;
ws.getCell('C1').value = new Date('2020-01-15T00:00:00Z');

const out = join(tmpdir(), `exceljs-esm-smoke-${Date.now()}.xlsx`);
await wb.xlsx.writeFile(out);

const wb2 = new Workbook();
await wb2.xlsx.readFile(out);
const ws2 = wb2.getWorksheet('Sheet1');
assert(ws2.getCell('A1').value === 'hello esm', 'A1 round-trip failed');
assert(ws2.getCell('B1').value === 42, 'B1 round-trip failed');

// CSV should not work until side-entry is loaded
let csvBlocked = false;
try {
  // Accessing the getter should throw until exceljs/csv is imported
  void wb.csv;
} catch (e) {
  csvBlocked = /CSV support is not loaded/.test(e.message);
}
assert(csvBlocked, 'csv should be gated until exceljs/csv is imported');

const {enableCsv} = await import('../lib/csv-entry.js');
enableCsv();
const csvPath = join(tmpdir(), `exceljs-esm-smoke-${Date.now()}.csv`);
await wb.csv.writeFile(csvPath);
const csvText = readFileSync(csvPath, 'utf8');
assert(csvText.includes('hello esm'), 'csv write failed');

unlinkSync(out);
unlinkSync(csvPath);

// Buffer write (common SvelteKit pattern: return file from +server)
const buf = await wb.xlsx.writeBuffer();
assert(buf && buf.byteLength > 100, 'writeBuffer failed');

console.log('esm smoke ok');
console.log('  - named Workbook import');
console.log('  - xlsx writeFile/readFile');
console.log('  - csv gated + enableCsv');
console.log('  - writeBuffer');
