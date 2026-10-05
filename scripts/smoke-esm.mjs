/**
 * Smoke test: native ESM import + builder write/load path.
 */
import {readFileSync, unlinkSync} from 'node:fs';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {workbook, writeBuffer, load, csv, ValueType} from '../dist/excel.js';
import {
  writeFile,
  readFile,
  readCsvFile,
  writeCsvFile,
  streamWrite,
  streamRead,
} from '../dist/node.js';
import {unzipSync, strFromU8} from 'fflate';

const assert = (cond, msg) => {
  if (!cond) throw new Error(msg || 'assertion failed');
};

assert(typeof workbook === 'function', 'named workbook export missing');
assert(typeof writeBuffer === 'function', 'named writeBuffer export missing');
assert(typeof load === 'function', 'named load export missing');
assert(csv && typeof csv.parse === 'function', 'named csv.parse export missing');
assert(typeof csv.stringify === 'function', 'named csv.stringify export missing');
assert(ValueType.String === 3, 'ValueType enum export missing');

const builder = workbook({creator: 'esm-smoke'})
  .sheet('Sheet1')
  .cell('A1', 'hello esm')
  .cell('B1', 42)
  .cell('C1', new Date('2020-01-15T00:00:00Z'));

const buf = await builder.writeBuffer();
assert(buf && buf.byteLength > 100, 'writeBuffer failed');

const files = unzipSync(buf);
const xml = Object.keys(files)
  .filter(n => n.endsWith('.xml'))
  .map(n => strFromU8(files[n]))
  .join('\n');
assert(xml.includes('hello esm'), 'A1 value missing from package');

// load round-trip + edit
const plain = await load(buf);
assert(plain.meta.creator === 'esm-smoke', 'load meta.creator mismatch');
assert(plain.sheets[0].name === 'Sheet1', 'load sheet name mismatch');
const a1 = plain.sheets[0].rows.find(r => r.number === 1).cells[1].value;
assert(a1 === 'hello esm', `load A1 mismatch: ${a1}`);

const edited = await workbook(plain).sheet('Sheet1').cell('A1', 'updated esm').writeBuffer();
const plain2 = await load(edited);
const a1b = plain2.sheets[0].rows.find(r => r.number === 1).cells[1].value;
assert(a1b === 'updated esm', `edit loop A1 mismatch: ${a1b}`);

const out = join(tmpdir(), `exceljs-esm-smoke-${Date.now()}.xlsx`);
await writeFile(
  out,
  workbook()
    .sheet('S')
    .rows([
      ['a', 1],
      ['b', 2],
    ]),
);
const disk = readFileSync(out);
assert(disk.byteLength > 100, 'writeFile failed');

const fromDisk = await readFile(out);
assert(fromDisk.sheets[0].rows[0].cells[1].value === 'a', 'readFile A1 mismatch');
unlinkSync(out);

// free writeBuffer
const buf2 = await writeBuffer(workbook().sheet('T').row([1, 2, 3]));
assert(buf2.byteLength > 100, 'free writeBuffer failed');

// CSV parse → sheet → writeBuffer; stringify / .csv()
const csvInit = await csv.parse('name,value\nalpha,1');
const fromCsv = await workbook().sheet('Csv', csvInit).writeBuffer();
assert(fromCsv.byteLength > 100, 'csv→xlsx failed');
const csvText = await workbook().sheet('Csv', csvInit).csv();
assert(csvText.includes('alpha'), `builder .csv failed: ${csvText}`);
const csvText2 = await csv.stringify(workbook().sheet('Csv', csvInit));
assert(csvText2.includes('name'), 'csv.stringify failed');

const csvPath = join(tmpdir(), `exceljs-esm-smoke-${Date.now()}.csv`);
await writeCsvFile(
  csvPath,
  workbook()
    .sheet('S')
    .rows([
      ['a', 1],
      ['b', 2],
    ]),
);
const csvDisk = await readCsvFile(csvPath);
assert(csvDisk.sheets[0].rows[0].cells[1].value === 'a', 'readCsvFile mismatch');
unlinkSync(csvPath);

// streamWrite / streamRead (node)
const streamPath = join(tmpdir(), `exceljs-esm-smoke-stream-${Date.now()}.xlsx`);
await streamWrite(streamPath, {
  sheets: [
    {
      name: 'Stream',
      rows: [
        ['id', 'name'],
        [1, 'alpha'],
        [2, 'beta'],
      ],
    },
  ],
});
const streamRows = [];
for await (const row of streamRead(streamPath)) {
  streamRows.push(row);
}
assert(streamRows.length >= 2, `streamRead expected rows, got ${streamRows.length}`);
assert(streamRows[0].sheetName === 'Stream', 'streamRead sheet name mismatch');
assert(
  streamRows.some(r => r.values[2] === 'alpha' || r.values[1] === 'alpha'),
  'streamRead values missing alpha',
);
unlinkSync(streamPath);

// Public surface: no accidental Doc Workbook / default ExcelJS export
const main = await import('../dist/excel.js');
assert(typeof main.default === 'undefined', 'main entry must not default-export ExcelJS');
assert(typeof main.Workbook !== 'function', 'main entry must not export Doc Workbook class');
const nodeMod = await import('../dist/node.js');
assert(typeof nodeMod.Workbook !== 'function', 'node entry must not export Doc Workbook class');
assert(typeof nodeMod.streamWrite === 'function', 'streamWrite missing on node entry');
assert(typeof nodeMod.streamRead === 'function', 'streamRead missing on node entry');

console.log('esm smoke ok');
console.log('  - named workbook / writeBuffer / load / csv');
console.log('  - builder writeBuffer + load edit loop');
console.log('  - node writeFile / readFile / readCsvFile / writeCsvFile');
console.log('  - node streamWrite / streamRead');
console.log('  - no Doc Workbook on public exports');
