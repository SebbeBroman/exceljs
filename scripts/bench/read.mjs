/* oxlint-disable no-console */
/**
 * Read-path benchmarks: viewWorkbook / readRows / load vs exceljs@4.
 *
 *   pnpm build && node --expose-gc scripts/bench/read.mjs
 *   node --expose-gc scripts/bench/read.mjs --rows 5000
 */
import {pathToFileURL} from 'node:url';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {bench, group, run, summary} from 'mitata';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '../..');

const {workbook, writeBuffer, load, viewWorkbook, readRows} = await import(
  pathToFileURL(path.join(root, 'dist/excel.js')).href
);

const args = process.argv.slice(2);
function flag(name, fallback) {
  const i = args.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const v = args[i + 1];
  if (v == null || v.startsWith('--')) return true;
  return v;
}

const ROWS = Number(flag('rows', 2000));
const COLS = Number(flag('cols', 8));

function makeGrid(nRows, nCols) {
  const rows = new Array(nRows);
  for (let r = 0; r < nRows; r++) {
    const row = new Array(nCols);
    for (let c = 0; c < nCols; c++) {
      row[c] = c === 0 ? `r${r}` : r * nCols + c;
    }
    rows[r] = row;
  }
  return rows;
}

const grid = makeGrid(ROWS, COLS);

const bufBuilder = await writeBuffer(workbook().sheet('data').rows(grid), {
  useSharedStrings: true,
  useStyles: false,
});

const exceljsWb = new ExcelJS.Workbook();
const ws = exceljsWb.addWorksheet('data');
for (const row of grid) ws.addRow(row);
const bufExceljs = await exceljsWb.xlsx.writeBuffer({
  useSharedStrings: true,
  useStyles: false,
});

const csvText = grid.map(r => r.join(',')).join('\n');
const csvBytes = new TextEncoder().encode(csvText);

console.log(
  JSON.stringify({
    suite: 'read: viewWorkbook / readRows / load / exceljs',
    rows: ROWS,
    cols: COLS,
    fixtureBytes: {
      excelTsXlsx: bufBuilder.byteLength,
      exceljsXlsx: bufExceljs.byteLength ?? bufExceljs.length,
      csv: csvBytes.byteLength,
    },
    node: process.version,
  }),
);

group(`xlsx ${ROWS}×${COLS}`, () => {
  summary(() => {
    bench('excel-ts load()', async () => {
      const wb = await load(bufBuilder);
      if (!wb.sheets?.length) throw new Error('empty');
    }).gc('inner');

    bench('excel-ts viewWorkbook + sheet(0).rows()', async () => {
      const view = await viewWorkbook(bufBuilder, {format: 'xlsx'});
      const rows = view.sheet(0).rows({values: 'string'});
      if (rows.length < 100) throw new Error('empty');
    }).gc('inner');

    bench('excel-ts viewWorkbook + rows({ end: 100 })', async () => {
      const view = await viewWorkbook(bufBuilder, {format: 'xlsx'});
      const rows = view.sheet(0).rows({start: 1, end: 100, values: 'string'});
      if (rows.length < 1) throw new Error('empty');
    }).gc('inner');

    bench('excel-ts readRows()', async () => {
      const rows = await readRows(bufBuilder, {format: 'xlsx'});
      if (rows.length < 100) throw new Error('empty');
    }).gc('inner');

    bench('exceljs@4 xlsx.load()', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(bufExceljs);
      if (!wb.worksheets?.length) throw new Error('empty');
    }).gc('inner');

    bench('exceljs@4 load + eachRow walk', async () => {
      const wb = new ExcelJS.Workbook();
      await wb.xlsx.load(bufExceljs);
      let n = 0;
      wb.worksheets[0].eachRow({includeEmpty: false}, row => {
        n += row.cellCount;
      });
      if (n < 100) throw new Error('empty');
    }).gc('inner');
  });
});

group(`csv ${ROWS}×${COLS}`, () => {
  summary(() => {
    bench('excel-ts viewWorkbook(csv) + rows()', async () => {
      const view = await viewWorkbook(csvBytes, {format: 'csv'});
      const rows = view.sheet(0).rows();
      if (rows.length < 100) throw new Error('empty');
    }).gc('inner');

    bench('excel-ts readRows(csv)', async () => {
      const rows = await readRows(csvBytes, {format: 'csv'});
      if (rows.length < 100) throw new Error('empty');
    }).gc('inner');
  });
});

await run({colors: process.stdout.isTTY});
