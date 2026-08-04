/* oxlint-disable no-console */
/**
 * E2E performance comparison: this package (builder) vs exceljs@4 (npm).
 *
 * Usage:
 *   pnpm build && pnpm bench
 *   pnpm build && pnpm bench -- --filter write
 *   pnpm build && pnpm bench -- --rows 2000
 *   node --expose-gc scripts/bench/e2e.mjs
 *
 * Contenders:
 *   - excel-ts (builder): public @sebbebroman/excel-ts API via dist/
 *   - exceljs@4:          npm exceljs (devDependency)
 *   - excel-ts (legacy):  internal Doc Workbook (same encoder as builder bridge)
 */
import {pathToFileURL} from 'node:url';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {bench, group, run, summary} from 'mitata';
import ExcelJS from 'exceljs';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '../..');

const {workbook, writeBuffer, load} = await import(pathToFileURL(path.join(root, 'dist/excel.js')).href);
const DocWorkbook = (await import(pathToFileURL(path.join(root, 'dist/lib/doc/workbook.js')).href)).default;

// --- CLI ---
const args = process.argv.slice(2);
function flag(name, fallback) {
  const i = args.indexOf(`--${name}`);
  if (i === -1) return fallback;
  const v = args[i + 1];
  if (v == null || v.startsWith('--')) return true;
  return v;
}

const ROWS = Number(flag('rows', 5000));
const COLS = Number(flag('cols', 8));
const FILTER = flag('filter', null); // 'write' | 'read' | 'roundtrip' | null
const INCLUDE_STYLES = flag('styles', false) === true || flag('styles', false) === 'true';
const SKIP_LEGACY = flag('no-legacy', false) === true || flag('no-legacy', false) === 'true';

// --- fixtures ---
function makeGrid(nRows, nCols) {
  const rows = new Array(nRows);
  for (let r = 0; r < nRows; r++) {
    const row = new Array(nCols);
    for (let c = 0; c < nCols; c++) {
      // mix types a bit: string id + numbers
      row[c] = c === 0 ? `r${r}` : r * nCols + c;
    }
    rows[r] = row;
  }
  return rows;
}

const grid = makeGrid(ROWS, COLS);
const gridSmall = makeGrid(Math.min(ROWS, 1000), COLS);

function bufLen(buf) {
  return buf?.byteLength ?? buf?.length ?? 0;
}

// Prebuild golden buffers (not timed) so read benches don't include write cost.
console.log(
  JSON.stringify({
    suite: 'e2e excel-ts vs exceljs',
    node: process.version,
    rows: ROWS,
    cols: COLS,
    styles: INCLUDE_STYLES,
    exceljs: '4.x (devDependency)',
    filter: FILTER,
  }),
);

const writeOpts = {useSharedStrings: true, useStyles: INCLUDE_STYLES};

async function writeBuilder(rows) {
  let b = workbook().sheet('data').rows(rows);
  if (INCLUDE_STYLES) b = b.style('1', {font: {bold: true}});
  return writeBuffer(b, writeOpts);
}

async function writeExceljs(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('data');
  if (INCLUDE_STYLES) {
    ws.getRow(1).font = {bold: true};
  }
  for (let i = 0; i < rows.length; i++) ws.addRow(rows[i]);
  return wb.xlsx.writeBuffer(writeOpts);
}

async function writeLegacyDoc(rows) {
  const wb = new DocWorkbook();
  const ws = wb.addWorksheet('data');
  if (INCLUDE_STYLES) {
    ws.getRow(1).font = {bold: true};
  }
  // batch when possible
  if (typeof ws.addRows === 'function') ws.addRows(rows);
  else for (let i = 0; i < rows.length; i++) ws.addRow(rows[i]);
  return wb.xlsx.writeBuffer(writeOpts);
}

async function readBuilder(bytes) {
  return load(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
}

async function readExceljs(bytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb;
}

async function readLegacyDoc(bytes) {
  const wb = new DocWorkbook();
  await wb.xlsx.load(bytes);
  return wb;
}

const [bufBuilder, bufExceljs, bufLegacy] = await Promise.all([
  writeBuilder(grid),
  writeExceljs(grid),
  SKIP_LEGACY ? Promise.resolve(null) : writeLegacyDoc(grid),
]);

console.log(
  JSON.stringify({
    fixtureBytes: {
      'excel-ts builder': bufLen(bufBuilder),
      'exceljs@4': bufLen(bufExceljs),
      ...(SKIP_LEGACY ? {} : {'excel-ts legacy Doc': bufLen(bufLegacy)}),
    },
  }),
);

function maybe(name, fn) {
  if (FILTER && !name.toLowerCase().includes(String(FILTER).toLowerCase())) return;
  fn();
}

// --- WRITE ---
maybe('write', () => {
  group(`write dense ${ROWS}×${COLS} → buffer`, () => {
    summary(() => {
      bench('excel-ts builder .rows().writeBuffer()', async () => {
        const buf = await writeBuilder(grid);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');

      bench('exceljs@4 addRow loop + writeBuffer', async () => {
        const buf = await writeExceljs(grid);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');

      if (!SKIP_LEGACY) {
        bench('excel-ts legacy DocWorkbook addRows + writeBuffer', async () => {
          const buf = await writeLegacyDoc(grid);
          if (bufLen(buf) < 500) throw new Error('tiny');
        }).gc('inner');
      }
    });
  });

  group(`write cell-by-cell ${gridSmall.length}×${COLS} → buffer`, () => {
    summary(() => {
      bench('excel-ts builder .cell()', async () => {
        let b = workbook().sheet('data');
        for (let r = 0; r < gridSmall.length; r++) {
          for (let c = 0; c < COLS; c++) {
            // A1-style: col letter simplistic for c < 26
            const addr = `${String.fromCharCode(65 + c)}${r + 1}`;
            b = b.cell(addr, gridSmall[r][c]);
          }
        }
        const buf = await b.writeBuffer(writeOpts);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');

      bench('exceljs@4 getCell().value', async () => {
        const wb = new ExcelJS.Workbook();
        const ws = wb.addWorksheet('data');
        for (let r = 0; r < gridSmall.length; r++) {
          for (let c = 0; c < COLS; c++) {
            ws.getCell(r + 1, c + 1).value = gridSmall[r][c];
          }
        }
        const buf = await wb.xlsx.writeBuffer(writeOpts);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');
    });
  });
});

// --- READ / PARSE ---
maybe('read', () => {
  group(`read/parse ${ROWS}×${COLS} from buffer`, () => {
    summary(() => {
      bench('excel-ts load() [from builder bytes]', async () => {
        const wb = await readBuilder(bufBuilder);
        if (!wb.sheets?.length) throw new Error('no sheets');
      }).gc('inner');

      bench('exceljs@4 xlsx.load() [from exceljs bytes]', async () => {
        const wb = await readExceljs(bufExceljs);
        if (!wb.worksheets?.length) throw new Error('no sheets');
      }).gc('inner');

      // Cross-read: parse the other package's output (format compatibility + cost)
      bench('excel-ts load() [from exceljs bytes]', async () => {
        const wb = await readBuilder(
          bufExceljs instanceof Uint8Array ? bufExceljs : new Uint8Array(bufExceljs),
        );
        if (!wb.sheets?.length) throw new Error('no sheets');
      }).gc('inner');

      bench('exceljs@4 xlsx.load() [from builder bytes]', async () => {
        const wb = await readExceljs(bufBuilder);
        if (!wb.worksheets?.length) throw new Error('no sheets');
      }).gc('inner');

      if (!SKIP_LEGACY) {
        bench('excel-ts legacy DocWorkbook xlsx.load()', async () => {
          const wb = await readLegacyDoc(bufLegacy);
          if (!wb.worksheets?.length) throw new Error('no sheets');
        }).gc('inner');
      }
    });
  });
});

// --- ROUND-TRIP ---
maybe('roundtrip', () => {
  group(`round-trip write→read ${ROWS}×${COLS}`, () => {
    summary(() => {
      bench('excel-ts builder writeBuffer → load', async () => {
        const buf = await writeBuilder(grid);
        const wb = await readBuilder(buf);
        const first = wb.sheets[0]?.rows?.[0]?.cells?.[1]?.value;
        if (first == null && first !== 0) throw new Error('empty');
      }).gc('inner');

      bench('exceljs@4 writeBuffer → load', async () => {
        const buf = await writeExceljs(grid);
        const wb = await readExceljs(buf);
        const v = wb.getWorksheet(1).getCell(1, 1).value;
        if (v == null && v !== 0) throw new Error('empty');
      }).gc('inner');

      if (!SKIP_LEGACY) {
        bench('excel-ts legacy Doc writeBuffer → load', async () => {
          const buf = await writeLegacyDoc(grid);
          const wb = await readLegacyDoc(buf);
          const v = wb.getWorksheet(1).getCell(1, 1).value;
          if (v == null && v !== 0) throw new Error('empty');
        }).gc('inner');
      }
    });
  });
});

await run({
  // quieter baseline; still prints mitata table
  colors: process.stdout.isTTY,
});
