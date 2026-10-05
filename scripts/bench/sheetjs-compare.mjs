/* oxlint-disable no-console */
/**
 * Size + perf: excel-ts vs SheetJS CE (xlsx) vs exceljs@4.
 *
 * Usage:
 *   pnpm add -D xlsx@0.18.5   # local only — do not commit to package.json
 *   pnpm build && node --expose-gc scripts/bench/sheetjs-compare.mjs
 *   node --expose-gc scripts/bench/sheetjs-compare.mjs -- --rows 5000
 */
import {pathToFileURL} from 'node:url';
import path from 'node:path';
import {fileURLToPath} from 'node:url';
import {readFileSync, existsSync, readdirSync, statSync} from 'node:fs';
import {gzipSync} from 'node:zlib';
import {createRequire} from 'node:module';
import {bench, group, run, summary} from 'mitata';
import ExcelJS from 'exceljs';

let XLSX;
try {
  XLSX = await import('xlsx');
} catch {
  console.error(
    'Missing optional dep `xlsx`. Install locally (do not add to package.json):\n' +
      '  pnpm add -D xlsx@0.18.5',
  );
  process.exit(1);
}

const require = createRequire(import.meta.url);
const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '../..');

const {workbook, writeBuffer, load, viewWorkbook, readRows} = await import(
  pathToFileURL(path.join(root, 'dist/excel.js')).href
);

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
const FILTER = flag('filter', null); // write | read | roundtrip | size | null

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

function bufLen(buf) {
  return buf?.byteLength ?? buf?.length ?? 0;
}

function dirSize(dir) {
  if (!existsSync(dir)) return 0;
  let total = 0;
  const walk = d => {
    for (const name of readdirSync(d)) {
      const p = path.join(d, name);
      const st = statSync(p);
      if (st.isDirectory()) walk(p);
      else total += st.size;
    }
  };
  walk(dir);
  return total;
}

function pkgMeta(name) {
  const pkgJsonPath = require.resolve(`${name}/package.json`);
  const pkg = JSON.parse(readFileSync(pkgJsonPath, 'utf8'));
  const pkgDir = path.dirname(pkgJsonPath);
  // Prefer browser/min entry for size comparison when present
  const candidates = [
    pkg.browser && typeof pkg.browser === 'string' ? path.join(pkgDir, pkg.browser) : null,
    pkg.unpkg ? path.join(pkgDir, pkg.unpkg) : null,
    pkg.jsdelivr ? path.join(pkgDir, pkg.jsdelivr) : null,
    pkg.module ? path.join(pkgDir, pkg.module) : null,
    pkg.main ? path.join(pkgDir, pkg.main) : null,
  ].filter(Boolean);

  // Known SheetJS / exceljs browser builds
  if (name === 'xlsx') {
    candidates.unshift(
      path.join(pkgDir, 'dist/xlsx.full.min.js'),
      path.join(pkgDir, 'xlsx.mjs'),
      path.join(pkgDir, 'xlsx.js'),
    );
  }
  if (name === 'exceljs') {
    candidates.unshift(
      path.join(pkgDir, 'dist/exceljs.min.js'),
      path.join(pkgDir, 'dist/exceljs.js'),
    );
  }

  let entry = null;
  for (const c of candidates) {
    if (existsSync(c) && statSync(c).isFile()) {
      entry = c;
      break;
    }
  }

  let entryRaw = 0;
  let entryGz = 0;
  if (entry) {
    const raw = readFileSync(entry);
    entryRaw = raw.byteLength;
    entryGz = gzipSync(raw).byteLength;
  }

  return {
    name,
    version: pkg.version,
    packageDirBytes: dirSize(pkgDir),
    entry: entry ? path.relative(pkgDir, entry) : null,
    entryRaw,
    entryGzip: entryGz,
  };
}

function formatBytes(n) {
  if (n < 1024) return `${n} B`;
  if (n < 1024 * 1024) return `${(n / 1024).toFixed(1)} KB`;
  return `${(n / (1024 * 1024)).toFixed(2)} MB`;
}

// --- size report ---
if (!FILTER || FILTER === 'size') {
  const excelTsDist = path.join(root, 'dist/excel.js');
  const excelTsRaw = readFileSync(excelTsDist);
  const excelTsEntry = {
    name: '@sebbebroman/exceljs (dist/excel.js)',
    version: JSON.parse(readFileSync(path.join(root, 'package.json'), 'utf8')).version,
    packageDirBytes: dirSize(path.join(root, 'dist')),
    entry: 'dist/excel.js',
    entryRaw: excelTsRaw.byteLength,
    entryGzip: gzipSync(excelTsRaw).byteLength,
  };

  const sizes = [excelTsEntry, pkgMeta('exceljs'), pkgMeta('xlsx')];
  console.log('\n=== Package / entry sizes (on disk) ===\n');
  console.log(
    'name'.padEnd(42),
    'ver'.padEnd(12),
    'pkg dir'.padStart(10),
    'entry raw'.padStart(12),
    'entry gzip'.padStart(12),
    'entry',
  );
  for (const s of sizes) {
    console.log(
      s.name.slice(0, 42).padEnd(42),
      String(s.version).padEnd(12),
      formatBytes(s.packageDirBytes).padStart(10),
      formatBytes(s.entryRaw).padStart(12),
      formatBytes(s.entryGzip).padStart(12),
      s.entry ?? '—',
    );
  }
  console.log(
    '\nNote: excel-ts dist/ is unminified ESM (tsc). exceljs/xlsx entries above are published builds (often minified). For apples-to-apples browser minify+gzip, see pnpm bench:browser (excel-ts vs exceljs).',
  );
}

// --- write helpers ---
async function writeExcelTs(rows) {
  return writeBuffer(workbook().sheet('data').rows(rows), {
    useSharedStrings: true,
    useStyles: false,
  });
}

async function writeExceljs(rows) {
  const wb = new ExcelJS.Workbook();
  const ws = wb.addWorksheet('data');
  for (let i = 0; i < rows.length; i++) ws.addRow(rows[i]);
  return wb.xlsx.writeBuffer({useSharedStrings: true, useStyles: false});
}

function writeSheetJS(rows) {
  const wb = XLSX.utils.book_new();
  const ws = XLSX.utils.aoa_to_sheet(rows);
  XLSX.utils.book_append_sheet(wb, ws, 'data');
  // array type → Uint8Array-ish buffer
  return XLSX.write(wb, {bookType: 'xlsx', type: 'array', compression: true});
}

function readSheetJS(bytes) {
  const wb = XLSX.read(bytes, {type: 'array', cellDates: false});
  const name = wb.SheetNames[0];
  return XLSX.utils.sheet_to_json(wb.Sheets[name], {header: 1});
}

async function readExcelTs(bytes) {
  return load(bytes instanceof Uint8Array ? bytes : new Uint8Array(bytes));
}

async function readExceljs(bytes) {
  const wb = new ExcelJS.Workbook();
  await wb.xlsx.load(bytes);
  return wb;
}

console.log(
  JSON.stringify({
    suite: 'excel-ts vs SheetJS CE vs exceljs@4',
    node: process.version,
    rows: ROWS,
    cols: COLS,
    filter: FILTER,
  }),
);

const [bufExcelTs, bufExceljs, bufSheetJS] = await Promise.all([
  writeExcelTs(grid),
  writeExceljs(grid),
  Promise.resolve(writeSheetJS(grid)),
]);

console.log(
  JSON.stringify({
    outputBytes: {
      'excel-ts': bufLen(bufExcelTs),
      'exceljs@4': bufLen(bufExceljs),
      'SheetJS CE (xlsx)': bufLen(bufSheetJS),
    },
  }),
);

function maybe(name, fn) {
  if (FILTER && FILTER !== 'size' && !name.toLowerCase().includes(String(FILTER).toLowerCase())) {
    return;
  }
  if (FILTER === 'size') return;
  fn();
}

maybe('write', () => {
  group(`write dense ${ROWS}×${COLS} → xlsx buffer`, () => {
    summary(() => {
      bench('excel-ts builder .rows().writeBuffer()', async () => {
        const buf = await writeExcelTs(grid);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');

      bench('exceljs@4 addRow + writeBuffer', async () => {
        const buf = await writeExceljs(grid);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');

      bench('SheetJS CE aoa_to_sheet + write(array)', () => {
        const buf = writeSheetJS(grid);
        if (bufLen(buf) < 500) throw new Error('tiny');
      }).gc('inner');
    });
  });
});

maybe('read', () => {
  group(`read/parse ${ROWS}×${COLS} from buffer`, () => {
    summary(() => {
      bench('excel-ts load() [full fidelity]', async () => {
        const wb = await readExcelTs(bufExcelTs);
        if (!wb.sheets?.length) throw new Error('no sheets');
      }).gc('inner');

      bench('excel-ts viewWorkbook + rows() [values]', async () => {
        const view = await viewWorkbook(bufExcelTs);
        const rows = view.sheet(0).rows({values: 'string'});
        if (!rows?.length) throw new Error('no rows');
      }).gc('inner');

      bench('excel-ts readRows() [first sheet strings]', async () => {
        const rows = await readRows(bufExcelTs);
        if (!rows?.length) throw new Error('no rows');
      }).gc('inner');

      bench('exceljs@4 xlsx.load()', async () => {
        const wb = await readExceljs(bufExceljs);
        if (!wb.worksheets?.length) throw new Error('no sheets');
      }).gc('inner');

      bench('SheetJS CE read + sheet_to_json', () => {
        const rows = readSheetJS(bufSheetJS);
        if (!rows?.length) throw new Error('no rows');
      }).gc('inner');

      // Cross-format: parse excel-ts output with SheetJS
      bench('SheetJS CE read excel-ts bytes', () => {
        const rows = readSheetJS(bufExcelTs);
        if (!rows?.length) throw new Error('no rows');
      }).gc('inner');

      bench('excel-ts load() SheetJS bytes', async () => {
        const wb = await readExcelTs(bufSheetJS);
        if (!wb.sheets?.length) throw new Error('no sheets');
      }).gc('inner');
    });
  });
});

maybe('roundtrip', () => {
  group(`round-trip write→read ${ROWS}×${COLS}`, () => {
    summary(() => {
      bench('excel-ts write → load', async () => {
        const buf = await writeExcelTs(grid);
        const wb = await readExcelTs(buf);
        if (!wb.sheets?.length) throw new Error('empty');
      }).gc('inner');

      bench('exceljs@4 write → load', async () => {
        const buf = await writeExceljs(grid);
        const wb = await readExceljs(buf);
        if (!wb.worksheets?.length) throw new Error('empty');
      }).gc('inner');

      bench('SheetJS CE write → read', () => {
        const buf = writeSheetJS(grid);
        const rows = readSheetJS(buf);
        if (!rows?.length) throw new Error('empty');
      }).gc('inner');
    });
  });
});

if (FILTER !== 'size') {
  await run({colors: process.stdout.isTTY});
}
