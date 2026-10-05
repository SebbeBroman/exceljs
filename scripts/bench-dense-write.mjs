/* oxlint-disable no-console */
/**
 * Phase 3 dense-write microbench.
 *
 * Compares:
 *   A) builder dense path: workbook().sheet().rows(N×K).writeBuffer()
 *   B) DocWorkbook addRow loop + xlsx.writeBuffer (legacy-style)
 *   C) builder materialize only (no zip) via internal compile path if available
 *
 * Usage:
 *   pnpm build && node scripts/bench-dense-write.mjs [rows]
 *
 * Default rows: 5000, 8 cols, 3 timed passes after warmup.
 */
import {pathToFileURL} from 'node:url';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');

const excelUrl = pathToFileURL(path.join(root, 'dist/excel.js')).href;
const docWbUrl = pathToFileURL(path.join(root, 'dist/lib/doc/workbook.js')).href;

const {workbook} = await import(excelUrl);
const DocWorkbook = (await import(docWbUrl)).default;

const nRows = parseInt(process.argv[2] || '5000', 10);
const nCols = 8;
const passes = 3;

function makeGrid(n) {
  const rows = new Array(n);
  for (let i = 0; i < n; i++) {
    const row = new Array(nCols);
    for (let c = 0; c < nCols; c++) {
      row[c] = c === 0 ? `r${i}` : i * nCols + c;
    }
    rows[i] = row;
  }
  return rows;
}

function hrMs(start) {
  return Number(process.hrtime.bigint() - start) / 1e6;
}

async function time(label, fn) {
  await fn(); // warmup
  const times = [];
  let heapMB = 0;
  for (let p = 0; p < passes; p++) {
    if (global.gc) global.gc();
    const t0 = process.hrtime.bigint();
    await fn();
    times.push(hrMs(t0));
    heapMB = process.memoryUsage().heapUsed / 1024 / 1024;
  }
  times.sort((a, b) => a - b);
  const median = Math.round(times[Math.floor(times.length / 2)] * 100) / 100;
  const result = {
    label,
    nRows,
    nCols,
    medianMs: median,
    timesMs: times.map(t => Math.round(t * 100) / 100),
    heapMB: Math.round(heapMB * 100) / 100,
  };
  console.log(JSON.stringify(result));
  return result;
}

const grid = makeGrid(nRows);
// Cell-by-cell baseline only for moderate N (very slow otherwise).
const cellByCellN = Math.min(nRows, 2000);
const cellGrid = nRows === cellByCellN ? grid : makeGrid(cellByCellN);

console.log(JSON.stringify({phase: 'start', nRows, nCols, node: process.version}));

const builder = await time('builder.rows dense writeBuffer', async () => {
  const buf = await workbook().sheet('data').rows(grid).writeBuffer({useSharedStrings: false});
  if (buf.byteLength < 1000) throw new Error('tiny buffer');
});

const docLoop = await time('DocWorkbook addRow loop writeBuffer', async () => {
  const wb = new DocWorkbook();
  const ws = wb.addWorksheet('data');
  for (let i = 0; i < grid.length; i++) {
    ws.addRow(grid[i]);
  }
  const buf = await wb.xlsx.writeBuffer({useSharedStrings: false, useStyles: false});
  if (!(buf && (buf.byteLength ?? buf.length) > 1000)) throw new Error('tiny buffer');
});

// Many consecutive .row() calls — optimizeOps fuses them into one .rows.
const fusedRows = await time('builder N×.row() fused writeBuffer', async () => {
  let b = workbook().sheet('data');
  for (let i = 0; i < grid.length; i++) {
    b = b.row(grid[i]);
  }
  const buf = await b.writeBuffer({useSharedStrings: false});
  if (buf.byteLength < 1000) throw new Error('tiny buffer');
});

let cellByCell = null;
if (cellByCellN <= 2500) {
  cellByCell = await time(`builder cell-by-cell writeBuffer (n=${cellByCellN})`, async () => {
    let b = workbook().sheet('data');
    for (let r = 0; r < cellGrid.length; r++) {
      for (let c = 0; c < nCols; c++) {
        const col = String.fromCharCode(65 + c);
        b = b.cell(`${col}${r + 1}`, cellGrid[r][c]);
      }
    }
    const buf = await b.writeBuffer({useSharedStrings: false});
    if (buf.byteLength < 500) throw new Error('tiny buffer');
  });
}

const speedupDoc =
  docLoop.medianMs > 0 ? Math.round((docLoop.medianMs / builder.medianMs) * 100) / 100 : null;
const speedupFused =
  fusedRows.medianMs > 0 ? Math.round((fusedRows.medianMs / builder.medianMs) * 100) / 100 : null;

console.log(
  JSON.stringify({
    phase: 'done',
    speedupVsDocLoop: speedupDoc,
    speedupDenseVsFusedRows: speedupFused,
    cellByCellMedianMs: cellByCell?.medianMs ?? null,
    note: 'speedup > 1 means left side of ratio is slower; dense .rows() is the target path',
  }),
);
