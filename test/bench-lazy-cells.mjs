/* oxlint-disable no-console */
/**
 * Focused write/fill microbench for lazy/flat cell storage experiments.
 * Usage: node --expose-gc test/bench-lazy-cells.mjs [count]
 */
import {createRequire} from 'node:module';
import {pathToFileURL} from 'node:url';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const root = path.resolve(__dirname, '..');

// Prefer built dist entries.
const excelUrl = pathToFileURL(path.join(root, 'dist/excel.js')).href;
const streamUrl = pathToFileURL(path.join(root, 'dist/lib/stream-xlsx-entry.js')).href;
const {Workbook} = await import(excelUrl);
const {WorkbookWriter} = await import(streamUrl);

const count = parseInt(process.argv[2] || '8000', 10);
const passes = 3;

function randomName(length = 5) {
  const possible = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  let t = '';
  for (let i = 0; i < length; i++) t += possible.charAt(Math.floor(Math.random() * possible.length));
  return t;
}
function randomNum(d) {
  return Math.round(Math.random() * d);
}

function hrNow() {
  const t = process.hrtime.bigint();
  return t;
}
function hrMs(start) {
  return Number(process.hrtime.bigint() - start) / 1e6;
}

async function run(label, factory, opts) {
  // warmup
  await factory(Math.min(count, 400), opts);
  const times = [];
  let heapAfter = 0;
  for (let pass = 0; pass < passes; pass++) {
    if (global.gc) global.gc();
    const start = hrNow();
    await factory(count, opts);
    times.push(hrMs(start));
    heapAfter = process.memoryUsage().heapUsed / 1024 / 1024;
  }
  times.sort((a, b) => a - b);
  const result = {
    label,
    count,
    medianMs: Math.round(times[1] * 100) / 100,
    timesMs: times.map(t => Math.round(t * 100) / 100),
    heapMB: Math.round(heapAfter * 100) / 100,
    opts,
  };
  console.log(JSON.stringify(result));
  return result;
}

function makeRow(i) {
  return {
    key: i,
    name: randomName(5),
    age: randomNum(100),
    addr1: randomName(16),
    addr2: randomName(10),
    num1: randomNum(10000),
    num2: randomNum(100000),
    num3: randomNum(1000000),
  };
}

const cols = [
  {header: 'Col 1', key: 'key', width: 25},
  {header: 'Col 2', key: 'name', width: 32},
  {header: 'Col 3', key: 'age', width: 21},
  {header: 'Col 4', key: 'addr1', width: 18},
  {header: 'Col 5', key: 'addr2', width: 8},
  {header: 'Col 6', key: 'num1', width: 8},
  {header: 'Col 7', key: 'num2', width: 8},
  {header: 'Col 8', key: 'num3', width: 32},
];

async function fillOnly(n) {
  const wb = new Workbook();
  const ws = wb.addWorksheet('data');
  ws.columns = cols.map(c => ({key: c.key}));
  for (let i = 0; i < n; i++) ws.addRow(makeRow(i));
  const m = wb.model;
  void m.worksheets[0].rows.length;
}

async function docWrite(n, opts) {
  const wb = new Workbook();
  const ws = wb.addWorksheet('data');
  ws.columns = cols.map(c =>
    c.key === 'num3' && opts.useStyles
      ? {...c, style: {font: {name: 'Comic Sans MS', size: 8, bold: true}}}
      : c,
  );
  for (let i = 0; i < n; i++) ws.addRow(makeRow(i));
  await wb.xlsx.writeFile('/tmp/exceljs-bench-lazy.xlsx', opts);
}

async function streamWrite(n, opts) {
  const wb = new WorkbookWriter({filename: '/tmp/exceljs-bench-lazy-s.xlsx', ...opts});
  const ws = wb.addWorksheet('data');
  ws.columns = cols.map(c =>
    c.key === 'num3' && opts.useStyles
      ? {...c, style: {font: {name: 'Comic Sans MS', size: 8, bold: true}}}
      : c,
  );
  for (let i = 0; i < n; i++) ws.addRow(makeRow(i)).commit();
  await wb.commit();
}

console.log(JSON.stringify({phase: 'start', count, node: process.version}));
await run('fill+model only', fillOnly, {});
await run('doc plain own', docWrite, {useStyles: false, useSharedStrings: false});
await run('doc styled shared', docWrite, {useStyles: true, useSharedStrings: true});
await run('stream plain own', streamWrite, {useStyles: false, useSharedStrings: false});
await run('stream styled shared', streamWrite, {useStyles: true, useSharedStrings: true});
console.log(JSON.stringify({phase: 'done'}));
