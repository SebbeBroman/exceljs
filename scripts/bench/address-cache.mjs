import {performance} from 'node:perf_hooks';
import {resolve} from 'node:path';
import {pathToFileURL} from 'node:url';
import {readFileSync, writeFileSync} from 'node:fs';
const {default: col} = await import(
  pathToFileURL(resolve(process.argv[2] || 'dist', 'lib/utils/col-cache.js'))
);
if (!global.gc) throw new Error('Run with --expose-gc');
global.gc();
const heap = process.memoryUsage().heapUsed;
const start = performance.now();
const value = col.n2l(16384);
const coldMs = performance.now() - start;
global.gc();
const retainedBytes = process.memoryUsage().heapUsed - heap;
let checksum = 0;
const warm = performance.now();
for (let i = 0; i < 100000; i++) checksum += col.n2l((i % 16384) + 1).length;
const warmMs = performance.now() - warm;
global.gc();
const addressHeap = process.memoryUsage().heapUsed;
for (let row = 1; row <= 100; row++) {
  for (let column = 1; column <= 100; column++) col.decodeAddress(col.encodeAddress(row, column));
}
global.gc();
const result = {
  value,
  coldMs,
  retainedBytes,
  warmMs,
  checksum,
  addressRetainedBytes: process.memoryUsage().heapUsed - addressHeap,
};
console.log(JSON.stringify(result));
if (process.argv[3]) {
  const file = resolve(process.argv[3]);
  const json = JSON.parse(readFileSync(file, 'utf8'));
  json.addresses = result;
  writeFileSync(file, JSON.stringify(json, null, 2) + '\n');
}
