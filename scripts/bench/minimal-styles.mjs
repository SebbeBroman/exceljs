import {performance} from 'node:perf_hooks';
import {pathToFileURL} from 'node:url';
import {resolve} from 'node:path';
import {existsSync, readFileSync, writeFileSync} from 'node:fs';
const directory = resolve(process.argv[2] || 'dist');
const newPath = resolve(directory, 'lib/xlsx/xform/style/minimal-styles.js');
const Styles = existsSync(newPath)
  ? (await import(pathToFileURL(newPath))).default
  : (await import(pathToFileURL(resolve(directory, 'lib/xlsx/xform/style/styles-xform.js'))))
      .default.Mock;
let checksum = 0;
const samples = [];
for (let iteration = 0; iteration < 14; iteration++) {
  const start = performance.now();
  for (let i = 0; i < 1000; i++) {
    const styles = new Styles();
    checksum += (styles.toXml ? await styles.toXml() : styles.xml).length;
  }
  if (iteration >= 3) samples.push(performance.now() - start);
}
const result = {
  workload: '1000 minimal styles constructions and serializations',
  medianMs: samples.sort((a, b) => a - b)[5],
  checksum,
};
console.log(JSON.stringify(result));
if (process.argv[3]) {
  const path = resolve(process.argv[3]);
  const json = JSON.parse(readFileSync(path, 'utf8'));
  json.minimalStyles = result;
  writeFileSync(path, JSON.stringify(json, null, 2) + '\n');
}
