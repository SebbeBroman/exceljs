/* oxlint-disable no-console */
/** Experimental native compression. Does not change the production ZIP writer. */
import {build} from 'esbuild';
import {readFileSync, writeFileSync, mkdirSync, copyFileSync} from 'node:fs';
import {join, resolve, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createServer} from 'node:http';
import puppeteer from 'puppeteer-core';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/browser-native-zip');
mkdirSync(out, {recursive: true});
const entry = `
import {workbook,writeBuffer,load} from ${JSON.stringify(join(root, 'dist/excel.js'))};
import BufferZipWriter from ${JSON.stringify(join(root, 'dist/lib/utils/buffer-zip.js'))};
import {Zip,ZipPassThrough,unzipSync,zipSync} from 'fflate';
class NativeDeflate extends ZipPassThrough {
 constructor(name){super(name);this.compression=8;}
 process(chunk,final){
  if(!final)throw new Error('Prototype accepts complete buffered entries only');
  new Response(new Blob([chunk]).stream().pipeThrough(new CompressionStream('deflate-raw'))).arrayBuffer()
   .then(bytes=>this.ondata(null,new Uint8Array(bytes),true),err=>this.ondata(err,null,true));
 }
}
function nativeZip(files){return new Promise((resolve,reject)=>{
 const chunks=[];let length=0;
 const zip=new Zip((error,chunk,final)=>{if(error){reject(error);return;}chunks.push(chunk);length+=chunk.length;if(final){const bytes=new Uint8Array(length);let offset=0;for(const part of chunks){bytes.set(part,offset);offset+=part.length;}resolve(bytes);}});
 for(const [name,bytes] of Object.entries(files)){const stream=new NativeDeflate(name);zip.add(stream);stream.push(bytes,true);}zip.end();
});}
const original=BufferZipWriter.prototype.toBytes;
export async function write(rows,engine){
 const fixed=new Date('2020-01-01T00:00:00Z');
 const builder=workbook({created:fixed,modified:fixed}).sheet('data').rows(rows);
 BufferZipWriter.prototype.toBytes=engine==='native'?function(){return nativeZip(this.files);}:engine.startsWith('mem')?function(level){return zipSync(this.files,{level,mem:Number(engine.slice(3))});}:original;
 try{return await writeBuffer(builder,{useSharedStrings:true,useStyles:false});}finally{BufferZipWriter.prototype.toBytes=original;}
}
export function identicalParts(a,b){const left=unzipSync(a),right=unzipSync(b);const keys=Object.keys(left).sort();if(keys.join()!==Object.keys(right).sort().join())throw new Error('ZIP entries differ');for(const key of keys){const x=left[key],y=right[key];if(x.length!==y.length||x.some((value,i)=>value!==y[i]))throw new Error('Uncompressed part differs: '+key);}return keys.length;}
export {load};
`;
await build({
  bundle: true,
  minify: true,
  platform: 'browser',
  format: 'esm',
  stdin: {contents: entry, resolveDir: root},
  outfile: join(out, 'experiment.mjs'),
  alias: {
    fs: join(root, 'lib/shims/fs-browser.ts'),
    module: join(root, 'lib/shims/node-module-browser.ts'),
    'node:module': join(root, 'lib/shims/node-module-browser.ts'),
  },
  define: {global: 'globalThis'},
});
copyFileSync(
  join(root, 'build/browser-libraries/sheetjs-read-write.mjs'),
  join(out, 'sheetjs.mjs'),
);
copyFileSync(join(root, 'node_modules/exceljs/dist/exceljs.min.js'), join(out, 'exceljs.min.js'));
writeFileSync(join(out, 'index.html'), '<!doctype html><script src="./exceljs.min.js"></script>');
const server = createServer((req, res) => {
  const name = req.url === '/' ? 'index.html' : req.url.slice(1);
  if (!['index.html', 'experiment.mjs', 'sheetjs.mjs', 'exceljs.min.js'].includes(name)) {
    res.writeHead(404);
    res.end();
    return;
  }
  res.setHeader('Content-Type', name.endsWith('html') ? 'text/html' : 'text/javascript');
  res.end(readFileSync(join(out, name)));
});
await new Promise(r => server.listen(0, '127.0.0.1', r));
let browser;
try {
  browser = await puppeteer.launch({
    executablePath:
      process.env.CHROME_PATH || '/Applications/Google Chrome.app/Contents/MacOS/Google Chrome',
    headless: true,
  });
  const page = await browser.newPage();
  await page.goto(`http://127.0.0.1:${server.address().port}`);
  const results = await page.evaluate(async () => {
    try {
      new CompressionStream('deflate-raw');
    } catch (error) {
      return {supported: false, error: error.message};
    }
    const api = await import('./experiment.mjs'),
      XLSX = await import('./sheetjs.mjs');
    const results = {supported: true, cases: {}};
    for (const count of [2000, 20000]) {
      const rows = Array.from({length: count}, (_, r) =>
        Array.from({length: 8}, (_, c) => (c === 0 ? 'r' + r : r * 8 + c)),
      );
      const expected = JSON.stringify(rows);
      const baseline = await api.write(rows, 'fflate'),
        native = await api.write(rows, 'native');
      const parts = api.identicalParts(baseline, native);
      const plain = await api.load(native);
      const cells = plain.sheets[0].rows.map(r =>
        Array.from({length: 8}, (_, c) => r.cells[c + 1].value),
      );
      if (JSON.stringify(cells) !== expected) throw new Error('Local native cross-read');
      const swb = XLSX.read(native, {type: 'array', dense: true});
      if (
        JSON.stringify(
          XLSX.utils.sheet_to_json(swb.Sheets[swb.SheetNames[0]], {header: 1, raw: true}),
        ) !== expected
      )
        throw new Error('SheetJS native cross-read');
      const ewb = new window.ExcelJS.Workbook();
      await ewb.xlsx.load(native);
      if (
        JSON.stringify(
          ewb.worksheets[0]
            .getSheetValues()
            .slice(1)
            .map(r => r.slice(1)),
        ) !== expected
      )
        throw new Error('ExcelJS native cross-read');
      const engines = ['fflate', 'native', 'mem4', 'mem6', 'mem8'];
      const sizes = {fflate: baseline.length, native: native.length};
      for (const engine of engines.slice(2)) {
        const bytes = await api.write(rows, engine);
        api.identicalParts(baseline, bytes);
        sizes[engine] = bytes.length;
      }
      const samples = Object.fromEntries(engines.map(x => [x, []]));
      for (let i = 0; i < 9; i++)
        for (let k = 0; k < engines.length; k++) {
          const engine = engines[(i + k) % engines.length];
          const start = performance.now();
          await api.write(rows, engine);
          if (i >= 2) samples[engine].push(performance.now() - start);
        }
      results.cases[count] = {
        partsIdentical: parts,
        outputBytes: sizes,
        timings: Object.fromEntries(
          Object.entries(samples).map(([name, x]) => [
            name,
            {medianMs: [...x].sort((a, b) => a - b)[3], samplesMs: x},
          ]),
        ),
      };
    }
    const special = [
      ['a&<b', 'é😀', 'a"b', true, null],
      ['two\nlines', -2, 0, false, ''],
    ];
    const a = await api.write(special, 'fflate'),
      b = await api.write(special, 'native');
    results.specialParts = api.identicalParts(a, b);
    return results;
  });
  writeFileSync(
    join(root, 'build/browser-profile/native-zip-results.json'),
    JSON.stringify(results, null, 2),
  );
  console.log(JSON.stringify(results, null, 2));
} finally {
  if (browser) await browser.close();
  await new Promise(r => server.close(r));
}
