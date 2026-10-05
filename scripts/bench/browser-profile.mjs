/* oxlint-disable no-console */
/** Phase timings, Chrome CPU samples, and esbuild contribution analysis.
 * Build first, then: node scripts/bench/browser-profile.mjs [--label name]
 */
import {build} from 'esbuild';
import {readFileSync, writeFileSync, mkdirSync} from 'node:fs';
import {join, resolve, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createRequire} from 'node:module';
import {createServer} from 'node:http';
import {gzipSync} from 'node:zlib';
import puppeteer from 'puppeteer-core';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/browser-profile');
mkdirSync(out, {recursive: true});
const labelIndex = process.argv.indexOf('--label');
const label = labelIndex < 0 ? 'baseline' : process.argv[labelIndex + 1];
if (!/^[\w-]+$/.test(label)) throw new Error('Invalid profile label');
const require = createRequire(import.meta.url);
const {SourceMapConsumer} = require(
  join(root, 'node_modules/.pnpm/source-map-js@1.2.1/node_modules/source-map-js'),
);
const entry = `
import {workbook,writeBuffer,load,viewWorkbook,csv} from ${JSON.stringify(join(root, 'dist/excel.js'))};
import {materializeDocWorkbook} from ${JSON.stringify(join(root, 'dist/lib/compile/ops-to-doc-workbook.js'))};
import DocWorkbook from ${JSON.stringify(join(root, 'dist/lib/doc/workbook.js'))};
import {docWorkbookToPlain} from ${JSON.stringify(join(root, 'dist/lib/compile/doc-to-plain.js'))};
import BufferZipWriter,{resolveZipLevel} from ${JSON.stringify(join(root, 'dist/lib/utils/buffer-zip.js'))};
const opts={useSharedStrings:true,useStyles:false};
const grid = n => Array.from({length:n},(_,r)=>Array.from({length:8},(_,c)=>c===0?'r'+r:r*8+c));
export async function setup(n=2000){
 const rows=grid(n);const b=workbook().sheet('data').rows(rows);const bytes=await writeBuffer(b,opts);
 const text=rows.map(r=>r.join(',')).join('\\n');
 globalThis.profileTasks={write:()=>writeBuffer(b,opts),load:()=>load(bytes),view:async()=>{const v=await viewWorkbook(bytes);return v.sheet(0).rows({values:'cell'});},csvParse:()=>csv.parse(text,{map:x=>x}),csvWrite:()=>csv.stringify(b)};
 globalThis.phaseTasks={rows,b,bytes};
}
export async function phases(){
 const {rows,b,bytes}=globalThis.phaseTasks;
 const times={};
 const sync=(key,fn)=>{const t=performance.now();const v=fn();times[key]=(times[key]||0)+performance.now()-t;return v;};
 const asyncTime=async(key,fn)=>{const t=performance.now();const v=await fn();times[key]=(times[key]||0)+performance.now()-t;return v;};
 const doc=sync('write: materialize document',()=>materializeDocWorkbook(b._ops));
 const model=sync('write: serialize document model',()=>doc.getXlsxModel?.() ?? doc.model);
 await asyncTime('write: prepare styles/strings/models',()=>doc.xlsx.prepareModel(model,opts));
 const zip=new BufferZipWriter();
 await asyncTime('write: render XML + encode UTF8',()=>doc.xlsx.renderParts(zip,model));
 sync('write: deflate + ZIP',()=>zip.toBytes(resolveZipLevel()));
 globalThis.__readPhaseTimes=times;
 const readDoc=new DocWorkbook();
 for(const [method,key] of [['_processWorksheetEntry','read: worksheet XML + fused cells'],['reconcile','read: reconcile package']]){
   const original=readDoc.xlsx[method].bind(readDoc.xlsx);
   readDoc.xlsx[method]=(...args)=>asyncTime(key,()=>original(...args));
 }
 const descriptor=Object.getOwnPropertyDescriptor(DocWorkbook.prototype,'model');
 Object.defineProperty(readDoc,'model',{get(){return descriptor.get.call(this);},set(v){sync('read: apply document model',()=>descriptor.set.call(this,v));}});
 await asyncTime('read: total document load',()=>readDoc.xlsx.load(bytes));
 sync('read: document to plain snapshot',()=>docWorkbookToPlain(readDoc));
 const view=await asyncTime('view: open ZIP + workbook/SST',()=>viewWorkbook(bytes));
 sync('view: parse/extract requested rows',()=>view.sheet(0).rows({values:'cell'}));
 return times;
}
`;
const plugin = {
  name: 'profile-unzip',
  setup(b) {
    b.onLoad({filter: /dist\/lib\/xlsx\/xlsx\.js$/}, ({path}) => {
      let s = readFileSync(path, 'utf8');
      s = s.replace(
        'const zipFiles = await unzipToFiles(buffer);',
        `const unzipStart=performance.now(); const zipFiles=await unzipToFiles(buffer); if(globalThis.__readPhaseTimes)globalThis.__readPhaseTimes['read: inflate ZIP']=performance.now()-unzipStart;`,
      );
      return {contents: s, loader: 'js', resolveDir: dirname(path)};
    });
  },
};
const file = join(out, 'profile.mjs');
await build({
  bundle: true,
  minify: true,
  keepNames: true,
  sourcemap: 'external',
  platform: 'browser',
  format: 'esm',
  target: 'es2022',
  stdin: {contents: entry, resolveDir: root},
  outfile: file,
  define: {global: 'globalThis'},
  alias: {
    fs: join(root, 'lib/shims/fs-browser.ts'),
    module: join(root, 'lib/shims/node-module-browser.ts'),
    'node:module': join(root, 'lib/shims/node-module-browser.ts'),
  },
  plugins: [plugin],
});
const consumer = new SourceMapConsumer(JSON.parse(readFileSync(file + '.map', 'utf8')));
const server = createServer((req, res) => {
  if (req.url === '/') {
    res.setHeader('Content-Type', 'text/html');
    res.end('<!doctype html><title>Browser profile</title>');
    return;
  }
  const name = req.url.slice(1);
  if (!['profile.mjs', 'profile.mjs.map'].includes(name)) {
    res.writeHead(404);
    res.end();
    return;
  }
  res.setHeader('Content-Type', 'text/javascript');
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
  await page.evaluate(async () => {
    globalThis.probe = await import('./profile.mjs');
  });
  const session = await page.createCDPSession();
  await session.send('Profiler.enable');
  await session.send('Profiler.setSamplingInterval', {interval: 100});
  const phases = {},
    cpu = {};
  for (const rows of [2000, 20000]) {
    await page.evaluate(n => probe.setup(n), rows);
    phases[rows] = await page.evaluate(async () => {
      await probe.phases();
      const all = [];
      for (let i = 0; i < 7; i++) all.push(await probe.phases());
      return Object.fromEntries(
        Object.keys(all[0]).map(key => [
          key,
          {
            medianMs: all.map(x => x[key]).sort((a, b) => a - b)[3],
            samplesMs: all.map(x => x[key]),
          },
        ]),
      );
    });
    for (const task of ['write', 'load', 'view', 'csvParse', 'csvWrite']) {
      await page.evaluate(async name => {
        for (let i = 0; i < 3; i++) await profileTasks[name]();
      }, task);
      await session.send('Profiler.start');
      const loops = rows === 2000 ? 40 : 8;
      const duration = await page.evaluate(
        async ({task, loops}) => {
          const start = performance.now();
          for (let i = 0; i < loops; i++) await profileTasks[task]();
          return performance.now() - start;
        },
        {task, loops},
      );
      const {profile} = await session.send('Profiler.stop');
      writeFileSync(join(out, `${label}-${rows}-${task}.cpuprofile`), JSON.stringify(profile));
      const nodes = new Map(profile.nodes.map(n => [n.id, n]));
      const functions = {},
        modules = {};
      let total = 0;
      for (let i = 0; i < (profile.samples || []).length; i++) {
        const node = nodes.get(profile.samples[i]);
        const frame = node.callFrame;
        const weight = (profile.timeDeltas?.[i] || 100) / 1000;
        total += weight;
        const pos = frame.url.endsWith('profile.mjs')
          ? consumer.originalPositionFor({line: frame.lineNumber + 1, column: frame.columnNumber})
          : null;
        const mod = pos?.source || frame.url || frame.functionName || '(unknown)';
        const key =
          (frame.functionName || '(anonymous)') + ' @ ' + mod + (pos?.line ? ':' + pos.line : '');
        functions[key] = (functions[key] || 0) + weight;
        modules[mod] = (modules[mod] || 0) + weight;
      }
      const summarize = xs =>
        Object.entries(xs)
          .sort((a, b) => b[1] - a[1])
          .map(([name, ms]) => ({name, selfMs: ms, share: ms / total}));
      cpu[`${rows}-${task}`] = {
        loops,
        durationMs: duration,
        sampledMs: total,
        functions: summarize(functions).slice(0, 30),
        modules: summarize(modules),
      };
      console.log(
        `${rows} ${task}: ${(duration / loops).toFixed(2)}ms/run; top self samples`,
        cpu[`${rows}-${task}`].functions.slice(0, 5),
      );
    }
  }
  const sizes = {};
  for (const name of ['local-write', 'local-read-write', 'local-view', 'local-csv']) {
    const meta = JSON.parse(
      readFileSync(join(root, 'build/browser-libraries', name + '.mjs.meta.json'), 'utf8'),
    );
    const contribution = {};
    for (const output of Object.values(meta.outputs))
      for (const [source, v] of Object.entries(output.inputs))
        contribution[source] = (contribution[source] || 0) + v.bytesInOutput;
    const group = {};
    for (const [source, bytes] of Object.entries(contribution)) {
      const key = source.includes('node_modules')
        ? 'dependencies'
        : source.includes('/xform/')
          ? 'XML transforms'
          : source.includes('/doc/')
            ? 'document model'
            : source.includes('/xlsx/xml/')
              ? 'XML templates'
              : source.includes('/xlsx/')
                ? 'XLSX orchestration'
                : source.includes('/compile/')
                  ? 'model bridges'
                  : source.includes('/builder/')
                    ? 'builder'
                    : source.includes('/read/')
                      ? 'values reader'
                      : source.includes('/fast-csv/')
                        ? 'fast-csv'
                        : source.includes('/utils/')
                          ? 'utilities'
                          : 'other';
      group[key] = (group[key] || 0) + bytes;
    }
    sizes[name] = {
      groups: group,
      modules: Object.entries(contribution)
        .sort((a, b) => b[1] - a[1])
        .map(([source, bytes]) => ({source, bytes})),
    };
  }
  const split = await build({
    bundle: true,
    minify: true,
    splitting: true,
    write: false,
    metafile: true,
    platform: 'browser',
    format: 'esm',
    target: 'es2022',
    stdin: {
      contents: `import {workbook,writeBuffer} from ${JSON.stringify(join(root, 'dist/excel.js'))}; export {workbook,writeBuffer};`,
      resolveDir: root,
    },
    outdir: join(out, 'split'),
    define: {global: 'globalThis'},
    alias: {
      fs: join(root, 'lib/shims/fs-browser.ts'),
      module: join(root, 'lib/shims/node-module-browser.ts'),
      'node:module': join(root, 'lib/shims/node-module-browser.ts'),
    },
  });
  const outputs = split.metafile.outputs;
  const initial = new Set();
  function visit(path) {
    if (initial.has(path)) return;
    initial.add(path);
    for (const dependency of outputs[path].imports)
      if (dependency.kind !== 'dynamic-import' && outputs[dependency.path]) visit(dependency.path);
  }
  const initialEntry = Object.keys(outputs).find(path => outputs[path].entryPoint === '<stdin>');
  visit(initialEntry);
  const splitFiles = split.outputFiles.map(file => {
    const key = file.path.slice(root.length + 1);
    return {
      path: key,
      raw: file.contents.length,
      gzip: gzipSync(file.contents).length,
      initial: initial.has(key),
    };
  });
  const splitWrite = {
    files: splitFiles,
    initialRaw: splitFiles.filter(x => x.initial).reduce((sum, x) => sum + x.raw, 0),
    initialGzip: splitFiles.filter(x => x.initial).reduce((sum, x) => sum + x.gzip, 0),
    totalGzip: splitFiles.reduce((sum, x) => sum + x.gzip, 0),
  };
  const report = {
    label,
    date: new Date().toISOString(),
    phases,
    cpu,
    sizes,
    splitWrite,
    notes:
      'Phase instrumentation and CPU sampling affect absolute timings. Self samples are exclusive, mapped through source maps where available. Gzip contributions are not additive; module sizes are minified bytesInOutput.',
  };
  writeFileSync(join(out, `${label}-profile.json`), JSON.stringify(report, null, 2));
  console.log(
    'phase medians',
    JSON.stringify(
      Object.fromEntries(
        Object.entries(phases).map(([n, x]) => [
          n,
          Object.fromEntries(Object.entries(x).map(([k, v]) => [k, Number(v.medianMs.toFixed(2))])),
        ]),
      ),
      null,
      2,
    ),
  );
} finally {
  if (browser) await browser.close();
  await new Promise(r => server.close(r));
}
