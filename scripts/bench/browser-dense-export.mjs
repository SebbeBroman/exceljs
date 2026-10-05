/* oxlint-disable no-console */
/** Plain-cell, single-sheet export experiment. Not a public API replacement. */
import {build} from 'esbuild';
import {readFileSync, writeFileSync, mkdirSync, copyFileSync} from 'node:fs';
import {join, resolve, dirname} from 'node:path';
import {fileURLToPath} from 'node:url';
import {createServer} from 'node:http';
import {gzipSync} from 'node:zlib';
import puppeteer from 'puppeteer-core';
const root = resolve(dirname(fileURLToPath(import.meta.url)), '../..');
const out = join(root, 'build/browser-dense-export');
mkdirSync(out, {recursive: true});
const entry = `
import {zipSync,strToU8} from 'fflate';
const head='<?xml version="1.0" encoding="UTF-8" standalone="yes"?>';
const ns='http://schemas.openxmlformats.org/spreadsheetml/2006/main';
const rel='http://schemas.openxmlformats.org/package/2006/relationships';
const office='http://schemas.openxmlformats.org/officeDocument/2006/relationships';
const escapes={'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&apos;'};
function escape(s){return s.replace(/[&<>"']/g,x=>escapes[x]).replace(/[\\x00-\\x08\\x0B\\x0C\\x0E-\\x1F\\x7F]/g,'');}
function colName(n){let name='';for(;n>0;n=Math.floor((n-1)/26))name=String.fromCharCode(65+(n-1)%26)+name;return name;}
export function writeRows(rows){
 const strings=[],indices=new Map(),xml=[];let max=0,count=0;
 const names=[];
 for(let r=0;r<rows.length;r++){
  const row=rows[r];max=Math.max(max,row.length);xml.push('<row r="'+(r+1)+'">');
  for(let c=0;c<row.length;c++){
   const value=row[c];if(value==null)continue;const address=(names[c]||(names[c]=colName(c+1)))+(r+1);
   if(typeof value==='string'){
    let index=indices.get(value);if(index===undefined){index=strings.length;indices.set(value,index);strings.push(value);}count++;
    xml.push('<c r="'+address+'" t="s"><v>'+index+'</v></c>');
   }else if(typeof value==='number'&&Number.isFinite(value)){
    xml.push('<c r="'+address+'"><v>'+value+'</v></c>');
   }else if(typeof value==='boolean'){
    xml.push('<c r="'+address+'" t="b"><v>'+(value?1:0)+'</v></c>');
   }else{throw new Error('Prototype supports only string/finite-number/boolean/null cells');}
  }xml.push('</row>');
 }
 const parts={
  '[Content_Types].xml':head+'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/></Types>',
  '_rels/.rels':head+'<Relationships xmlns="'+rel+'"><Relationship Id="rId1" Type="'+office+'/officeDocument" Target="xl/workbook.xml"/></Relationships>',
  'xl/workbook.xml':head+'<workbook xmlns="'+ns+'" xmlns:r="'+office+'"><sheets><sheet name="data" sheetId="1" r:id="rId1"/></sheets></workbook>',
  'xl/_rels/workbook.xml.rels':head+'<Relationships xmlns="'+rel+'"><Relationship Id="rId1" Type="'+office+'/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="'+office+'/sharedStrings" Target="sharedStrings.xml"/></Relationships>',
  'xl/sharedStrings.xml':head+'<sst xmlns="'+ns+'" count="'+count+'" uniqueCount="'+strings.length+'">'+strings.map(s=>'<si><t xml:space="preserve">'+escape(s)+'</t></si>').join('')+'</sst>',
  'xl/worksheets/sheet1.xml':head+'<worksheet xmlns="'+ns+'"><dimension ref="A1:'+colName(Math.max(1,max))+Math.max(1,rows.length)+'"/><sheetData>'+xml.join('')+'</sheetData></worksheet>',
 };
 return zipSync(Object.fromEntries(Object.entries(parts).map(([name,text])=>[name,strToU8(text)])),{level:1});
}
`;
await build({
  bundle: true,
  minify: true,
  platform: 'browser',
  format: 'esm',
  stdin: {contents: entry, resolveDir: root},
  outfile: join(out, 'dense.mjs'),
});
for (const name of ['local-read-write.mjs', 'sheetjs-read-write.mjs', 'exceljs.min.js'])
  copyFileSync(join(root, 'build/browser-libraries', name), join(out, name));
writeFileSync(join(out, 'index.html'), '<!doctype html><script src="./exceljs.min.js"></script>');
const server = createServer((req, res) => {
  const name = req.url === '/' ? 'index.html' : req.url.slice(1);
  if (
    ![
      'index.html',
      'dense.mjs',
      'local-read-write.mjs',
      'sheetjs-read-write.mjs',
      'exceljs.min.js',
    ].includes(name)
  ) {
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
  const cases = await page.evaluate(async () => {
    const dense = await import('./dense.mjs'),
      local = await import('./local-read-write.mjs'),
      XLSX = await import('./sheetjs-read-write.mjs');
    const result = {};
    for (const count of [2000, 20000]) {
      const rows = Array.from({length: count}, (_, r) =>
        Array.from({length: 8}, (_, c) => (c === 0 ? 'r' + r : r * 8 + c)),
      );
      const bytes = dense.writeRows(rows);
      const expected = JSON.stringify(rows);
      const model = await local.load(bytes);
      if (
        JSON.stringify(
          model.sheets[0].rows.map(r => Array.from({length: 8}, (_, c) => r.cells[c + 1].value)),
        ) !== expected
      )
        throw new Error('local read');
      const wb = XLSX.read(bytes, {type: 'array', dense: true});
      if (
        JSON.stringify(XLSX.utils.sheet_to_json(wb.Sheets.data, {header: 1, raw: true})) !==
        expected
      )
        throw new Error('SheetJS read');
      const ex = new window.ExcelJS.Workbook();
      await ex.xlsx.load(bytes);
      if (
        JSON.stringify(
          ex.worksheets[0]
            .getSheetValues()
            .slice(1)
            .map(r => r.slice(1)),
        ) !== expected
      )
        throw new Error('ExcelJS read');
      const tasks = {
        dense: () => dense.writeRows(rows),
        full: () =>
          local.writeBuffer(local.workbook().sheet('data').rows(rows), {
            useSharedStrings: true,
            useStyles: false,
          }),
      };
      const samples = {dense: [], full: []};
      for (let i = 0; i < 9; i++)
        for (const name of i % 2 ? ['full', 'dense'] : ['dense', 'full']) {
          const start = performance.now();
          await tasks[name]();
          if (i >= 2) samples[name].push(performance.now() - start);
        }
      result[count] = {
        outputBytes: bytes.length,
        validatedReaders: 3,
        timings: Object.fromEntries(
          Object.entries(samples).map(([name, x]) => [
            name,
            {medianMs: [...x].sort((a, b) => a - b)[3], samplesMs: x},
          ]),
        ),
      };
    }
    const special = [
      ['é😀', 'a&<b', ' leading ', true, null],
      ['two\nlines', -2, 0, false, ''],
    ];
    const wb = XLSX.read(dense.writeRows(special), {type: 'array', dense: true});
    const cells = XLSX.utils.sheet_to_json(wb.Sheets.data, {header: 1, raw: true, defval: null});
    if (JSON.stringify(cells) !== JSON.stringify(special)) throw new Error('Special values');
    return result;
  });
  const raw = readFileSync(join(out, 'dense.mjs'));
  const report = {
    raw: raw.length,
    gzip: gzipSync(raw).length,
    cases,
    limits:
      'Single sheet, plain string/finite-number/boolean/null cells only. No styles, formulas, dates, metadata, tables, images or protection. Explicit prototype, not a replacement for the full API.',
  };
  writeFileSync(
    join(root, 'build/browser-profile/dense-export-results.json'),
    JSON.stringify(report, null, 2),
  );
  console.log(JSON.stringify(report, null, 2));
} finally {
  if (browser) await browser.close();
  await new Promise(r => server.close(r));
}
