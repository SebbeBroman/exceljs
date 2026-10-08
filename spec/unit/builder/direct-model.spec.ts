import {describe, it, expect} from 'vite-plus/test';
import {readFile} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {normalizeWorkbook} from '../../utils/normalize-workbook.js';

import {unzipSync, zipSync, strFromU8, strToU8} from 'fflate';
import {workbook, writeBuffer, load} from '../../../excel.js';
import type {WorkbookBuilder} from '../../../lib/builder/workbook-builder.js';
import type {Workbook} from '../../../lib/model/types.js';

const meta = {creator: 'parity', created: new Date('2020-01-01'), modified: new Date('2020-01-01')};
function parts(buffer: Uint8Array): Record<string, string | number[]> {
  return Object.fromEntries(Object.entries(unzipSync(buffer)).map(([name, bytes]) => [name, /\.(xml|rels)$/.test(name) ? strFromU8(bytes) : [...bytes]]));
}
const goldenDir = new URL('./data/direct-model/', import.meta.url);
const hashes: Record<string, string> = JSON.parse(await readFile(new URL('projections.json', goldenDir), 'utf8'));
function digest(v: unknown): string {return createHash('sha256').update(normalizeWorkbook(v)).digest('hex');}
async function baselineWrite(builder: WorkbookBuilder, sharedStrings = true): Promise<Uint8Array> {
  // Preserve the historical fixture lookup; compare its OOXML with the new prepared model.
  const ops = builder._ops.map(op => op.op === 'sheetProtection' && !op.model.hashValue
    ? {op: 'protect', sheet: op.sheet, options: Object.fromEntries(Object.entries(op.model).filter(([key]) => key !== 'sheet'))}
    : op);
  return readFile(new URL(`${digest({ops, sharedStrings})}.xlsx`, goldenDir));
}
async function assertBaselineProjection(bytes: Uint8Array, options?: {ignoreNodes: string[]}): Promise<void> {
  expect(digest(await load(bytes, options))).toBe(hashes[digest(parts(bytes))]);
}
async function parity(builder: WorkbookBuilder, sharedStrings = true): Promise<void> {
  const old = await baselineWrite(builder, sharedStrings);
  const direct = await writeBuffer(builder, {useSharedStrings: sharedStrings});
  expect(parts(direct)).toEqual(parts(old));
  await assertBaselineProjection(old);
  await assertBaselineProjection(direct);
}

describe('direct OOXML model compatibility', () => {
  it('preserves default metadata when empty or undefined fields are supplied', async () => {
    await parity(workbook({...meta, creator: '', lastModifiedBy: '', title: undefined}).sheet('S').row([1]));
  });
  it('matches legacy XML and snapshots for sparse values, dates, formulas, rich text and errors', async () => {
    await parity(workbook(meta).sheet('Values').rows([
      ['text', 123, false, null, new Date('2022-05-06'), {error: '#N/A'}],
      [{richText: [{text: 'bold', font: {bold: true}}, {text: ' plain'}]}],
    ]).cell('C7', {text: 'link', hyperlink: 'https://example.com', tooltip: 'tip'})
      .cell('D7', {formula: 'B1*2', result: 246})
      .cell('E7', {formula: 'B1+1', result: 124, shareType: 'shared', ref: 'E7:E8'} as {formula: string; result: number})
      .cell('E8', {sharedFormula: 'E7', result: 125}).cell('Z10000', 'far'));
  });
  it('matches legacy inline strings and keyed columns with appended headers', async () => {
    await parity(workbook(meta).sheet('Data').row(['title']).columns([
      {header: ['Name', 'ignored'], key: 'name', width: 24, style: {font: {italic: true}}},
      {header: 'Score', key: 'score', hidden: true, outlineLevel: 1},
    ]).row({name: 'Ada', score: 10}).row({name: 'Bob', score: 20}).style('A3', {font: {bold: true}}).style('3', {numFmt: '0.00'}), false);
  });
  it('matches legacy merged styles, notes, validation, conditional formatting and print settings', async () => {
    await parity(workbook(meta).sheet('Features').rows([[1, 2], [3, 4]])
      .style('A1:B2', {font: {bold: true}, border: {bottom: {style: 'thin'}}})
      .style('A1', {font: {italic: true}}).merge('C3:D4').cell('C3', 'merged')
      .note('A1', 'comment').dataValidation('B2', {type: 'whole', operator: 'between', formulae: [1, 10]})
      .conditionalFormatting({ref: 'A1:B2', rules: [{type: 'expression', priority: 1, formulae: ['TRUE'], style: {font: {italic: true}}}]})
      .views([{state: 'frozen', ySplit: 1}]).pageSetup({orientation: 'landscape', printArea: 'A1:D4', printTitlesRow: '1:1'})
      .headerFooter({oddHeader: 'Header', oddFooter: '&P'}).protect({sheet: true, selectLockedCells: false})
      .definedName('Cells', 'Features!$A$1:$B$2').sheet('Empty').sheet('Last').row([42]));
  });
  it('matches legacy table placement and totals, including later cell edits', async () => {
    await parity(workbook(meta).sheet('Tables').row(['before']).table({name: 'People', ref: 'B3', totalsRow: true,
      columns: [{name: 'Name'}, {name: 'Score', totalsRowFunction: 'sum', totalsRowResult: 30, style: {numFmt: '0.00'}} as {name: string; totalsRowFunction: 'sum'; style: {numFmt: string}}],
      rows: [['Ada', 10], ['Bob', 20]],
    }).cell('C4', 15, {font: {bold: true}}).row(['after']));
  });
  it('matches legacy XML for image anchors with custom column widths', async () => {
    const builder = workbook(meta).sheet('Images').columns([{width: 24}, {width: 12}]).row(['image']);
    const id = builder.image({extension: 'png', buffer: Buffer.from('iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==', 'base64')});
    builder.image(id, {tl: {col: 0.5, row: 1.5}, br: {col: 2.5, row: 4.5}, hyperlinks: {hyperlink: 'https://example.com', tooltip: 'image'}});
    const old = await baselineWrite(builder);
    const direct = await writeBuffer(builder);
    expect(parts(direct)).toEqual(parts(old));
    const loaded = await load(old);
    expect(loaded.sheets[0].images).toHaveLength(1);
    const reencoded = await writeBuffer(loaded);
    expect(digest((await load(reencoded)).sheets[0].images)).toBe(digest(loaded.sheets[0].images));
  });
  it('writes plain row metadata without mutating the snapshot or adding column headers', async () => {
    const snapshot: Workbook = {meta, sheets: [{id: 7, name: 'Plain', columns: [{header: 'Do not insert', width: 20}], rows: [
      {number: 2, height: 30, hidden: true, style: {font: {bold: true}}, cells: {1: {value: 'kept'}}},
      {number: 4, height: 25, cells: {}},
    ]}]};
    const copy = structuredClone(snapshot);
    const bytes = await writeBuffer(snapshot);
    expect(snapshot).toEqual(copy);
    const loaded = await load(bytes);
    expect(loaded.sheets[0].rows.map(r => r.number)).toEqual([2]);
    expect(loaded.sheets[0].rows[0]).toMatchObject({height: 30, hidden: true});
    expect(loaded.sheets[0].rows[0].cells[1].style?.font?.bold).toBe(true);
    expect(parts(bytes)['xl/worksheets/sheet1.xml']).toContain('r="4" ht="25"');
  });
  it('matches legacy notes and styles on blank cells', async () => {
    await parity(workbook(meta).sheet('Blank').note('A1', 'only note').style('B2', {font: {bold: true}}).row([]).row([null]).row(['kept']).note('D5', 'blank cell note'));
  });
  it('matches legacy interleaved row and cell edits across sheets', async () => {
    let seed = 1729;
    const random = (max: number) => {seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0; return seed % max;};
    for (let fixture = 0; fixture < 12; fixture++) {
      const builder = workbook(meta).sheet('A').columns([{key: 'name', style: {font: {italic: true}}}, {key: 'value'}]).sheet('B');
      for (let step = 0; step < 25; step++) {
        builder.sheet(random(2) ? 'A' : 'B');
        const address = `${String.fromCharCode(65 + random(4))}${1 + random(12)}`;
        switch (random(5)) {
          case 0: builder.row([`r${step}`, random(100)]); break;
          case 1: builder.cell(address, random(100)); break;
          case 2: builder.cell(address, `v${step}`, {font: {bold: true}}); break;
          case 3: builder.style(address, {alignment: {horizontal: 'center'}}); break;
          case 4: builder.row({name: `k${step}`, value: step}); break;
        }
      }
      await parity(builder);
    }
  });
  it('matches legacy row filtering in existing Excel fixtures', async () => {
    for (const file of ['1519293514-KRISHNAPATNAM_LINE_UP.xlsx', 'duplicateRowTest.xlsx']) {
      const bytes = await readFile(new URL(`../../integration/data/${file}`, import.meta.url));
      await assertBaselineProjection(bytes);
    }
  });
  it('matches legacy projections when ignored nodes or addressless cells use the classic parser', async () => {
    const bytes = await baselineWrite(workbook(meta).sheet('S').rows([[1, 2], [3, 4]]));
    const files = unzipSync(bytes);
    const xml = strFromU8(files['xl/worksheets/sheet1.xml']).replace('r="B1"', '');
    files['xl/worksheets/sheet1.xml'] = strToU8(xml);
    const patched = zipSync(files);
    await assertBaselineProjection(patched, {ignoreNodes: ['dimension']});
  });
});
