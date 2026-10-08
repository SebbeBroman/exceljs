import {describe, it, expect, afterEach} from 'vite-plus/test';

import {workbook, writeBuffer, load} from '../../../excel.js';
import type {Workbook} from '../../../lib/model/types.js';
import {
  canUseFastSheetData,
  isFastSheetDataEnabled,
  setFastSheetDataEnabled,
  splitSheetData,
} from '../../../lib/xlsx/xform/sheet/fast-sheet-data.js';
import {entryToString, unzipToFiles} from '../../../lib/utils/zip-reader.js';

async function sheetXmlOf(buf: Uint8Array): Promise<string> {
  const files = await unzipToFiles(buf);
  return entryToString(files['xl/worksheets/sheet1.xml']!);
}

async function loadBoth(buf: Uint8Array): Promise<{fast: Workbook; full: Workbook}> {
  setFastSheetDataEnabled(true);
  const fast = await load(buf);
  setFastSheetDataEnabled(false);
  try {
    const full = await load(buf);
    return {fast, full};
  } finally {
    setFastSheetDataEnabled(true);
  }
}

describe('fast-sheet-data', () => {
  afterEach(() => {
    setFastSheetDataEnabled(true);
  });

  it('is enabled by default', () => {
    expect(isFastSheetDataEnabled()).toBe(true);
  });

  it('splits sheet XML at sheetData', async () => {
    const buf = await writeBuffer(workbook().sheet('S').rows([[1, 2]]));
    const xml = await sheetXmlOf(buf);
    const split = splitSheetData(xml);
    expect(split).not.toBeNull();
    expect(split!.head).toContain('<worksheet');
    expect(split!.content).toContain('<row');
    expect(split!.tail).toContain('</worksheet>');
    expect(canUseFastSheetData(split!.content)).toBe(true);
  });

  it('rejects denylisted sheetData content', () => {
    expect(canUseFastSheetData('<row><c r="A1"><is><r><rPr><sz val="11"/></rPr><t>x</t></r></is></c></row>')).toBe(
      false,
    );
    expect(canUseFastSheetData('<row><c r="A1"><is><r><rPh>foo</rPh><t>x</t></r></is></c></row>')).toBe(
      false,
    );
    expect(canUseFastSheetData('<row><c r="A1"><v>1</v></c></row><extLst><ext>x</ext></extLst>')).toBe(false);
    expect(canUseFastSheetData('<row><c r="A1"><v>1</v></c></row>')).toBe(true);
  });

  it('returns null for malformed splits', () => {
    expect(splitSheetData('<worksheet><sheetDatas>')).toBeNull();
    expect(splitSheetData('<worksheet><sheetData><row/></worksheet>')).toBeNull();
  });

  it('matches classic load on styled values, dates, booleans, errors', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('S')
        .rows([
          ['name', 'score', 'active', 'ratio', 'when', 'bad'],
          ['Ada', 95, true, 1.5, new Date(Date.UTC(2024, 0, 15)), {error: '#N/A'}],
          ['Bob', 80, false, 2.25, new Date(Date.UTC(2023, 5, 30, 12, 30)), {error: '#DIV/0!'}],
        ])
        .style('A1:F1', {font: {bold: true}})
        .style('B2:B3', {numFmt: '0.00'}),
    );
    const {fast, full} = await loadBoth(buf);
    expect(fast).toEqual(full);
  });

  it('matches on date-formatted numbers, shared formulas, hyperlinks, notes', async () => {
    const b = workbook().sheet('S');
    b.cell('A1', 45000, {numFmt: 'yyyy-mm-dd'});
    b.cell('A2', {formula: 'A1+1', result: 45001});
    b.cell('A3', {formula: 'A1+2', result: 45002});
    b.cell('B1', {text: 'click', hyperlink: 'https://example.com'});
    b.note('C1', 'a note');
    b.merge('D1:E2');
    b.cell('D1', 'merged');
    const buf = await writeBuffer(b);
    const {fast, full} = await loadBoth(buf);
    expect(fast).toEqual(full);
    // Sanity: the interesting values actually survived.
    const row1 = fast.sheets[0]!.rows.find(r => r.number === 1)!;
    expect(row1.cells[1]!.value).toBeInstanceOf(Date);
    expect(row1.cells[2]!.value).toEqual({text: 'click', hyperlink: 'https://example.com'});
  });

  it('matches on rich text, inline strings, row attrs, cols, merges', async () => {
    const b = workbook().sheet('S');
    b.cell('A1', {richText: [{text: 'plain '}, {text: 'bold'}]});
    b.cell('A3', 'sparse');
    b.cell('C3', 42);
    const buf = await writeBuffer(b);
    const {fast, full} = await loadBoth(buf);
    expect(fast).toEqual(full);
  });

  it('matches on multi-sheet workbooks with per-sheet features', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('A', s =>
          s
            .rows([
              ['x', 1],
              ['y', 2],
            ])
            .columns([{header: 'X'}, {header: 'Y'}]),
        )
        .sheet('Empty')
        .sheet('C', s => s.cell('B2', 'far').dataValidation('B2', {type: 'whole', formulae: [1, 5]})),
    );
    const {fast, full} = await loadBoth(buf);
    expect(fast).toEqual(full);
    expect(fast.sheets.map(s => s.name)).toEqual(['A', 'Empty', 'C']);
  });

  it('falls back safely on denylisted content', async () => {
    // Inject inline rich text with per-run fonts (rPr) straight into the
    // package: the fused path must decline and the classic path must agree.
    const {zipSync, unzipSync} = await import('fflate');
    const buf = await writeBuffer(workbook().sheet('S').cell('A1', 'hi'));
    const files = unzipSync(buf) as Record<string, Uint8Array>;
    const xml = new TextDecoder().decode(files['xl/worksheets/sheet1.xml']!);
    const injected =
      '<row r="1"><c r="A1" t="inlineStr"><is><r><rPr><b/></rPr><t>bold</t></r>' +
      '<r><t>plain</t></r></is></c><c r="B1"><v>42</v></c></row>';
    const patched = xml.replace(/<sheetData>[\s\S]*<\/sheetData>/, `<sheetData>${injected}</sheetData>`);
    expect(patched).not.toBe(xml);
    expect(canUseFastSheetData(injected)).toBe(false);
    files['xl/worksheets/sheet1.xml'] = new TextEncoder().encode(patched);
    const patchedBuf = zipSync(files);
    const {fast, full} = await loadBoth(patchedBuf);
    expect(fast).toEqual(full);
  });

  it('fast parser agrees with classic on the bench-shaped grid', async () => {
    const grid: unknown[][] = [];
    for (let r = 0; r < 500; r++) {
      const row: unknown[] = new Array(8);
      for (let c = 0; c < 8; c++) row[c] = c === 0 ? `r${r}` : r * 8 + c;
      grid.push(row);
    }
    const buf = await writeBuffer(workbook().sheet('data').rows(grid as never), {
      useSharedStrings: true,
      useStyles: false,
    });
    const xml = await sheetXmlOf(buf);
    const split = splitSheetData(xml)!;
    expect(canUseFastSheetData(split.content)).toBe(true);
    const {fast, full} = await loadBoth(buf);
    expect(fast).toEqual(full);
  });
});
