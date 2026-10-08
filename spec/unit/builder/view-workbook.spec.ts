import {viewCsv, readCsvRows} from '../../../csv.js';
import {describe, it, expect} from 'vite-plus/test';

import {
  viewWorkbook,
  isWorkbookView,
  workbook,
  writeBuffer,
} from '../../../excel.js';

describe('viewWorkbook', () => {
  it('opens CSV and exposes sheetNames + rows slices', async () => {
    const view = await viewCsv('a,b,c\n1,2,3\n4,5,6\n7,8,9');
    expect(isWorkbookView(view)).toBe(true);
    expect(view.format).toBe('csv');
    expect(view.sheetNames).toEqual(['Sheet1']);

    const sheet = view.sheet(0);
    expect(sheet.name).toBe('Sheet1');
    expect(sheet.rows()).toEqual([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
      ['4', '5', '6'],
      ['7', '8', '9'],
    ]);
    expect(sheet.rows({start: 2, end: 3, cols: {start: 1, end: 2}})).toEqual([
      ['1', '2'],
      ['4', '5'],
    ]);
    expect(sheet.rows({start: 2, end: 3, cols: ['A', 'C']})).toEqual([
      ['1', '3'],
      ['4', '6'],
    ]);
  });

  it('records() uses header row of the slice', async () => {
    const view = await viewCsv('name,score\nAda,98\nBob,70');
    const recs = view.sheet(0).records({header: true});
    expect(recs).toEqual([
      {name: 'Ada', score: '98'},
      {name: 'Bob', score: '70'},
    ]);
  });

  it('opens xlsx from builder bytes', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('Roster')
        .rows([
          ['Name', 'Club'],
          ['Ada', 'Blue'],
          ['Bob', 'Red'],
        ]),
    );
    const view = await viewWorkbook(buf, {filename: 'r.xlsx'});
    expect(view.sheetNames).toEqual(['Roster']);
    expect(view.sheet('Roster').rows({start: 2, end: 2})).toEqual([['Ada', 'Blue']]);
  });

  it('workbook(view) builds a writable builder', async () => {
    const view = await viewCsv('x,y\n1,2');
    const out = await workbook(view).sheet('Sheet1').cell('A1', 'z').writeBuffer();
    const again = await viewWorkbook(out, {format: 'xlsx'});
    expect(again.sheet(0).rows({start: 1, end: 1})[0]?.[0]).toBe('z');
  });

  it('readRows is sugar over the view', async () => {
    const rows = await readCsvRows('a,b\n1,2');
    expect(rows).toEqual([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('rejects legacy .xls extension', async () => {
    let err: Error | undefined;
    try {
      await viewWorkbook(new Uint8Array([1, 2, 3]), {filename: 'old.xls'});
    } catch (e) {
      err = e as Error;
    }
    expect(err?.message).toMatch(/Unsupported format/);
  });

  it('reads only the requested sheet in a multi-sheet workbook', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('A', s => s.rows([['only-a'], ['a2']]))
        .sheet('B', s => s.rows([['only-b'], ['b2'], ['b3']])),
    );
    const view = await viewWorkbook(buf, {format: 'xlsx'});
    expect(view.sheetNames).toEqual(['A', 'B']);
    // Only touch sheet B — lazy path must not require materializing A
    const rows = view.sheet('B').rows();
    expect(rows).toEqual([['only-b'], ['b2'], ['b3']]);
    expect(view.sheet(1).rows({start: 2, end: 2})).toEqual([['b2']]);
  });

  it('end-limited slice returns the first N rows', async () => {
    const data = Array.from({length: 50}, (_, i) => [`r${i + 1}`, i + 1]);
    const buf = await writeBuffer(workbook().sheet('Big').rows(data));
    const view = await viewWorkbook(buf, {format: 'xlsx'});
    const rows = view.sheet(0).rows({start: 1, end: 5});
    expect(rows).toHaveLength(5);
    expect(rows[0]).toEqual(['r1', '1']);
    expect(rows[4]).toEqual(['r5', '5']);
  });

  it('xlsx records() and workbook(view) round-trip simple values', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('Roster')
        .rows([
          ['Name', 'Score'],
          ['Ada', 98],
          ['Bob', 70],
        ]),
    );
    const view = await viewWorkbook(buf, {format: 'xlsx'});
    const recs = view.sheet(0).records({header: true});
    expect(recs).toEqual([
      {Name: 'Ada', Score: '98'},
      {Name: 'Bob', Score: '70'},
    ]);
    const out = await workbook(view).writeBuffer();
    const again = await viewWorkbook(out, {format: 'xlsx'});
    expect(again.sheetNames).toEqual(['Roster']);
    expect(again.sheet(0).rows()).toEqual([
      ['Name', 'Score'],
      ['Ada', '98'],
      ['Bob', '70'],
    ]);
  });
});
