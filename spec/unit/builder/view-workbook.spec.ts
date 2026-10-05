import {describe, it, expect} from 'vite-plus/test';
import {
  viewWorkbook,
  isWorkbookView,
  readRows,
  workbook,
  writeBuffer,
} from '../../../excel.ts';

describe('viewWorkbook', () => {
  it('opens CSV and exposes sheetNames + rows slices', async () => {
    const view = await viewWorkbook('a,b,c\n1,2,3\n4,5,6\n7,8,9', {format: 'csv'});
    expect(isWorkbookView(view)).to.equal(true);
    expect(view.format).to.equal('csv');
    expect(view.sheetNames).to.deep.equal(['Sheet1']);

    const sheet = view.sheet(0);
    expect(sheet.name).to.equal('Sheet1');
    expect(sheet.rows()).to.deep.equal([
      ['a', 'b', 'c'],
      ['1', '2', '3'],
      ['4', '5', '6'],
      ['7', '8', '9'],
    ]);
    expect(sheet.rows({start: 2, end: 3, cols: {start: 1, end: 2}})).to.deep.equal([
      ['1', '2'],
      ['4', '5'],
    ]);
    expect(sheet.rows({start: 2, end: 3, cols: ['A', 'C']})).to.deep.equal([
      ['1', '3'],
      ['4', '6'],
    ]);
  });

  it('records() uses header row of the slice', async () => {
    const view = await viewWorkbook('name,score\nAda,98\nBob,70', {format: 'csv'});
    const recs = view.sheet(0).records({header: true});
    expect(recs).to.deep.equal([
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
    expect(view.sheetNames).to.deep.equal(['Roster']);
    expect(view.sheet('Roster').rows({start: 2, end: 2})).to.deep.equal([['Ada', 'Blue']]);
  });

  it('workbook(view) builds a writable builder', async () => {
    const view = await viewWorkbook('x,y\n1,2', {format: 'csv'});
    const out = await workbook(view).sheet('Sheet1').cell('A1', 'z').writeBuffer();
    const again = await viewWorkbook(out, {format: 'xlsx'});
    expect(again.sheet(0).rows({start: 1, end: 1})[0]?.[0]).to.equal('z');
  });

  it('readRows is sugar over the view', async () => {
    const rows = await readRows('a,b\n1,2', {format: 'csv'});
    expect(rows).to.deep.equal([
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
    expect(err?.message).to.match(/Unsupported format/);
  });

  it('reads only the requested sheet in a multi-sheet workbook', async () => {
    const buf = await writeBuffer(
      workbook()
        .sheet('A', s => s.rows([['only-a'], ['a2']]))
        .sheet('B', s => s.rows([['only-b'], ['b2'], ['b3']])),
    );
    const view = await viewWorkbook(buf, {format: 'xlsx'});
    expect(view.sheetNames).to.deep.equal(['A', 'B']);
    // Only touch sheet B — lazy path must not require materializing A
    const rows = view.sheet('B').rows();
    expect(rows).to.deep.equal([['only-b'], ['b2'], ['b3']]);
    expect(view.sheet(1).rows({start: 2, end: 2})).to.deep.equal([['b2']]);
  });

  it('end-limited slice returns the first N rows', async () => {
    const data = Array.from({length: 50}, (_, i) => [`r${i + 1}`, i + 1]);
    const buf = await writeBuffer(workbook().sheet('Big').rows(data));
    const view = await viewWorkbook(buf, {format: 'xlsx'});
    const rows = view.sheet(0).rows({start: 1, end: 5});
    expect(rows).to.have.length(5);
    expect(rows[0]).to.deep.equal(['r1', '1']);
    expect(rows[4]).to.deep.equal(['r5', '5']);
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
    expect(recs).to.deep.equal([
      {Name: 'Ada', Score: '98'},
      {Name: 'Bob', Score: '70'},
    ]);
    const out = await workbook(view).writeBuffer();
    const again = await viewWorkbook(out, {format: 'xlsx'});
    expect(again.sheetNames).to.deep.equal(['Roster']);
    expect(again.sheet(0).rows()).to.deep.equal([
      ['Name', 'Score'],
      ['Ada', '98'],
      ['Bob', '70'],
    ]);
  });
});
