import {describe, it, expect} from 'vitest';
import {readRows, workbook, writeBuffer} from '../../../excel.ts';

describe('readRows', () => {
  it('parses CSV text into string[][]', async () => {
    const rows = await readRows('name,score\nAda,98\nBob,  70  \n\n', {
      format: 'csv',
    });
    expect(rows).to.deep.equal([
      ['name', 'score'],
      ['Ada', '98'],
      ['Bob', '70'],
    ]);
  });

  it('sniffs CSV from file name', async () => {
    const enc = new TextEncoder();
    const buf = enc.encode('a,b\n1,2');
    const rows = await readRows(buf, {name: 'fighters.csv'});
    expect(rows).to.deep.equal([
      ['a', 'b'],
      ['1', '2'],
    ]);
  });

  it('parses xlsx first sheet to string[][]', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('Roster')
        .rows([
          ['Name', 'Club'],
          ['Ada', 'Blue'],
          ['Bob', 'Red'],
        ]),
    );
    const rows = await readRows(xlsx, {name: 'roster.xlsx'});
    expect(rows[0]).to.deep.equal(['Name', 'Club']);
    expect(rows[1]).to.deep.equal(['Ada', 'Blue']);
    expect(rows[2]).to.deep.equal(['Bob', 'Red']);
  });

  it('selects sheet by name', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('A', s => s.row(['only-a']))
        .sheet('B', s => s.row(['only-b'])),
    );
    const rows = await readRows(xlsx, {format: 'xlsx', sheet: 'B'});
    expect(rows[0]?.[0]).to.equal('only-b');
  });

  it('stringifies numbers and skips blank rows by default', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('S')
        .cell('A1', 'x')
        .cell('A2', 42)
        .cell('A4', 'y'), // row 3 empty
    );
    const rows = await readRows(xlsx, {format: 'xlsx'});
    // dense rows 1,2,4 — blank row 3 skipped
    expect(rows.map(r => r[0])).to.deep.equal(['x', '42', 'y']);
  });

  it('honours end slice on xlsx', async () => {
    const data = Array.from({length: 30}, (_, i) => [`r${i + 1}`, String(i + 1)]);
    const xlsx = await writeBuffer(workbook().sheet('T').rows(data));
    const rows = await readRows(xlsx, {format: 'xlsx', end: 3});
    expect(rows).to.deep.equal([
      ['r1', '1'],
      ['r2', '2'],
      ['r3', '3'],
    ]);
  });

  it('end-limited large sheet returns first N rows', async () => {
    const data = Array.from({length: 200}, (_, i) => [i + 1, `v${i + 1}`]);
    const xlsx = await writeBuffer(workbook().sheet('L').rows(data));
    const rows = await readRows(xlsx, {format: 'xlsx', start: 1, end: 10});
    expect(rows).to.have.length(10);
    expect(rows[0]).to.deep.equal(['1', 'v1']);
    expect(rows[9]).to.deep.equal(['10', 'v10']);
  });
});
