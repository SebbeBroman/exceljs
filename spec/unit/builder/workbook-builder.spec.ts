import {describe, it, expect} from 'vite-plus/test';

import {workbook, load} from '../../../excel.js';
import {writeBuffer} from '../../../lib/xlsx/write-buffer.js';
import {compileToPlainWorkbook} from '../../../lib/compile/ops-to-model.js';
import {unzipSync, strFromU8} from 'fflate';

describe('workbook builder', () => {
  it('records sheet rows and builds a plain snapshot', () => {
    const plain = workbook({creator: 'test'})
      .sheet('Data')
      .rows([
        ['a', 1],
        ['b', 2],
      ])
      .build();

    expect(plain.meta.creator).toBe('test');
    expect(plain.sheets).toHaveLength(1);
    expect(plain.sheets[0]!.name).toBe('Data');
    expect(plain.sheets[0]!.rows).toHaveLength(2);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('a');
    expect(plain.sheets[0]!.rows[0]!.cells[2]!.value).toBe(1);
  });

  it('supports callback sheet form', () => {
    const plain = workbook()
      .sheet('A', s => s.row([1]).row([2]))
      .sheet('B', s => s.row(['x']))
      .build();

    expect(plain.sheets.map(s => s.name)).toEqual(['A', 'B']);
    expect(plain.sheets[0]!.rows).toHaveLength(2);
    expect(plain.sheets[1]!.rows[0]!.cells[1]!.value).toBe('x');
  });

  it('requires an active sheet for row ops', () => {
    expect(() => workbook().row([1])).toThrow(/No active sheet/);
  });

  it('applies cell and merge ops into the plain model', () => {
    const plain = workbook()
      .sheet('S')
      .cell('A1', 'hello')
      .cell('B2', 42)
      .merge('A1:B1')
      .build();

    expect(plain.sheets[0]!.rows.find(r => r.number === 1)!.cells[1]!.value).toBe('hello');
    expect(plain.sheets[0]!.rows.find(r => r.number === 2)!.cells[2]!.value).toBe(42);
    expect(plain.sheets[0]!.merges).toEqual(['A1:B1']);
  });

  it('writeBuffer produces a zip with workbook.xml', async () => {
    const buf = await workbook()
      .sheet('Q1')
      .rows([
        ['Product', 'Revenue'],
        ['Widgets', 12000],
      ])
      .style('A1:B1', {font: {bold: true}})
      .writeBuffer();

    expect(buf.byteLength).toBeGreaterThan(1000);
    const files = unzipSync(buf);
    const names = Object.keys(files);
    expect(names.some(n => n.includes('workbook.xml'))).toBe(true);
    expect(names.some(n => n.includes('sheet'))).toBe(true);

    // Strings may live in sharedStrings.xml depending on write options
    const allXml = names
      .filter(n => n.endsWith('.xml'))
      .map(n => strFromU8(files[n]!))
      .join('\n');
    expect(allXml).toContain('Widgets');
    expect(allXml).toContain('Product');
  });

  it('writeBuffer free function accepts a builder', async () => {
    const b = workbook().sheet('X').row([1, 2, 3]);
    const buf = await writeBuffer(b);
    expect(buf.byteLength).toBeGreaterThan(500);
  });

  it('optimizes empty rows arrays away', () => {
    const b = workbook().sheet('S').rows([]).row([1]);
    const plain = compileToPlainWorkbook(b._ops);
    expect(plain.sheets[0]!.rows).toHaveLength(1);
  });

  it('columns with headers produce a header row in plain model', () => {
    const plain = workbook()
      .sheet('People')
      .columns([
        {header: 'Name', key: 'name', width: 20},
        {header: 'Age', key: 'age', width: 10},
      ])
      .row({name: 'Ada', age: 36})
      .build();

    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('Name');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).toBe('Ada');
  });

  it('title row before columns-with-headers is kept in build and writeBuffer', async () => {
    const b = workbook()
      .sheet('Report')
      .row(['Report'])
      .columns([
        {header: 'A', key: 'a'},
        {header: 'B', key: 'b'},
      ])
      .row(['1', '2']);

    const plain = b.build();
    expect(plain.sheets[0]!.rows).toHaveLength(3);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('Report');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).toBe('A');
    expect(plain.sheets[0]!.rows[1]!.cells[2]!.value).toBe('B');
    expect(plain.sheets[0]!.rows[2]!.cells[1]!.value).toBe('1');

    const roundtrip = await load(await b.writeBuffer());
    const sheet = roundtrip.sheets[0]!;
    const byRow = (n: number) => sheet.rows.find(r => r.number === n)!;
    expect(byRow(1).cells[1]!.value).toBe('Report');
    expect(byRow(2).cells[1]!.value).toBe('A');
    expect(byRow(2).cells[2]!.value).toBe('B');
    expect(byRow(3).cells[1]!.value).toBe('1');
  });

  it('columns-first still places headers on row 1', () => {
    const plain = workbook()
      .sheet('S')
      .columns([{header: 'A'}, {header: 'B'}])
      .row(['1', '2'])
      .build();

    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('A');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).toBe('1');
  });

  it('SheetInit.title emits title before columns and rows', () => {
    const plain = workbook()
      .sheet('Report', {
        title: {
          text: 'Q1 Revenue',
          style: {font: {bold: true, size: 16}},
          merge: 'A1:B1',
        },
        columns: [
          {header: 'Product', key: 'p', width: 20},
          {header: 'Revenue', key: 'r', width: 12},
        ],
        rows: [{p: 'Widgets', r: 12000}],
      })
      .build();

    const sheet = plain.sheets[0]!;
    expect(sheet.rows[0]!.cells[1]!.value).toBe('Q1 Revenue');
    expect(sheet.rows[1]!.cells[1]!.value).toBe('Product');
    expect(sheet.rows[2]!.cells[1]!.value).toBe('Widgets');
    expect(sheet.merges).toEqual(['A1:B1']);
  });

  it('fluent .title() works before columns', () => {
    const plain = workbook()
      .sheet('S')
      .title('Hello')
      .columns([{header: 'A'}])
      .row(['1'])
      .build();

    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('Hello');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).toBe('A');
    expect(plain.sheets[0]!.rows[2]!.cells[1]!.value).toBe('1');
  });

  it('all-empty headers suppress the header row', () => {
    const plain = workbook()
      .sheet('S')
      .columns([{header: ''}, {}])
      .row(['1', '2'])
      .build();

    expect(plain.sheets[0]!.rows).toHaveLength(1);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('1');
  });

  it('styled title applies style to the merge range', () => {
    const plain = workbook()
      .sheet('S')
      .title({text: 'T', style: {font: {bold: true}}, merge: 'A1:B1'})
      .build();

    expect(plain.sheets[0]!.merges).toEqual(['A1:B1']);
    const row1 = plain.sheets[0]!.rows[0]!;
    expect(row1.cells[1]!.style?.font).toMatchObject({bold: true});
    expect(row1.cells[2]!.style?.font).toMatchObject({bold: true});
  });

  it('keyed rows resolve after title + columns', () => {
    const plain = workbook()
      .sheet('S', {
        title: 'T',
        columns: [
          {header: 'A', key: 'a'},
          {header: 'B', key: 'b'},
        ],
        rows: [{a: '1', b: '2'}],
      })
      .build();

    expect(plain.sheets[0]!.rows[2]!.cells[1]!.value).toBe('1');
    expect(plain.sheets[0]!.rows[2]!.cells[2]!.value).toBe('2');
  });

  it('rejects invalid title merge ranges', () => {
    expect(() => workbook().sheet('S').title({text: 'T', merge: 'not-a-range'})).toThrow(
      /Invalid title merge range/,
    );
  });

  it('sparse far cells do not OOM the coalesce path', () => {
    const plain = workbook().sheet('S').cell('A1', 'hi').cell('Z1000000', 'far').build();
    const rows = plain.sheets[0]!.rows;
    expect(rows.find(r => r.number === 1)!.cells[1]!.value).toBe('hi');
    expect(rows.find(r => r.number === 1000000)!.cells[26]!.value).toBe('far');
  });
});
