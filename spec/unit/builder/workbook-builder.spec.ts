import {describe, it, expect} from 'vitest';
import {workbook} from '../../../excel.ts';
import {writeBuffer} from '../../../lib/xlsx/write-buffer.ts';
import {compileToPlainWorkbook} from '../../../lib/compile/ops-to-model.ts';
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

    expect(plain.meta.creator).to.equal('test');
    expect(plain.sheets).to.have.length(1);
    expect(plain.sheets[0]!.name).to.equal('Data');
    expect(plain.sheets[0]!.rows).to.have.length(2);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).to.equal('a');
    expect(plain.sheets[0]!.rows[0]!.cells[2]!.value).to.equal(1);
  });

  it('supports callback sheet form', () => {
    const plain = workbook()
      .sheet('A', s => s.row([1]).row([2]))
      .sheet('B', s => s.row(['x']))
      .build();

    expect(plain.sheets.map(s => s.name)).to.deep.equal(['A', 'B']);
    expect(plain.sheets[0]!.rows).to.have.length(2);
    expect(plain.sheets[1]!.rows[0]!.cells[1]!.value).to.equal('x');
  });

  it('requires an active sheet for row ops', () => {
    expect(() => workbook().row([1])).to.throw(/No active sheet/);
  });

  it('applies cell and merge ops into the plain model', () => {
    const plain = workbook()
      .sheet('S')
      .cell('A1', 'hello')
      .cell('B2', 42)
      .merge('A1:B1')
      .build();

    expect(plain.sheets[0]!.rows.find(r => r.number === 1)!.cells[1]!.value).to.equal('hello');
    expect(plain.sheets[0]!.rows.find(r => r.number === 2)!.cells[2]!.value).to.equal(42);
    expect(plain.sheets[0]!.merges).to.deep.equal(['A1:B1']);
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

    expect(buf.byteLength).to.be.greaterThan(1000);
    const files = unzipSync(buf);
    const names = Object.keys(files);
    expect(names.some(n => n.includes('workbook.xml'))).to.equal(true);
    expect(names.some(n => n.includes('sheet'))).to.equal(true);

    // Strings may live in sharedStrings.xml depending on write options
    const allXml = names
      .filter(n => n.endsWith('.xml'))
      .map(n => strFromU8(files[n]!))
      .join('\n');
    expect(allXml).to.include('Widgets');
    expect(allXml).to.include('Product');
  });

  it('writeBuffer free function accepts a builder', async () => {
    const b = workbook().sheet('X').row([1, 2, 3]);
    const buf = await writeBuffer(b);
    expect(buf.byteLength).to.be.greaterThan(500);
  });

  it('optimizes empty rows arrays away', () => {
    const b = workbook().sheet('S').rows([]).row([1]);
    const plain = compileToPlainWorkbook(b._ops);
    expect(plain.sheets[0]!.rows).to.have.length(1);
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

    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).to.equal('Name');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).to.equal('Ada');
  });
});
