import {describe, it, expect} from 'vite-plus/test';
import {workbook, load, writeBuffer} from '../../../excel.ts';
import {workbook as nodeWorkbook, writeFile, readFile} from '../../../node.ts';
import {mkdtemp, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';

describe('load + edit loop', () => {
  it('write → load → plain model values', async () => {
    const buf = await workbook({creator: 'roundtrip'})
      .sheet('Sheet1')
      .cell('A1', 'hello')
      .cell('B1', 42)
      .cell('C1', true)
      .cell('D1', new Date('2020-01-15T00:00:00.000Z'))
      .merge('A2:B2')
      .cell('A2', 'merged')
      .style('A1', {font: {bold: true}})
      .writeBuffer();

    const plain = await load(buf);

    expect(plain.meta.creator).to.equal('roundtrip');
    expect(plain.sheets).to.have.length(1);
    expect(plain.sheets[0]!.name).to.equal('Sheet1');

    const row1 = plain.sheets[0]!.rows.find(r => r.number === 1)!;
    expect(row1.cells[1]!.value).to.equal('hello');
    expect(row1.cells[2]!.value).to.equal(42);
    expect(row1.cells[3]!.value).to.equal(true);
    expect(row1.cells[4]!.value).to.be.instanceOf(Date);
    expect((row1.cells[4]!.value as Date).toISOString()).to.equal('2020-01-15T00:00:00.000Z');
    expect(row1.cells[1]!.style?.font?.bold).to.equal(true);

    expect(plain.sheets[0]!.merges).to.include('A2:B2');
    const row2 = plain.sheets[0]!.rows.find(r => r.number === 2)!;
    expect(row2.cells[1]!.value).to.equal('merged');
    // merge slave B2 should not appear as its own value cell
    expect(row2.cells[2]).to.equal(undefined);
  });

  it('load → workbook(data).cell → writeBuffer → load', async () => {
    const original = await workbook()
      .sheet('Sheet1')
      .rows([
        ['a', 1],
        ['b', 2],
      ])
      .writeBuffer();

    const wb = await load(original);
    expect(wb.sheets[0]!.rows[0]!.cells[1]!.value).to.equal('a');

    const out = await workbook(wb).sheet('Sheet1').cell('A1', 'updated').writeBuffer();

    const again = await load(out);
    const row1 = again.sheets[0]!.rows.find(r => r.number === 1)!;
    expect(row1.cells[1]!.value).to.equal('updated');
    expect(row1.cells[2]!.value).to.equal(1);
    const row2 = again.sheets[0]!.rows.find(r => r.number === 2)!;
    expect(row2.cells[1]!.value).to.equal('b');
    expect(row2.cells[2]!.value).to.equal(2);
  });

  it('load accepts ArrayBuffer', async () => {
    const buf = await writeBuffer(workbook().sheet('S').row([1, 2, 3]));
    const ab = buf.buffer.slice(buf.byteOffset, buf.byteOffset + buf.byteLength);
    const plain = await load(ab);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).to.equal(1);
  });

  it('preserves formulas on round-trip', async () => {
    const buf = await workbook()
      .sheet('Calc')
      .cell('A1', 10)
      .cell('B1', {formula: 'A1*2', result: 20})
      .writeBuffer();

    const plain = await load(buf);
    const b1 = plain.sheets[0]!.rows[0]!.cells[2]!.value as {formula?: string; result?: number};
    expect(b1.formula).to.equal('A1*2');
    expect(b1.result).to.equal(20);
  });
});

describe('node readFile / writeFile round-trip', () => {
  it('writeFile → readFile → edit → writeFile', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-load-'));
    const path = join(dir, 'book.xlsx');
    try {
      await writeFile(
        path,
        nodeWorkbook({creator: 'node-rt'}).sheet('Data').rows([
          ['name', 'n'],
          ['alpha', 1],
        ]),
      );

      const plain = await readFile(path);
      expect(plain.meta.creator).to.equal('node-rt');
      expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).to.equal('alpha');

      const outPath = join(dir, 'book2.xlsx');
      await writeFile(outPath, workbook(plain).sheet('Data').cell('B2', 99));
      const again = await readFile(outPath);
      expect(again.sheets[0]!.rows.find(r => r.number === 2)!.cells[2]!.value).to.equal(99);
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });
});
