import {describe, it, expect} from 'vitest';
import {workbook, load} from '../../../excel.ts';
import {
  compileToPlainWorkbook,
  isDenseRectangularOps,
  optimizeOps,
} from '../../../lib/compile/ops-to-model.ts';
import {_isDenseMaterialize} from '../../../lib/compile/ops-to-doc-workbook.ts';
import type {BuilderOp} from '../../../lib/builder/ops.ts';

describe('optimizeOps', () => {
  it('drops empty rows and cells ops', () => {
    const ops: BuilderOp[] = [
      {op: 'sheet', name: 'S'},
      {op: 'rows', sheet: 'S', values: []},
      {op: 'cells', sheet: 'S', map: {}},
      {op: 'row', sheet: 'S', values: [1]},
    ];
    const out = optimizeOps(ops);
    expect(out.filter(o => o.op === 'rows' || o.op === 'row')).to.have.length(1);
    expect(out.some(o => o.op === 'cells')).to.equal(false);
  });

  it('fuses consecutive row/rows on the same sheet', () => {
    const ops: BuilderOp[] = [
      {op: 'sheet', name: 'S'},
      {op: 'row', sheet: 'S', values: [1]},
      {op: 'row', sheet: 'S', values: [2]},
      {op: 'rows', sheet: 'S', values: [[3], [4]]},
    ];
    const out = optimizeOps(ops);
    const rowsOps = out.filter(o => o.op === 'rows');
    expect(rowsOps).to.have.length(1);
    if (rowsOps[0]!.op === 'rows') {
      expect(rowsOps[0].values).to.have.length(4);
    }
  });

  it('does not fuse rows across sheets', () => {
    const ops: BuilderOp[] = [
      {op: 'sheet', name: 'A'},
      {op: 'row', sheet: 'A', values: [1]},
      {op: 'sheet', name: 'B'},
      {op: 'row', sheet: 'B', values: [2]},
    ];
    const out = optimizeOps(ops);
    expect(out.filter(o => o.op === 'rows')).to.have.length(2);
  });

  it('last-write-wins for the same cell address', () => {
    const ops: BuilderOp[] = [
      {op: 'sheet', name: 'S'},
      {op: 'cell', sheet: 'S', address: 'A1', value: 'old'},
      {op: 'cell', sheet: 'S', address: 'B1', value: 1},
      {op: 'cell', sheet: 'S', address: 'A1', value: 'new'},
    ];
    const out = optimizeOps(ops);
    const cells = out.filter(o => o.op === 'cell');
    expect(cells).to.have.length(2);
    const a1 = cells.find(o => o.op === 'cell' && o.address === 'A1');
    expect(a1 && a1.op === 'cell' && a1.value).to.equal('new');
  });

  it('last-write-wins across cells map and cell ops', () => {
    const ops: BuilderOp[] = [
      {op: 'sheet', name: 'S'},
      {op: 'cells', sheet: 'S', map: {A1: 'x', B1: 'y'}},
      {op: 'cell', sheet: 'S', address: 'A1', value: 'z'},
    ];
    const out = optimizeOps(ops);
    const cellsMap = out.find(o => o.op === 'cells');
    expect(cellsMap && cellsMap.op === 'cells' && cellsMap.map).to.deep.equal({B1: 'y'});
    const a1 = out.find(o => o.op === 'cell' && o.address === 'A1');
    expect(a1 && a1.op === 'cell' && a1.value).to.equal('z');
  });

  it('does not mutate original ops', () => {
    const rows: BuilderOp = {op: 'rows', sheet: 'S', values: [[1]]};
    const row: BuilderOp = {op: 'row', sheet: 'S', values: [2]};
    const ops: BuilderOp[] = [{op: 'sheet', name: 'S'}, rows, row];
    optimizeOps(ops);
    expect(rows.op === 'rows' && rows.values).to.have.length(1);
  });
});

describe('dense rectangular path', () => {
  it('detects dense ops for .rows() only', () => {
    const b = workbook().sheet('Data').rows([
      ['a', 1],
      ['b', 2],
    ]);
    expect(isDenseRectangularOps(b._ops)).to.equal(true);
    expect(_isDenseMaterialize(b._ops)).to.equal(true);
  });

  it('is not dense when styles/merges/cells present', () => {
    const styled = workbook().sheet('S').rows([[1]]).style('A1', {font: {bold: true}});
    expect(_isDenseMaterialize(styled._ops)).to.equal(false);

    const merged = workbook().sheet('S').rows([[1, 2]]).merge('A1:B1');
    expect(_isDenseMaterialize(merged._ops)).to.equal(false);

    const patched = workbook().sheet('S').rows([[1]]).cell('A2', 9);
    expect(_isDenseMaterialize(patched._ops)).to.equal(false);
  });

  it('dense write → load matches general (cell-by-cell) path values', async () => {
    const data = [
      ['name', 'qty', 'ok'],
      ['alpha', 10, true],
      ['beta', 20, false],
      ['gamma', 30, true],
    ];

    // Dense path: bulk .rows()
    const denseBuf = await workbook({creator: 'dense'})
      .sheet('Data')
      .rows(data)
      .writeBuffer();

    // General path: random cell writes for the same grid
    let gen = workbook({creator: 'dense'}).sheet('Data');
    for (let r = 0; r < data.length; r++) {
      for (let c = 0; c < data[r]!.length; c++) {
        const col = String.fromCharCode(65 + c);
        gen = gen.cell(`${col}${r + 1}`, data[r]![c]!);
      }
    }
    const generalBuf = await gen.writeBuffer();

    const densePlain = await load(denseBuf);
    const generalPlain = await load(generalBuf);

    expect(densePlain.sheets[0]!.name).to.equal(generalPlain.sheets[0]!.name);
    expect(densePlain.sheets[0]!.rows).to.have.length(generalPlain.sheets[0]!.rows.length);

    for (const row of densePlain.sheets[0]!.rows) {
      const gRow = generalPlain.sheets[0]!.rows.find(r => r.number === row.number)!;
      expect(gRow).to.not.equal(undefined);
      for (const [col, cell] of Object.entries(row.cells)) {
        expect(gRow.cells[Number(col)]!.value).to.deep.equal(cell.value);
      }
    }
  });

  it('fused consecutive .row() matches .rows() plain model', () => {
    const fused = workbook().sheet('S').row([1, 'a']).row([2, 'b']).row([3, 'c']);
    const bulk = workbook()
      .sheet('S')
      .rows([
        [1, 'a'],
        [2, 'b'],
        [3, 'c'],
      ]);

    expect(_isDenseMaterialize(fused._ops)).to.equal(true);
    const a = compileToPlainWorkbook(fused._ops);
    const b = compileToPlainWorkbook(bulk._ops);
    expect(a.sheets[0]!.rows).to.deep.equal(b.sheets[0]!.rows);
  });

  it('usedFlags.styles false for dense rows-only builder', () => {
    const b = workbook().sheet('S').rows([[1, 2]]);
    expect(b._used.styles).to.equal(false);
    expect(b._used.merges).to.equal(false);
  });
});
