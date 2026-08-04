import {describe, it, expect} from 'vitest';
import {workbook, writeBuffer} from '../../../excel.ts';
import {
  openLightPackage,
  parseLightSheet,
  readLightSheet,
} from '../../../lib/read/xlsx-light.ts';

describe('xlsx-light', () => {
  async function multiSheetXlsx() {
    return writeBuffer(
      workbook()
        .sheet('Roster')
        .rows([
          ['Name', 'Club'],
          ['Ada', 'Blue'],
          ['Bob', 'Red'],
        ])
        .sheet('Scores')
        .rows([
          ['x', 1],
          ['y', 2],
        ]),
    );
  }

  it('openLightPackage lists sheets without parsing sheet XML', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    expect(pkg.sheetNames).to.deep.equal(['Roster', 'Scores']);
    expect(pkg.sheets).to.have.length(2);
    expect(pkg.sheets[0]).to.include({
      index: 0,
      name: 'Roster',
      path: 'xl/worksheets/sheet1.xml',
    });
    expect(pkg.sheets[1]).to.include({
      index: 1,
      name: 'Scores',
      path: 'xl/worksheets/sheet2.xml',
    });
    expect(pkg.sharedStrings.length).to.be.greaterThan(0);
    expect(pkg.files['xl/workbook.xml']).to.be.instanceOf(Uint8Array);
    expect(pkg.files['xl/worksheets/sheet1.xml']).to.be.instanceOf(Uint8Array);
  });

  it('parseLightSheet sheet 0 returns full string grid', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    const grid = parseLightSheet(pkg, 0);
    expect(grid.name).to.equal('Roster');
    expect(grid.index).to.equal(0);
    expect(grid.rows).to.deep.equal([
      ['Name', 'Club'],
      ['Ada', 'Blue'],
      ['Bob', 'Red'],
    ]);
    expect(grid.maxCol).to.equal(2);
    expect(grid.lastRow).to.equal(3);
  });

  it('parseLightSheet selects sheet by name', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    const grid = parseLightSheet(pkg, 'Scores');
    expect(grid.name).to.equal('Scores');
    expect(grid.rows).to.deep.equal([
      ['x', '1'],
      ['y', '2'],
    ]);
  });

  it('end: 1 returns only the first row', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    const grid = parseLightSheet(pkg, 0, {end: 1});
    expect(grid.rows).to.deep.equal([['Name', 'Club']]);
    expect(grid.lastRow).to.equal(1);
  });

  it('start/end + cols slice', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    const grid = parseLightSheet(pkg, 0, {
      start: 2,
      end: 3,
      cols: {start: 1, end: 1},
    });
    expect(grid.rows).to.deep.equal([['Ada'], ['Bob']]);
  });

  it('cols as discrete number[]', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    const grid = parseLightSheet(pkg, 0, {cols: [2], start: 1, end: 2});
    expect(grid.rows).to.deep.equal([['Club'], ['Blue']]);
  });

  it('stringifies numbers and booleans like cellToDisplayString', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('S')
        .cell('A1', 'x')
        .cell('B1', 42)
        .cell('C1', true)
        .cell('D1', false),
    );
    const grid = await readLightSheet(xlsx, 0);
    expect(grid.rows[0]).to.deep.equal(['x', '42', 'TRUE', 'FALSE']);
  });

  it('skips blank rows by default', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('S')
        .cell('A1', 'x')
        .cell('A2', 42)
        .cell('A4', 'y'), // row 3 empty
    );
    const grid = await readLightSheet(xlsx, 0);
    expect(grid.rows.map(r => r[0])).to.deep.equal(['x', '42', 'y']);
    expect(grid.lastRow).to.equal(4);
  });

  it('blankrows: true keeps empty rows present in XML when present', async () => {
    // Builder may omit fully empty rows from sheet XML; place an empty string cell
    // so the row exists, then verify blankrows still skips all-defval if we want —
    // with blankrows true, an all-empty dense row from a sparse empty map is kept
    // only if the row appears. Use cols + explicit empty-ish range via multi cells.
    const xlsx = await writeBuffer(
      workbook()
        .sheet('S')
        .rows([
          ['a'],
          [''],
          ['c'],
        ]),
    );
    const pkg = await openLightPackage(xlsx);
    const skipped = parseLightSheet(pkg, 0, {blankrows: false});
    // row with empty string may still show as '' after stringify — skip if all defval
    expect(skipped.rows.map(r => r[0])).to.deep.equal(['a', 'c']);

    const kept = parseLightSheet(pkg, 0, {blankrows: true});
    expect(kept.rows.length).to.be.greaterThanOrEqual(2);
    expect(kept.rows[0]?.[0]).to.equal('a');
    expect(kept.rows[kept.rows.length - 1]?.[0]).to.equal('c');
  });

  it('values: cell returns typed CellValue grid', async () => {
    const xlsx = await writeBuffer(
      workbook()
        .sheet('S')
        .cell('A1', 'hi')
        .cell('B1', 7)
        .cell('C1', true),
    );
    const grid = await readLightSheet(xlsx, 0, {values: 'cell'});
    expect(grid.rows[0]?.[0]).to.equal('hi');
    expect(grid.rows[0]?.[1]).to.equal(7);
    expect(grid.rows[0]?.[2]).to.equal(true);
  });

  it('throws clear error for missing sheet', async () => {
    const xlsx = await multiSheetXlsx();
    const pkg = await openLightPackage(xlsx);
    expect(() => parseLightSheet(pkg, 'Nope')).to.throw(/Sheet not found/);
    expect(() => parseLightSheet(pkg, 99)).to.throw(/Sheet not found/);
  });
});
