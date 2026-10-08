import {csv} from '../../../csv.js';
import {describe, it, expect} from 'vite-plus/test';

import {mkdtemp, readFile, rm, writeFile as fsWriteFile} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {workbook, load} from '../../../excel.js';
import {readCsvFile, writeCsvFile} from '../../../node.js';

describe('csv named API', () => {
  it('parses CSV text into SheetInit for builder.sheet', async () => {
    const init = await csv.parse('name,value\nalpha,1\nbeta,2');
    expect(init.rows).toHaveLength(3);

    const plain = workbook().sheet('Data', init).build();
    expect(plain.sheets[0]!.name).toBe('Data');
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('name');
    expect(plain.sheets[0]!.rows[1]!.cells[1]!.value).toBe('alpha');
    expect(plain.sheets[0]!.rows[1]!.cells[2]!.value).toBe(1);
  });

  it('parse → builder → writeBuffer round-trips values', async () => {
    const init = await csv.parse('a,b\n10,20');
    const buf = await workbook().sheet('S', init).writeBuffer();
    const plain = await load(buf);
    expect(plain.sheets[0]!.rows[0]!.cells[1]!.value).toBe('a');
    expect(plain.sheets[0]!.rows[1]!.cells[2]!.value).toBe(20);
  });

  it('stringifies a builder sheet', async () => {
    const text = await csv.stringify(
      workbook()
        .sheet('Data')
        .rows([
          ['name', 'value'],
          ['alpha', 1],
        ]),
    );
    expect(text).toBe('name,value\nalpha,1');
  });

  it('csv.stringify exports an explicitly selected sheet', async () => {
    const text = await csv.stringify(workbook()
      .sheet('A')
      .row([1, 2])
      .sheet('B')
      .row(['x', 'y']), {sheetName: 'B'});
    expect(text).toBe('x,y');
  });

  it('csv.stringify respects sheetName', async () => {
    const wb = workbook()
      .sheet('A')
      .row([1])
      .sheet('B')
      .row([2])
      .build();
    const text = await csv.stringify(wb, {sheetName: 'A'});
    expect(text).toBe('1');
  });

  it('quotes fields that need it', async () => {
    const text = await csv.stringify(workbook().sheet('S').row(['hello, world', 'a"b']));
    expect(text).toBe('"hello, world","a""b"');
  });

  it('maps formula results and dates on stringify', async () => {
    const text = await csv.stringify(
      workbook()
        .sheet('S')
        .cell('A1', {formula: '1+1', result: 2})
        .cell('B1', new Date(Date.UTC(2020, 0, 15, 0, 0, 0))),
      {dateFormat: 'YYYY-MM-DD', dateUTC: true},
    );
    expect(text).toBe('2,2020-01-15');
  });
});

describe('node csv file helpers', () => {
  it('readCsvFile / writeCsvFile round-trip', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-csv-'));
    const inPath = join(dir, 'in.csv');
    const outPath = join(dir, 'out.csv');
    try {
      await fsWriteFile(inPath, 'a,b\n1,2\n', 'utf8');
      const wb = await readCsvFile(inPath, {sheetName: 'Import'});
      expect(wb.sheets[0]!.name).toBe('Import');
      expect(wb.sheets[0]!.rows[1]!.cells[1]!.value).toBe(1);

      await writeCsvFile(outPath, wb);
      const written = await readFile(outPath, 'utf8');
      expect(written.replace(/\r\n/g, '\n')).toBe('a,b\n1,2');
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });
});
