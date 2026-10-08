import {describe, it, expect} from 'vite-plus/test';

import {createWriteStream} from 'node:fs';
import {mkdtemp, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {finished} from 'node:stream/promises';
import {readFile, streamWrite, streamRead} from '../../../node.js';

const ROW_COUNT = 5000;

async function* asyncRows(n: number): AsyncGenerator<RowTuple> {
  yield ['id', 'name', 'value'];
  for (let i = 1; i <= n; i++) {
    yield [i, `name-${i}`, i * 10];
  }
}

type RowTuple = (string | number)[];

function* syncRows(n: number): Generator<RowTuple> {
  yield ['id', 'name', 'value'];
  for (let i = 1; i <= n; i++) {
    yield [i, `name-${i}`, i * 10];
  }
}

describe('excel/node streamWrite', () => {
  it('writes multi-thousand rows (declarative + async iterable) and loads back', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-stream-'));
    const path = join(dir, 'big.xlsx');
    try {
      await streamWrite(path, {
        useSharedStrings: true,
        sheets: [
          {
            name: 'Data',
            columns: [
              {header: 'id', key: 'id', width: 10},
              {header: 'name', key: 'name', width: 20},
              {header: 'value', key: 'value', width: 12},
            ],
            // header also comes from columns; use plain array rows without extra header
            rows: (async function* () {
              for (let i = 1; i <= ROW_COUNT; i++) {
                yield [i, `name-${i}`, i * 10] as RowTuple;
              }
            })(),
          },
        ],
      });

      const wb = await readFile(path);
      expect(wb.sheets).toHaveLength(1);
      expect(wb.sheets[0]!.name).toBe('Data');
      // column headers are row 1; data rows follow
      const rows = wb.sheets[0]!.rows;
      expect(rows.length).toBeGreaterThanOrEqual(ROW_COUNT);

      // sample first data-ish row values
      const first = rows.find(r => r.number === 1);
      expect(first).toBeTruthy();
      // headers from columns
      expect(first!.cells[1]?.value).toBe('id');

      const mid = rows.find(r => r.number === 2501);
      expect(mid).toBeTruthy();
      expect(mid!.cells[1]?.value).toBe(2500);
      expect(mid!.cells[2]?.value).toBe('name-2500');
      expect(mid!.cells[3]?.value).toBe(25000);

      const last = rows.find(r => r.number === ROW_COUNT + 1);
      expect(last).toBeTruthy();
      expect(last!.cells[1]?.value).toBe(ROW_COUNT);
      expect(last!.cells[2]?.value).toBe(`name-${ROW_COUNT}`);
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });

  it('supports callback style with for-await rows', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-stream-cb-'));
    const path = join(dir, 'cb.xlsx');
    try {
      await streamWrite(
        path,
        async w => {
          const sheet = w.sheet('S');
          await sheet.rows(asyncRows(200));
        },
        {useSharedStrings: false, useStyles: false},
      );

      const wb = await readFile(path);
      expect(wb.sheets[0]!.name).toBe('S');
      // header + 200 data
      expect(wb.sheets[0]!.rows.length).toBe(201);
      const r2 = wb.sheets[0]!.rows.find(r => r.number === 2);
      expect(r2!.cells[1]?.value).toBe(1);
      expect(r2!.cells[2]?.value).toBe('name-1');
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });

  it('supports sync iterable and Writable stream destination', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-stream-ws-'));
    const path = join(dir, 'writable.xlsx');
    try {
      const out = createWriteStream(path);
      const writePromise = streamWrite(out, {
        sheets: [{name: 'A', rows: syncRows(50)}],
      });
      await writePromise;
      await finished(out);

      const wb = await readFile(path);
      expect(wb.sheets[0]!.rows.length).toBe(51);
      expect(wb.sheets[0]!.rows.find(r => r.number === 51)!.cells[1]?.value).toBe(50);
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });

  it('streamRead yields rows after streamWrite', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-stream-rd-'));
    const path = join(dir, 'round.xlsx');
    try {
      await streamWrite(path, async w => {
        const s = w.sheet('R');
        s.row(['a', 'b']);
        s.row([1, 2]);
        s.row([3, 4]);
      });

      const collected: Array<{sheetName: string; rowNumber: number; a: unknown; b: unknown}> = [];
      for await (const row of streamRead(path)) {
        collected.push({
          sheetName: row.sheetName,
          rowNumber: row.rowNumber,
          a: row.values[1],
          b: row.values[2],
        });
      }

      expect(collected.length).toBeGreaterThanOrEqual(3);
      expect(collected[0]!.sheetName).toBe('R');
      expect(collected.some(r => r.a === 1 && r.b === 2)).toBe(true);
      expect(collected.some(r => r.a === 3 && r.b === 4)).toBe(true);
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });

  it('rejects declarative spec with no sheets', async () => {
    await expect(
      streamWrite(join(tmpdir(), 'nope.xlsx'), {sheets: []} as never),
    ).rejects.toThrow(/at least one sheet/);
  });
});
