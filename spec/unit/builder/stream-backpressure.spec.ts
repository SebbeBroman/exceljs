import {describe, it, expect} from 'vite-plus/test';

import {Writable, PassThrough} from 'node:stream';
import {streamWrite} from '../../../node.js';
import {load} from '../../../excel.js';

function* rows(count: number): Generator<(string | number)[]> {
  for (let i = 0; i < count; i++) yield [i, `row-${i}`, 'abc'.repeat(100)];
}

describe('native streaming backpressure', () => {
  it('bounds destination queue and round-trips two sheets with STORE compression', async () => {
    const chunks: Buffer[] = [];
    let peakQueue = 0;
    const output = new Writable({
      highWaterMark: 1024,
      write(chunk, _encoding, callback) {
        chunks.push(Buffer.from(chunk));
        peakQueue = Math.max(peakQueue, output.writableLength);
        setImmediate(callback);
      },
    });
    await streamWrite(output, {
      zip: {store: true},
      sheets: [
        {name: 'A', rows: rows(4000)},
        {name: 'B', rows: rows(4000)},
      ],
    });
    // At most one XML batch, ZIP headers and metadata can await the slow sink.
    expect(peakQueue).toBeLessThan(100000);
    const workbook = await load(Buffer.concat(chunks));
    expect(workbook.sheets.map(sheet => sheet.rows.length)).toEqual([4000, 4000]);
    expect(workbook.sheets[1]!.rows[3999]!.cells[2]!.value).toBe('row-3999');
    // Duplex destinations finish writing before the caller consumes their readable side.
    const pass = new PassThrough();
    await streamWrite(pass, {sheets: [{name: 'Only', rows: [[1]]}]});
    expect(pass.writableFinished).toBe(true);
    expect((await load(pass.read())).sheets[0]!.rows[0]!.cells[1]!.value).toBe(1);
  });

  it('rejects destination errors while producing rows', async () => {
    const output = new Writable({
      write(_chunk, _encoding, callback) {
        callback(new Error('disk full'));
      },
    });
    await expect(streamWrite(output, {sheets: [{name: 'A', rows: rows(4000)}]})).rejects.toThrow(
      'disk full',
    );
  });

  it('rejects a destination that closes without finishing', async () => {
    const output = new Writable({
      write(_chunk, _encoding, callback) {
        output.destroy();
        callback();
      },
    });
    await expect(streamWrite(output, {sheets: [{name: 'A', rows: rows(4000)}]})).rejects.toThrow();
  });

  it('propagates producer failures and closes the destination', async () => {
    const output = new Writable({
      write(_chunk, _encoding, callback) {
        callback();
      },
    });
    await expect(
      streamWrite(output, async workbook => {
        workbook.sheet('A').row([1]);
        throw new Error('producer failed');
      }),
    ).rejects.toThrow('producer failed');
    expect(output.destroyed).toBe(true);
  });

  it('keeps queued XML intact for callback sheets written out of order', async () => {
    const chunks: Buffer[] = [];
    const output = new Writable({
      write(chunk, _encoding, callback) {
        chunks.push(Buffer.from(chunk));
        callback();
      },
    });
    await streamWrite(
      output,
      workbook => {
        const a = workbook.sheet('A');
        const b = workbook.sheet('B');
        for (let i = 0; i < 2000; i++) {
          b.row([`b-${i}`]);
          a.row([`a-${i}`]);
        }
      },
      {zip: {store: true}},
    );
    const workbook = await load(Buffer.concat(chunks));
    expect(workbook.sheets[1]!.rows[1999]!.cells[1]!.value).toBe('b-1999');
    expect(workbook.sheets[0]!.rows[1999]!.cells[1]!.value).toBe('a-1999');
  });
});
