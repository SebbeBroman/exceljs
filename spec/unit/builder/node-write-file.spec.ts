import {describe, it, expect} from 'vite-plus/test';

import {mkdtemp, readFile, rm} from 'node:fs/promises';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {workbook, writeFile} from '../../../node.js';
import {unzipSync, strFromU8} from 'fflate';

describe('excel/node writeFile', () => {
  it('writes an xlsx file to disk', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'excel-ts-'));
    const path = join(dir, 'out.xlsx');
    try {
      await writeFile(
        path,
        workbook({creator: 'node-test'}).sheet('S').rows([
          ['a', 1],
          ['b', 2],
        ]),
      );
      const bytes = new Uint8Array(await readFile(path));
      expect(bytes.byteLength).toBeGreaterThan(500);
      const files = unzipSync(bytes);
      const sheetName = Object.keys(files).find(n => n.includes('worksheets/sheet'));
      expect(sheetName).toBeTruthy();
      expect(strFromU8(files[sheetName!]!)).toContain('a');
    } finally {
      await rm(dir, {recursive: true, force: true});
    }
  });
});
