import {readFile, mkdtemp, rm} from 'node:fs/promises';
import {createHash} from 'node:crypto';
import {tmpdir} from 'node:os';
import {join} from 'node:path';
import {describe, it, expect} from 'vite-plus/test';
import {unzipSync, strFromU8} from 'fflate';
import {load} from '../../../excel.js';
import {streamWrite, streamRead} from '../../../node.js';
import {fixture} from './data/stream-fixture.mjs';
const projections: Record<string, {hash: string} | {error: string}> = JSON.parse(await readFile(new URL('./data/legacy-projections.json', import.meta.url), 'utf8'));
const streamProjections: Record<string, {hash: string; rows: number} | {error: string}> = JSON.parse(await readFile(new URL('./data/legacy-stream-projections.json', import.meta.url), 'utf8'));

function normalize(value: unknown): string {
  return JSON.stringify(value, (key, item) => key === 'worksheet' ? undefined : item);
}
function parts(bytes: Uint8Array): Record<string, string> {
  return Object.fromEntries(Object.entries(unzipSync(bytes)).map(([key, value]) => [key, strFromU8(value)]));
}

describe('saved pre-removal compatibility fixtures', () => {
  for (const [file, baseline] of Object.entries(projections)) {
    it(`preserves the buffered projection of ${file}`, async () => {
      const bytes = await readFile(new URL(`../../integration/data/${file}`, import.meta.url));
      if ('error' in baseline) {
        await expect(load(bytes)).rejects.toThrow(baseline.error);
      } else {
        const digest = createHash('sha256').update(normalize(await load(bytes))).digest('hex');
        expect(digest).to.equal(baseline.hash);
      }
    });
  }
  for (const [file, baseline] of Object.entries(streamProjections)) {
    it(`preserves the streamed projection of ${file}`, async () => {
      const input = new URL(`../../integration/data/${file}`, import.meta.url);
      const consume = async () => {
        const hash = createHash('sha256'); let rows = 0;
        for await (const row of streamRead(input.pathname)) {hash.update(JSON.stringify(row)); rows++;}
        return {hash: hash.digest('hex'), rows};
      };
      if ('error' in baseline) await expect(consume()).rejects.toThrow(baseline.error);
      else expect(await consume()).to.deep.equal(baseline);
    });
  }
  for (const strings of [false, true]) for (const styles of [false, true]) {
    it(`preserves streaming XML and rows (strings=${strings}, styles=${styles})`, async () => {
      const dir = await mkdtemp(join(tmpdir(), 'stream-parity-'));
      try {
        const output = join(dir, 'actual.xlsx');
        const baselineUrl = new URL(`./data/stream-${strings}-${styles}.xlsx`, import.meta.url);
        await streamWrite(output, fixture(strings, styles));
        expect(parts(await readFile(output))).to.deep.equal(parts(await readFile(baselineUrl)));
        const rows: unknown[] = [];
        for await (const row of streamRead(output)) rows.push(row);
        const baseline = JSON.parse(await readFile(new URL(baselineUrl.href + '.json'), 'utf8'));
        expect(JSON.parse(normalize(rows))).to.deep.equal(baseline);
      } finally {
        await rm(dir, {recursive: true, force: true});
      }
    });
  }  for (const strings of [false, true]) for (const styles of [false, true]) {
    for (const options of [{styles: 'cache'}, {sharedStrings: 'ignore'}, {sharedStrings: 'emit'}, {hyperlinks: 'cache'}, {worksheets: 'ignore'}]) {
      it(`preserves stream read options ${JSON.stringify({strings, styles, options})}`, async () => {
        const input = new URL(`./data/stream-${strings}-${styles}.xlsx`, import.meta.url);
        const suffix = Object.entries(options).flat().join('-');
        const baseline = JSON.parse(await readFile(new URL(input.href + '.' + suffix + '.json'), 'utf8'));
        const consume = async () => {
          const rows: unknown[] = [];
          for await (const row of streamRead(input.pathname, options)) rows.push(row);
          return JSON.parse(normalize(rows));
        };
        if (baseline.error) await expect(consume()).rejects.toThrow(baseline.error);
        else expect(await consume()).to.deep.equal(baseline.rows);
      });
    }
  }

});
