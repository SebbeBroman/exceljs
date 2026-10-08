import {describe, it, expect} from 'vite-plus/test';

import DefinedNames from '../../../lib/model/defined-names.js';
import Range from '../../../lib/model/range.js';

describe('compact defined names', () => {
  it('coalesces adjacent cells and overlapping ranges, preserving quoted sheet names', () => {
    const names = new DefinedNames();
    for (const cell of ['A1', 'A2', 'B1', 'B2']) names.add(`'blo ,!rt'!${cell}`, 'Square');
    names.add("'blo ,!rt'!$A$1:$B$2", 'Square');
    names.add('Other!$C$4', 'Square');
    expect(names.model).toEqual([{name: 'Square', ranges: ["'blo ,!rt'!$A$1:$B$2", 'Other!$C$4']}]);
  });
  it('retains the whole Excel grid as one reference', () => {
    const names = new DefinedNames();
    names.add('Data!A1:XFD1048576', 'Everything');
    expect(names.model).toEqual([{name: 'Everything', ranges: ['Data!$A$1:$XFD$1048576']}]);
  });
  it('skips invalid loaded references and resets prior state', () => {
    const names = new DefinedNames();
    names.add('Old!A1', 'Old');
    names.model = [
      {name: 'invalid', ranges: ['"="', '#REF!']},
      {name: 'valid', ranges: ['Sheet3!$A$2:$F$2228']},
    ];
    expect(names.model).toEqual([{name: 'valid', ranges: ['Sheet3!$A$2:$F$2228']}]);
  });
  it('preserves the union of deterministic overlapping rectangles without duplicate coverage', () => {
    let seed = 17;
    const random = () => ((seed = (Math.imul(seed, 1664525) + 1013904223) >>> 0) % 12) + 1;
    for (let sample = 0; sample < 25; sample++) {
      const names = new DefinedNames();
      const expected = new Set<string>();
      for (let n = 0; n < 15; n++) {
        const range = new Range(random(), random(), random(), random(), 'S');
        names.add(range.$shortRange, 'Union');
        range.forEachAddress(cell => expected.add(cell));
      }
      const actual = new Set<string>();
      let count = 0;
      for (const ref of names.model[0].ranges)
        new Range(ref).forEachAddress(cell => {
          actual.add(cell);
          count++;
        });
      expect([...actual].sort()).toEqual([...expected].sort());
      expect(count).toBe(actual.size);
    }
  });
});
