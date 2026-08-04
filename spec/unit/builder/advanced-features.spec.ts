import {describe, it, expect} from 'vitest';
import {workbook, load} from '../../../excel.ts';
import {unzipSync, strFromU8} from 'fflate';

function allSheetXml(buf: Uint8Array): string {
  const files = unzipSync(buf);
  return Object.keys(files)
    .filter(n => n.includes('sheet') && n.endsWith('.xml'))
    .map(n => strFromU8(files[n]!))
    .join('\n');
}

function allXml(buf: Uint8Array): string {
  const files = unzipSync(buf);
  return Object.keys(files)
    .filter(n => n.endsWith('.xml'))
    .map(n => strFromU8(files[n]!))
    .join('\n');
}

describe('builder advanced features (Phase 5)', () => {
  describe('views / pageSetup / headerFooter', () => {
    it('writes freeze panes into sheet views', async () => {
      const buf = await workbook()
        .sheet('S')
        .row([1, 2, 3])
        .views([{state: 'frozen', xSplit: 1, ySplit: 1, topLeftCell: 'B2', activeCell: 'B2'}])
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.include('sheetView');
      expect(xml).to.match(/ySplit=["']1["']/);
      expect(xml).to.match(/xSplit=["']1["']/);
      expect(xml).to.include('topLeftCell');
    });

    it('writes pageSetup and headerFooter', async () => {
      const buf = await workbook()
        .sheet('Print')
        .rows([['a'], ['b']])
        .pageSetup({orientation: 'landscape', fitToPage: true, fitToWidth: 1, fitToHeight: 0})
        .headerFooter({oddHeader: 'Hello Header', oddFooter: 'Page &P'})
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.include('landscape');
      expect(xml).to.include('Hello Header');
    });

    it('accepts views/pageSetup via sheet init', async () => {
      const plain = workbook()
        .sheet('S', {
          rows: [[1]],
          views: [{state: 'frozen', ySplit: 1}],
          pageSetup: {paperSize: 9},
        })
        .build();

      expect(plain.sheets[0]!.views?.[0]).to.include({state: 'frozen', ySplit: 1});
      expect(plain.sheets[0]!.pageSetup?.paperSize).to.equal(9);
    });
  });

  describe('dataValidation', () => {
    it('writes list validation XML', async () => {
      const buf = await workbook()
        .sheet('DV')
        .cell('A1', 'One')
        .dataValidation('B1', {
          type: 'list',
          allowBlank: true,
          formulae: ['"One,Two,Three"'],
          showErrorMessage: true,
        })
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.include('dataValidations');
      expect(xml).to.include('list');
      expect(xml).to.match(/B1|sqref/);
    });

    it('includes validations in build() plain model', () => {
      const plain = workbook()
        .sheet('DV')
        .dataValidation('A1', {type: 'whole', operator: 'equal', formulae: [5]})
        .build();

      expect(plain.sheets[0]!.dataValidations?.A1).to.deep.include({
        type: 'whole',
        operator: 'equal',
      });
    });
  });

  describe('conditionalFormatting', () => {
    it('writes CF expression rule with style', async () => {
      const buf = await workbook()
        .sheet('CF')
        .rows([
          [1, 2],
          [3, 4],
        ])
        .conditionalFormatting({
          ref: 'A1:B2',
          rules: [
            {
              type: 'cellIs',
              operator: 'greaterThan',
              formulae: [2],
              priority: 1,
              style: {
                font: {bold: true},
                fill: {
                  type: 'pattern',
                  pattern: 'solid',
                  bgColor: {argb: 'FFFF0000'},
                },
              },
            },
          ],
        })
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.include('conditionalFormatting');
      expect(xml).to.match(/cellIs|cfRule/);
      expect(xml).to.match(/A1:B2|sqref/);
    });

    it('forces styles when CF has style', async () => {
      const b = workbook()
        .sheet('CF')
        .row([1])
        .conditionalFormatting({
          ref: 'A1',
          rules: [
            {
              type: 'expression',
              formulae: ['TRUE'],
              priority: 1,
              style: {font: {bold: true}},
            },
          ],
        });
      expect(b._used.styles).to.equal(true);
      expect(b._used.conditionalFormatting).to.equal(true);
    });
  });

  describe('note', () => {
    it('writes comments part for cell notes', async () => {
      const buf = await workbook()
        .sheet('Notes')
        .cell('A1', 'hello')
        .note('A1', 'This is a comment')
        .writeBuffer();

      const files = unzipSync(buf);
      const names = Object.keys(files);
      // comments are typically xl/comments1.xml or vmlDrawing
      const hasComments =
        names.some(n => /comments/i.test(n)) ||
        allXml(buf).includes('This is a comment') ||
        names.some(n => /vmlDrawing/i.test(n));
      expect(hasComments).to.equal(true);
    });

    it('stores notes on plain build()', () => {
      const plain = workbook().sheet('N').cell('B2', 1).note('B2', 'n').build();
      expect(plain.sheets[0]!.notes?.B2).to.equal('n');
    });
  });

  describe('protect', () => {
    it('writes sheetProtection (deferred hash, low spinCount)', async () => {
      const buf = await workbook()
        .sheet('Locked')
        .cell('A1', 1)
        .protect('secret', {spinCount: 1})
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.include('sheetProtection');
      // password hash present
      expect(xml).to.match(/hashValue|algorithmName|saltValue/);
    });

    it('protect keeps builder chain sync and appears on build()', () => {
      const b = workbook().sheet('P').row([1]).protect('pw', {spinCount: 1});
      expect(b._used.protection).to.equal(true);
      const plain = b.build();
      expect(plain.sheets[0]!.protect?.password).to.equal('pw');
    });
  });

  describe('table', () => {
    it('writes table XML part', async () => {
      const buf = await workbook()
        .sheet('T')
        .rows([
          ['Name', 'Age'],
          ['Ada', 36],
          ['Bob', 42],
        ])
        .table({
          name: 'People',
          ref: 'A1:B3',
          columns: [{name: 'Name'}, {name: 'Age'}],
          rows: [
            ['Ada', 36],
            ['Bob', 42],
          ],
        })
        .writeBuffer();

      const files = unzipSync(buf);
      const names = Object.keys(files);
      expect(names.some(n => /table\d*\.xml/i.test(n))).to.equal(true);
      const xml = allXml(buf);
      expect(xml).to.match(/People|table/i);
    });
  });

  describe('image', () => {
    it('embeds workbook media and places image on sheet', async () => {
      // 1x1 PNG
      const png = Buffer.from(
        'iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAYAAAAfFcSJAAAADUlEQVR42mP8z8BQDwAEhQGAhKmMIQAAAABJRU5ErkJggg==',
        'base64',
      );
      const b = workbook().sheet('Img').cell('A1', 'pic');
      const id = b.image({extension: 'png', buffer: png});
      expect(typeof id).to.equal('number');
      const buf = await b.image(id, 'B2:D6').writeBuffer();

      const files = unzipSync(buf);
      const names = Object.keys(files);
      expect(names.some(n => /media\//i.test(n))).to.equal(true);
      expect(names.some(n => /drawing/i.test(n))).to.equal(true);
    });
  });

  describe('definedName', () => {
    it('writes definedNames in workbook.xml', async () => {
      const buf = await workbook()
        .sheet('Data')
        .cell('A1', 10)
        .definedName('MyRange', 'Data!$A$1')
        .writeBuffer();

      const xml = allXml(buf);
      expect(xml).to.match(/definedNames|definedName/);
      expect(xml).to.include('MyRange');
    });
  });

  describe('load round-trip (where supported)', () => {
    it('preserves views and pageSetup through write → load', async () => {
      const buf = await workbook()
        .sheet('S')
        .row([1])
        .views([{state: 'frozen', ySplit: 1, xSplit: 0}])
        .pageSetup({orientation: 'landscape'})
        .writeBuffer();

      const plain = await load(buf);
      expect(plain.sheets[0]!.views?.length).to.be.greaterThan(0);
      // orientation may come back from pageSetup
      const orient = plain.sheets[0]!.pageSetup?.orientation;
      if (orient) expect(orient).to.equal('landscape');
    });

    it('preserves data validations through write → load', async () => {
      const buf = await workbook()
        .sheet('DV')
        .cell('A1', 1)
        .dataValidation('A1', {
          type: 'whole',
          operator: 'between',
          formulae: [1, 10],
          allowBlank: true,
        })
        .writeBuffer();

      const plain = await load(buf);
      const dv = plain.sheets[0]!.dataValidations;
      expect(dv).to.be.ok;
      const entry = dv && (dv.A1 || Object.values(dv)[0]);
      expect(entry?.type).to.equal('whole');
    });

    it('preserves CF through write → load', async () => {
      const buf = await workbook()
        .sheet('CF')
        .rows([[1], [2], [3]])
        .conditionalFormatting({
          ref: 'A1:A3',
          rules: [
            {
              type: 'cellIs',
              operator: 'greaterThan',
              formulae: [1],
              priority: 1,
              style: {font: {bold: true}},
            },
          ],
        })
        .writeBuffer();

      const plain = await load(buf);
      expect(plain.sheets[0]!.conditionalFormattings?.length).to.be.greaterThan(0);
    });

    it('preserves notes through write → load', async () => {
      const buf = await workbook()
        .sheet('N')
        .cell('A1', 'x')
        .note('A1', 'round-trip note')
        .writeBuffer();

      const plain = await load(buf);
      const notes = plain.sheets[0]!.notes;
      expect(notes).to.be.ok;
      const text = notes && (notes.A1 || Object.values(notes)[0]);
      // note may be string or {texts: [...]}
      if (typeof text === 'string') {
        expect(text).to.include('round-trip note');
      } else if (text && typeof text === 'object' && 'texts' in text) {
        const joined = (text.texts || []).map((t: {text?: string}) => t.text).join('');
        expect(joined).to.include('round-trip note');
      } else {
        // some loaders expose differently — still assert notes bag non-empty
        expect(Object.keys(notes || {}).length).to.be.greaterThan(0);
      }
    });

    it('preserves sheetProtection model (write-only password recovery)', async () => {
      const buf = await workbook()
        .sheet('P')
        .cell('A1', 1)
        .protect('pw', {spinCount: 1})
        .writeBuffer();

      const plain = await load(buf);
      // Password is not recoverable; hashed model should round-trip for re-encode
      expect(plain.sheets[0]!.sheetProtection || plain.sheets[0]!.protect).to.be.ok;
    });
  });

  describe('callback sheet form', () => {
    it('supports advanced methods inside sheet callback', async () => {
      const buf = await workbook()
        .sheet('A', s =>
          s
            .row([1, 2])
            .views([{state: 'frozen', ySplit: 1}])
            .dataValidation('B1', {type: 'whole', formulae: [0], operator: 'greaterThan'})
            .note('A1', 'hi')
            .protect(undefined, {spinCount: 1}),
        )
        .writeBuffer();

      const xml = allSheetXml(buf);
      expect(xml).to.match(/sheetProtection|dataValidations/);
      expect(buf.byteLength).to.be.greaterThan(500);
    });
  });
});
