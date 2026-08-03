// @vitest-migrated
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import __json_0 from './data/sheet.1.0.json';
import __json_1 from './data/sheet.1.1.json';
import __json_2 from './data/sheet.1.3.json';
import __json_3 from './data/sheet.1.4.json';
import __json_4 from './data/sheet.2.0.json';
import __json_5 from './data/sheet.2.1.json';
import __json_6 from './data/sheet.3.1.json';
import __json_7 from './data/sheet.5.0.json';
import __json_8 from './data/sheet.5.1.json';
import __json_9 from './data/sheet.5.3.json';
import __json_10 from './data/sheet.5.4.json';
import __json_11 from './data/sheet.6.1.json';
import __json_12 from './data/sheet.6.3.json';
import __json_13 from './data/sheet.7.0.json';
import __json_14 from './data/sheet.7.1.json';
import __json_15 from './data/sheet.4.0.json';
import __json_16 from './data/sheet.4.0.json';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const Enums = verquire('doc/enums');
const XmlStream = verquire('utils/xml-stream');
const WorksheetXform = verquire('xlsx/xform/sheet/worksheet-xform');

const SharedStringsXform = verquire('xlsx/xform/strings/shared-strings-xform');
const StylesXform = verquire('xlsx/xform/style/styles-xform');

const fakeStyles = {
  addStyleModel(style, cellType) {
    if (cellType === Enums.ValueType.Date) {
      return 1;
    }
    if (style && style.font) {
      return 2;
    }
    return 0;
  },
  getStyleModel(id) {
    switch (id) {
      case 1:
        return {numFmt: 'mm-dd-yy'};
      case 2:
        return {
          font: {
            underline: true,
            size: 11,
            color: {theme: 10},
            name: 'Calibri',
            family: 2,
            scheme: 'minor',
          },
        };
      default:
        return null;
    }
  },
};

const fakeHyperlinkMap = {
  B6: 'https://www.npmjs.com/package/exceljs',
};

function fixDate(model) {
  model.rows[3].cells[1].value = new Date(model.rows[3].cells[1].value);
  return model;
}

const expectations = [
  {
    title: 'Sheet 1',
    create: () => new WorksheetXform(),
    initialModel: fixDate(__json_0),
    preparedModel: fixDate(__json_1),
    xml: fs.readFileSync(`${__dirname}/data/sheet.1.2.xml`).toString(),
    parsedModel: __json_2,
    reconciledModel: fixDate(__json_3),
    tests: ['prepare', 'render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 2 - Data Validations',
    create: () => new WorksheetXform(),
    initialModel: __json_4,
    preparedModel: __json_5,
    xml: fs.readFileSync(`${__dirname}/data/sheet.2.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 3 - Empty Sheet',
    create: () => new WorksheetXform(),
    preparedModel: __json_6,
    xml: fs.readFileSync(`${__dirname}/data/sheet.3.2.xml`).toString(),
    tests: ['render'],
    options: {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    },
  },
  {
    title: 'Sheet 5 - Shared Formulas',
    create: () => new WorksheetXform(),
    initialModel: __json_7,
    preparedModel: __json_8,
    xml: fs.readFileSync(`${__dirname}/data/sheet.5.2.xml`).toString(),
    parsedModel: __json_9,
    reconciledModel: __json_10,
    tests: ['prepare-render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 6 - AutoFilter',
    create: () => new WorksheetXform(),
    preparedModel: __json_11,
    xml: fs.readFileSync(`${__dirname}/data/sheet.6.2.xml`).toString(),
    parsedModel: __json_12,
    tests: ['render', 'parse'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
  {
    title: 'Sheet 7 - Row Breaks',
    create: () => new WorksheetXform(),
    initialModel: __json_13,
    preparedModel: __json_14,
    xml: fs.readFileSync(`${__dirname}/data/sheet.7.2.xml`).toString(),
    tests: ['prepare', 'render'],
    options: {
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
      hyperlinkMap: fakeHyperlinkMap,
      styles: fakeStyles,
      formulae: {},
      siFormulae: 0,
    },
  },
];

describe('WorksheetXform', () => {
  testXformHelper(expectations);

  it('hyperlinks must be after dataValidations', () => {
    const xform = new WorksheetXform();
    const model = __json_15;
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      sharedStrings: new SharedStringsXform(),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const {xml} = xmlStream;
    const iHyperlinks = xml.indexOf('hyperlinks');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iHyperlinks).not.to.equal(-1);
    expect(iDataValidations).not.to.equal(-1);
    expect(iHyperlinks).to.be.greaterThan(iDataValidations);
  });

  it('conditionalFormattings must be before dataValidations', async () => {
    const xform = new WorksheetXform();
    await xform.installCfXforms();
    const model = __json_16;
    const xmlStream = new XmlStream();
    const options = {
      styles: new StylesXform(true),
      hyperlinks: [],
    };
    xform.prepare(model, options);
    xform.render(xmlStream, model);

    const {xml} = xmlStream;
    const iConditionalFormatting = xml.indexOf('conditionalFormatting');
    const iDataValidations = xml.indexOf('dataValidations');
    expect(iConditionalFormatting).not.to.equal(-1);
    expect(iDataValidations).not.to.equal(-1);
    expect(iConditionalFormatting).to.be.lessThan(iDataValidations);
  });
});
