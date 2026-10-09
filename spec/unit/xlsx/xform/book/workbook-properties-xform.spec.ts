import WorkbookPropertiesXformParser from '../../../../../lib/xlsx/parser/book/workbook-properties-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import WorkbookPropertiesXform from '../../../../../lib/xlsx/xform/book/workbook-properties-xform.js';

const expectations = [
  {
    title: 'default',
    create() {
      return new WorkbookPropertiesXform();
    },
    preparedModel: {},
    xml:
      '<workbookPr defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
    parsedModel: {},
    tests: ['render', 'renderIn'],
  },
  {
    title: 'date1904',
    create() {
      return new WorkbookPropertiesXform();
    },
    createParser() {
      return new WorkbookPropertiesXformParser();
    },
    preparedModel: {date1904: true},
    xml:
      '<workbookPr date1904="1" defaultThemeVersion="164011" filterPrivacy="1"></workbookPr>',
    parsedModel: {date1904: true},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('WorkbookPropertiesXform', () => {
  testXformHelper(expectations);
});
