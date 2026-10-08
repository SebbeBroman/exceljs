import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import WorkbookCalcPropertiesXform from '../../../../../lib/xlsx/xform/book/workbook-calc-properties-xform.js';

const expectations = [
  {
    title: 'default',
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: {},
    xml: '<calcPr calcId="171027"></calcPr>',
    parsedModel: {},
    tests: ['render', 'renderIn'],
  },
  {
    title: 'fullCalcOnLoad',
    create() {
      return new WorkbookCalcPropertiesXform();
    },
    preparedModel: {fullCalcOnLoad: true},
    xml: '<calcPr calcId="171027" fullCalcOnLoad="1"></calcPr>',
    parsedModel: {},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('WorkbookCalcPropertiesXform', () => {
  testXformHelper(expectations);
});
