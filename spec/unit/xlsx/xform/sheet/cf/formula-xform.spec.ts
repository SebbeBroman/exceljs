import {describe} from 'vite-plus/test';
import testXformHelper from '../../test-xform-helper.js';

import FormulaXform from '../../../../../../lib/xlsx/xform/sheet/cf/formula-xform.js';

const expectations = [
  {
    title: 'formula',
    create() {
      return new FormulaXform();
    },
    preparedModel: 'ROW()',
    xml: '<formula>ROW()</formula>',
    parsedModel: 'ROW()',
    tests: ['render', 'parse'],
  },
];

describe('FormulaXform', () => {
  testXformHelper(expectations);
});
