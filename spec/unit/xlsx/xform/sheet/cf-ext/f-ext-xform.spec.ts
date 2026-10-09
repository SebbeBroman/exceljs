import FExtXformParser from '../../../../../../lib/xlsx/parser/sheet/cf-ext/f-ext-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../../test-xform-helper.js';

import FExtXform from '../../../../../../lib/xlsx/xform/sheet/cf-ext/f-ext-xform.js';

const expectations = [
  {
    title: 'formula',
    create() {
      return new FExtXform();
    },
    createParser() {
      return new FExtXformParser();
    },
    preparedModel: '7',
    xml: '<xm:f>7</xm:f>',
    parsedModel: '7',
    tests: ['render', 'parse'],
  },
];

describe('FExtXform', () => {
  testXformHelper(expectations);
});
