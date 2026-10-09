import DimensionXformParser from '../../../../../lib/xlsx/parser/sheet/dimension-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import DimensionXform from '../../../../../lib/xlsx/xform/sheet/dimension-xform.js';

const expectations = [
  {
    title: 'Dimension',
    create() {
      return new DimensionXform();
    },
    createParser() {
      return new DimensionXformParser();
    },
    preparedModel: 'A1:F5',
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<dimension ref="A1:F5"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('DimensionXform', () => {
  testXformHelper(expectations);
});
