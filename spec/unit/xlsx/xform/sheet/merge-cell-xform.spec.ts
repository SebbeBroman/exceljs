import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import MergeCellXform from '../../../../../lib/xlsx/xform/sheet/merge-cell-xform.js';

const expectations = [
  {
    title: 'Merge',
    create() {
      return new MergeCellXform();
    },
    preparedModel: 'B2:C4',
    xml: '<mergeCell ref="B2:C4"/>',
    parsedModel: 'B2:C4',
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('MergeCellXform', () => {
  testXformHelper(expectations);
});
