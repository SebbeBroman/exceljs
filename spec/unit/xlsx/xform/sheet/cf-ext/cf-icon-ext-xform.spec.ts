import {describe} from 'vite-plus/test';
import testXformHelper from '../../test-xform-helper.js';

import CfIconExtXform from '../../../../../../lib/xlsx/xform/sheet/cf-ext/cf-icon-ext-xform.js';

const expectations = [
  {
    title: 'range',
    create() {
      return new CfIconExtXform();
    },
    preparedModel: {iconSet: '3Triangles', iconId: 7},
    xml: '<x14:cfIcon iconSet="3Triangles" iconId="7" />',
    parsedModel: {iconSet: '3Triangles', iconId: 7},
    tests: ['render', 'parse'],
  },
];

describe('CfIconExtXform', () => {
  testXformHelper(expectations);
});
