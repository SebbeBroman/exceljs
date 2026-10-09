import PageSetupPropertiesXformParser from '../../../../../lib/xlsx/parser/sheet/page-setup-properties-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import PageSetupPropertiesXform from '../../../../../lib/xlsx/xform/sheet/page-setup-properties-xform.js';

const expectations = [
  {
    title: 'fitToPage',
    create() {
      return new PageSetupPropertiesXform();
    },
    createParser() {
      return new PageSetupPropertiesXformParser();
    },
    preparedModel: {fitToPage: true},
    xml: '<pageSetUpPr fitToPage="1"/>',
    parsedModel: {fitToPage: true},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('PageSetupPropertiesXform', () => {
  testXformHelper(expectations);
});
