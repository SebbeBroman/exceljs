import PageMarginsXformParser from '../../../../../lib/xlsx/parser/sheet/page-margins-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import PageMarginsXform from '../../../../../lib/xlsx/xform/sheet/page-margins-xform.js';

const expectations = [
  {
    title: 'normal',
    create() {
      return new PageMarginsXform();
    },
    createParser() {
      return new PageMarginsXformParser();
    },
    preparedModel: {
      left: 0.7,
      right: 0.7,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3,
    },
    xml:
      '<pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>',
    parsedModel: {
      left: 0.7,
      right: 0.7,
      top: 0.75,
      bottom: 0.75,
      header: 0.3,
      footer: 0.3,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('PageMarginsXform', () => {
  testXformHelper(expectations);
});
