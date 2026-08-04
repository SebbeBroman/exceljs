// @vitest-migrated
import testXformHelper from '../../test-xform-helper.js';

const FExtXform = verquire('xlsx/xform/sheet/cf-ext/f-ext-xform');

const expectations = [
  {
    title: 'formula',
    create() {
      return new FExtXform();
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
