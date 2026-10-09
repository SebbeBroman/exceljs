import BlipXformParser from '../../../../../lib/xlsx/parser/drawing/blip-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import BlipXform from '../../../../../lib/xlsx/xform/drawing/blip-xform.js';

const expectations = [
  {
    title: 'full',
    create() {
      return new BlipXform();
    },
    createParser() {
      return new BlipXformParser();
    },
    preparedModel: {rId: 'rId1'},
    xml:
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />',
    parsedModel: {rId: 'rId1'},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('BlipXform', () => {
  testXformHelper(expectations);
});
