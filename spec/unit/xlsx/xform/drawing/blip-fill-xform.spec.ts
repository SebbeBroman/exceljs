import BlipFillXformParser from '../../../../../lib/xlsx/parser/drawing/blip-fill-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import BlipFillXform from '../../../../../lib/xlsx/xform/drawing/blip-fill-xform.js';

const expectations = [
  {
    title: 'normal',
    create() {
      return new BlipFillXform();
    },
    createParser() {
      return new BlipFillXformParser();
    },
    preparedModel: {rId: 'rId1'},
    xml:
      '<xdr:blipFill>' +
      '<a:blip xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" r:embed="rId1" cstate="print" />' +
      '<a:stretch><a:fillRect /></a:stretch>' +
      '</xdr:blipFill>',
    parsedModel: {rId: 'rId1'},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('BlipFillXform', () => {
  testXformHelper(expectations);
});
