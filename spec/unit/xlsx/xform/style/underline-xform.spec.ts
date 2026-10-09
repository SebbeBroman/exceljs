import UnderlineXformParser from '../../../../../lib/xlsx/parser/style/underline-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import UnderlineXform from '../../../../../lib/xlsx/xform/style/underline-xform.js';

const expectations = [
  {
    title: 'single',
    create() {
      return new UnderlineXform();
    },
    createParser() {
      return new UnderlineXformParser();
    },
    preparedModel: true,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<u/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'double',
    create() {
      return new UnderlineXform();
    },
    createParser() {
      return new UnderlineXformParser();
    },
    preparedModel: 'double',
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '<u val="double"/>',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'false',
    create() {
      return new UnderlineXform();
    },
    preparedModel: false,
    get parsedModel() {
      return this.preparedModel;
    },
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('UnderlineXform', () => {
  testXformHelper(expectations);
});
