import StringXformParser from '../../../../../lib/xlsx/parser/simple/string-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import StringXform from '../../../../../lib/xlsx/xform/simple/string-xform.js';

const expectations = [
  {
    title: 'hello',
    create() {
      return new StringXform({tag: 'string', attr: 'val'});
    },
    createParser() {
      return new StringXformParser({tag: 'string', attr: 'val'});
    },
    preparedModel: 'Hello, World!',
    xml: '<string val="Hello, World!"/>',
    parsedModel: 'Hello, World!',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'empty',
    create() {
      return new StringXform({tag: 'string', attr: 'val'});
    },
    createParser() {
      return new StringXformParser({tag: 'string', attr: 'val'});
    },
    preparedModel: '',
    xml: '<string val=""/>',
    parsedModel: '',
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'undefined',
    create() {
      return new StringXform({tag: 'string', attr: 'val'});
    },
    preparedModel: undefined,
    xml: '',
    tests: ['render', 'renderIn'],
  },
];

describe('StringXform', () => {
  testXformHelper(expectations);
});
