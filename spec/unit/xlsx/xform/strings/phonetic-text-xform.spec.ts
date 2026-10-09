import PhoneticTextXformParser from '../../../../../lib/xlsx/parser/strings/phonetic-text-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import PhoneticTextXform from '../../../../../lib/xlsx/xform/strings/phonetic-text-xform.js';

const expectations = [
  {
    title: 'text',
    create() {
      return new PhoneticTextXform();
    },
    createParser() {
      return new PhoneticTextXformParser();
    },
    preparedModel: {text: 'Hello, World!', sb: 0, eb: 1},
    xml: '<rPh sb="0" eb="1"><t>Hello, World!</t></rPh>',
    parsedModel: {text: 'Hello, World!', sb: 0, eb: 1},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Katakana',
    create() {
      return new PhoneticTextXform();
    },
    createParser() {
      return new PhoneticTextXformParser();
    },
    preparedModel: {sb: 0, eb: 2, text: 'ヤクワリ'},
    xml: '<rPh sb="0" eb="2"><t>ヤクワリ</t></rPh>',
    parsedModel: {sb: 0, eb: 2, text: 'ヤクワリ'},
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('PhoneticTextXform', () => {
  testXformHelper(expectations);
});
