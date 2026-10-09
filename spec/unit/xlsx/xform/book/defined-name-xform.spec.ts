import DefinedNameXformParser from '../../../../../lib/xlsx/parser/book/defined-name-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import DefinedNameXform from '../../../../../lib/xlsx/xform/book/defined-name-xform.js';

const expectations = [
  {
    title: 'Defined Names',
    create() {
      return new DefinedNameXform();
    },
    createParser() {
      return new DefinedNameXformParser();
    },
    preparedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1']},
    xml: '<definedName name="foo">bar!$A$1:$C$1</definedName>',
    parsedModel: {name: 'foo', ranges: ['bar!$A$1:$C$1']},
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'Print Area',
    create() {
      return new DefinedNameXform();
    },
    createParser() {
      return new DefinedNameXformParser();
    },
    preparedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    xml:
      '<definedName name="_xlnm.Print_Area" localSheetId="0">bar!$A$1:$C$10</definedName>',
    parsedModel: {
      name: '_xlnm.Print_Area',
      localSheetId: 0,
      ranges: ['bar!$A$1:$C$10'],
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'String with something that looks like a range',
    create() {
      return new DefinedNameXform();
    },
    createParser() {
      return new DefinedNameXformParser();
    },
    preparedModel: {name: 'foo', ranges: []},
    xml: '<definedName name="foo">"OFFSET($A$10;0;0;0;1)"</definedName>',
    parsedModel: {name: 'foo', ranges: []},
    tests: ['parse'],
  },
];

describe('DefinedNameXform', () => {
  testXformHelper(expectations);
});
