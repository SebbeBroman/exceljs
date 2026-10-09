import PrintOptionsXformParser from '../../../../../lib/xlsx/parser/sheet/print-options-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import PrintOptionsXform from '../../../../../lib/xlsx/xform/sheet/print-options-xform.js';

const expectations = [
  {
    title: 'empty',
    create() {
      return new PrintOptionsXform();
    },
    preparedModel: {},
    xml: '',
    parsedModel: {},
    tests: ['render', 'renderIn'],
  },
  {
    title: 'gridlines',
    create() {
      return new PrintOptionsXform();
    },
    createParser() {
      return new PrintOptionsXformParser();
    },
    preparedModel: {showGridLines: true},
    xml: '<printOptions gridLines="1"/>',
    parsedModel: {
      showGridLines: true,
      showRowColHeaders: false,
      horizontalCentered: false,
      verticalCentered: false,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'headers',
    create() {
      return new PrintOptionsXform();
    },
    createParser() {
      return new PrintOptionsXformParser();
    },
    preparedModel: {showRowColHeaders: true},
    xml: '<printOptions headings="1"/>',
    parsedModel: {
      showGridLines: false,
      showRowColHeaders: true,
      horizontalCentered: false,
      verticalCentered: false,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'centered',
    create() {
      return new PrintOptionsXform();
    },
    createParser() {
      return new PrintOptionsXformParser();
    },
    preparedModel: {horizontalCentered: true, verticalCentered: true},
    xml: '<printOptions horizontalCentered="1" verticalCentered="1"/>',
    parsedModel: {
      showGridLines: false,
      showRowColHeaders: false,
      horizontalCentered: true,
      verticalCentered: true,
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('PrintOptionsXform', () => {
  testXformHelper(expectations);
});
