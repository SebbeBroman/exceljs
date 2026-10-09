import WorkbookXformParser from '../../../../../lib/xlsx/parser/book/workbook-xform.js';
import {describe} from 'vite-plus/test';
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import __json_0 from './data/book.1.1.json' with {type: 'json'};
import __json_1 from './data/book.1.3.json' with {type: 'json'};
import __json_2 from './data/book.2.3.json' with {type: 'json'};
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

import WorkbookXform from '../../../../../lib/xlsx/xform/book/workbook-xform.js';

const expectations = [
  {
    title: 'book.1',
    create() {
      return new WorkbookXform();
    },
    createParser() {
      return new WorkbookXformParser();
    },
    preparedModel: __json_0,
    xml: fs
      .readFileSync(`${__dirname}/data/book.1.2.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    parsedModel: __json_1,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'book.2 - no properties',
    create() {
      return new WorkbookXform();
    },
    createParser() {
      return new WorkbookXformParser();
    },
    xml: fs
      .readFileSync(`${__dirname}/data/book.2.2.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    parsedModel: __json_2,
    tests: ['parse'],
  },
];

describe('WorkbookXform', () => {
  testXformHelper(expectations);
});
