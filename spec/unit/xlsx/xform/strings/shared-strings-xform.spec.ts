// @vitest-migrated
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import __json_0 from './data/sharedStrings.json';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const SharedStringsXform = verquire('xlsx/xform/strings/shared-strings-xform');

const expectations = [
  {
    title: 'Shared Strings',
    create() {
      return new SharedStringsXform();
    },
    preparedModel: __json_0,
    xml: fs.readFileSync(`${__dirname}/data/sharedStrings.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('SharedStringsXform', () => {
  testXformHelper(expectations);
});
