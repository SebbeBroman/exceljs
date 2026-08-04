// @vitest-migrated
import __req_0 from './data/table.1.1.json';
import __req_1 from './data/table.1.3.json';
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const TableXform = verquire('xlsx/xform/table/table-xform');

const expectations = [
  {
    title: 'showing filter',
    create() {
      return new TableXform();
    },
    initialModel: null,
    preparedModel: __req_0,
    xml: fs.readFileSync(`${__dirname}/data/table.1.2.xml`).toString(),
    parsedModel: __req_1,
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('TableXform', () => {
  testXformHelper(expectations);
});
