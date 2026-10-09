
import {describe} from 'vite-plus/test';
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

import AppXform from '../../../../../lib/xlsx/xform/core/app-xform.js';

const expectations = [
  {
    title: 'app.01',
    create() {
      return new AppXform();
    },
    preparedModel: {worksheets: [{name: 'Sheet1'}]},
    xml: fs
      .readFileSync(`${__dirname}/data/app.01.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render', 'renderIn'],
  },
  {
    title: 'app.02',
    create() {
      return new AppXform();
    },
    preparedModel: {
      worksheets: [{name: 'Sheet1'}, {name: 'Sheet2'}],
      company: 'Cyber Sapiens, Ltd.',
      manager: 'Guyon Roche',
    },
    xml: fs
      .readFileSync(`${__dirname}/data/app.02.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    tests: ['render', 'renderIn'],
  },
];

describe('AppXform', () => {
  testXformHelper(expectations);
});
