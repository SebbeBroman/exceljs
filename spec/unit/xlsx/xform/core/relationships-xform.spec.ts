import RelationshipsXformParser from '../../../../../lib/xlsx/parser/core/relationships-xform.js';
import {describe} from 'vite-plus/test';
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import __json_0 from './data/worksheet.rels.1.json' with {type: 'json'};
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

import RelationshipsXform from '../../../../../lib/xlsx/xform/core/relationships-xform.js';

const expectations = [
  {
    title: 'worksheet.rels',
    create() {
      return new RelationshipsXform();
    },
    createParser() {
      return new RelationshipsXformParser();
    },
    preparedModel: __json_0,
    xml: fs
      .readFileSync(`${__dirname}/data/worksheet.rels.xml`)
      .toString()
      .replace(/\r\n/g, '\n'),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('RelationshipsXform', () => {
  testXformHelper(expectations);
});
