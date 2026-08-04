// @vitest-migrated
import __req_0 from './data/drawing.1.0.js';
import __req_1 from './data/drawing.1.1.js';
import __req_2 from './data/drawing.1.3.js';
import __req_3 from './data/drawing.1.4.js';
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DrawingXform = verquire('xlsx/xform/drawing/drawing-xform');

const options = {
  rels: {
    rId1: {Target: '../media/image1.jpg'},
    rId2: {Target: '../media/image2.jpg'},
  },
  mediaIndex: {image1: 0, image2: 1},
  media: [{}, {}],
};

const expectations = [
  {
    title: 'Drawing 1',
    create() {
      return new DrawingXform({tag: 'xdr:from'});
    },
    initialModel: __req_0,
    preparedModel: __req_1,
    xml: fs.readFileSync(`${__dirname}/data/drawing.1.2.xml`).toString(),
    parsedModel: __req_2,
    reconciledModel: __req_3,
    tests: ['prepare', 'render', 'renderIn', 'parse', 'reconcile'],
    options,
  },
];

describe('DrawingXform', () => {
  testXformHelper(expectations);
});
