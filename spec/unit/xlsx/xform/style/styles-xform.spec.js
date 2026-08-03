// @vitest-migrated
import fs from 'node:fs';
import testXformHelper from '../test-xform-helper.js';
import __json_0 from './data/styles.1.1.json';
import path from 'node:path';
import {fileURLToPath} from 'node:url';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const StylesXform = verquire('xlsx/xform/style/styles-xform');
const XmlStream = verquire('utils/xml-stream');

const expectations = [
  {
    title: 'Styles with fonts',
    create() {
      return new StylesXform();
    },
    preparedModel: __json_0,
    xml: fs.readFileSync(`${__dirname}/data/styles.1.2.xml`).toString(),
    get parsedModel() {
      return this.preparedModel;
    },
    tests: ['render', 'renderIn', 'parse'],
  },
];

describe('StylesXform', () => {
  testXformHelper(expectations);

  describe('As StyleManager', () => {
    it('Renders empty model', () => {
      const stylesXform = new StylesXform(true);
      const expectedXml = fs
        .readFileSync(`${__dirname}/data/styles.2.2.xml`)
        .toString();

      const xmlStream = new XmlStream();
      stylesXform.render(xmlStream);

      expect(xmlStream.xml).xml.to.equal(expectedXml);
    });
  });
});
