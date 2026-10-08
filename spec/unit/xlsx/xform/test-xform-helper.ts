import {describe, it, expect} from 'vite-plus/test';
import {PassThrough} from 'node:stream';
import CompyXform from './compy-xform.js';
import {cloneDeep} from '../../../utils/clone-deep.js';
import {normalizeXml} from '../../../utils/normalize-xml.js';
import parseSax from '../../../../lib/utils/parse-sax.js';
import XmlStream from '../../../../lib/utils/xml-stream.js';
import BooleanXform from '../../../../lib/xlsx/xform/simple/boolean-xform.js';
import type BaseXform from '../../../../lib/xlsx/xform/base-xform.js';

interface Expectation {
  title: string;
  create(): BaseXform;
  tests: string[];
  options?: Record<string, unknown>;
  [key: string]: unknown;
}

function getExpectation(expectation: Expectation, name: string) {
  if (!Object.hasOwn(expectation, name))
    throw new Error(`Expectation missing required field: ${name}`);
  return cloneDeep(expectation[name]);
}

function composite(expectation: Expectation): CompyXform {
  const child = expectation.create();
  return new CompyXform({
    tag: 'compy',
    children: [
      {name: 'pre', xform: new BooleanXform({tag: 'pre', attr: 'val'})},
      {name: child.tag, xform: child},
      {name: 'post', xform: new BooleanXform({tag: 'post', attr: 'val'})},
    ],
  });
}

async function parse(xform: BaseXform, xml: string) {
  const stream = new PassThrough();
  stream.end(xml);
  return xform.parse(parseSax(stream));
}

const its = {
  prepare(e: Expectation) {
    it('Prepare Model', () => {
      const model = getExpectation(e, 'initialModel');
      e.create().prepare(model, e.options);
      expect(cloneDeep(model, false)).toEqual(getExpectation(e, 'preparedModel'));
    });
  },
  render(e: Expectation) {
    it('Render to XML', () => {
      const xml = new XmlStream();
      e.create().render(xml, getExpectation(e, 'preparedModel'), 0);
      expect(normalizeXml(xml.xml)).toBe(normalizeXml(getExpectation(e, 'xml') as string));
    });
  },
  'prepare-render'(e: Expectation) {
    it('Prepare and Render to XML', () => {
      const model = getExpectation(e, 'initialModel');
      const xform = e.create();
      xform.prepare(model, e.options);
      expect(normalizeXml(xform.toXml(model))).toBe(
        normalizeXml(getExpectation(e, 'xml') as string),
      );
    });
  },
  renderIn(e: Expectation) {
    it('Render in Composite to XML', () => {
      const model = {pre: true, child: getExpectation(e, 'preparedModel'), post: true};
      const child = e.create();
      const xform = composite({...e, create: () => child});
      // Render fixtures use the model key "child" independently of the XML tag.
      xform.map[child.tag!].name = 'child';
      expect(normalizeXml(xform.toXml(model))).toBe(
        normalizeXml(`<compy><pre/>${getExpectation(e, 'xml')}<post/></compy>`),
      );
    });
  },
  parseIn(e: Expectation) {
    it('Parse within composite', async () => {
      const xform = composite(e);
      const child = xform.map[Object.keys(xform.map)[1]!];
      const model = await parse(xform, `<compy><pre/>${getExpectation(e, 'xml')}<post/></compy>`);
      expect(cloneDeep(model, false)).toEqual({
        pre: true,
        [child.tag!]: getExpectation(e, 'parsedModel'),
        post: true,
      });
    });
  },
  parse(e: Expectation) {
    it('Parse to Model', async () => {
      const model = await parse(e.create(), getExpectation(e, 'xml') as string);
      expect(cloneDeep(model, false)).toEqual(getExpectation(e, 'parsedModel'));
    });
  },
  reconcile(e: Expectation) {
    it('Reconcile Model', () => {
      const model = getExpectation(e, 'parsedModel');
      e.create().reconcile(model, e.options);
      expect(cloneDeep(model, false)).toEqual(getExpectation(e, 'reconciledModel'));
    });
  },
};

export default function testXform(expectations: Expectation[]): void {
  for (const expectation of expectations)
    describe(expectation.title, () => {
      for (const test of expectation.tests) {
        if (!Object.hasOwn(its, test)) throw new Error(`Unknown transform test: ${test}`);
        its[test as keyof typeof its](expectation);
      }
    });
}
