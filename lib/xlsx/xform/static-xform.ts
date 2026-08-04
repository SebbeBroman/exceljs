import BaseXform from './base-xform.js';
import type {XmlStreamLike, XmlNode} from './base-xform.js';
import XmlStream from '../../utils/xml-stream.js';

// const model = {
//   tag: 'name',
//   $: {attr: 'value'},
//   c: [
//     { tag: 'child' }
//   ],
//   t: 'some text'
// };

export interface StaticXmlModel {
  tag: string;
  $?: Record<string, unknown>;
  c?: StaticXmlModel[];
  t?: string;
}

function build(xmlStream: XmlStreamLike, model: StaticXmlModel): void {
  xmlStream.openNode(model.tag, model.$);
  if (model.c) {
    model.c.forEach(child => {
      build(xmlStream, child);
    });
  }
  if (model.t) {
    xmlStream.writeText(model.t);
  }
  xmlStream.closeNode();
}

class StaticXform extends BaseXform {
  // Static tree model (separate from BaseXform.parse `model`)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _model: any;
  _xml?: string;

  constructor(model: StaticXmlModel) {
    super();

    // This class is an optimisation for static (unimportant and unchanging) xml
    // It is stateless - apart from its static model and so can be used as a singleton
    // Being stateless - it will only track entry to and exit from it's root xml tag during parsing and nothing else
    // Known issues:
    //    since stateless - parseOpen always returns true. Parent xform must know when to start using this xform
    //    if the root tag is recursive, the parsing will behave unpredictably
    this._model = model;
  }

  override render(xmlStream: XmlStreamLike): void {
    if (!this._xml) {
      const stream = new XmlStream() as unknown as XmlStreamLike;
      build(stream, this._model);
      this._xml = stream.xml;
    }
    xmlStream.writeXml(this._xml);
  }

  override parseOpen(): boolean {
    return true;
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    switch (name) {
      case this._model.tag:
        return false;
      default:
        return true;
    }
  }
}

export default StaticXform;
export {StaticXform};
