import BaseXform from './base-xform.js';
import type {XmlStreamLike} from './base-xform.js';
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
  // Static tree model (separate from the transform model)
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _model: any;
  _xml?: string;

  constructor(model: StaticXmlModel) {
    super();

    // Cache the serialization of an unchanging XML tree.
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
}

export default StaticXform;
export {StaticXform};
