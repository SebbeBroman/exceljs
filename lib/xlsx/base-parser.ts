import XformState from './xform/xform-state.js';
import type {XmlNode, SaxParser} from './xform/xform-state.js';
export type * from './xform/xform-state.js';
let parseSaxPromise: Promise<typeof import('../utils/parse-sax.js').default> | undefined;
function loadParseSax(): Promise<typeof import('../utils/parse-sax.js').default> {
  if (!parseSaxPromise) {
    parseSaxPromise = import('../utils/parse-sax.js').then(m => m.default);
  }
  return parseSaxPromise;
}

class BaseXform<TModel = unknown> extends XformState<TModel> {
  async parse(saxParser: SaxParser): Promise<TModel | null | undefined> {
    for await (const events of saxParser) {
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          this.parseOpen(value as XmlNode);
        } else if (eventType === 'text') {
          this.parseText(value as string);
        } else if (eventType === 'closetag') {
          if (!this.parseClose((value as {name: string}).name)) {
            return this.model;
          }
        }
      }
    }
    return this.model;
  }
  async parseStream(stream: any): Promise<TModel | null | undefined> {
    const parseSax = await loadParseSax();
    return this.parse(parseSax(stream) as SaxParser);
  }
}
export default BaseXform;
export {BaseXform};
