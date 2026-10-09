import XmlStream from '../../utils/xml-stream.js';
import XformState from './xform-state.js';
import type {XmlStreamLike} from './xform-state.js';
export type * from './xform-state.js';

class BaseXform<TModel = unknown> extends XformState<TModel> {
  get xml(): string {
    // convenience function to get the xml of this.model
    // useful for manager types that are built during the prepare phase
    return this.toXml(this.model);
  }
  toXml(model?: any): string {
    const xmlStream = new XmlStream() as unknown as XmlStreamLike;
    this.render(xmlStream, model);
    return xmlStream.xml;
  }
}
export default BaseXform;
export {BaseXform};
