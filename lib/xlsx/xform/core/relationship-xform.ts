import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export type RelationshipModel = Record<string, string>;

class RelationshipXform extends BaseXform<RelationshipModel> {
  override render(xmlStream: XmlStreamLike, model?: RelationshipModel | null): void {
    xmlStream.leafNode('Relationship', model!);
  }
}

export default RelationshipXform;
export {RelationshipXform};
