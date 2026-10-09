import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import RelationshipXform from './relationship-xform.js';
import type {RelationshipModel} from './relationship-xform.js';

class RelationshipsXform extends BaseXform<RelationshipModel[]> {
  declare map: {Relationship: RelationshipXform};
  _values?: RelationshipModel[];

  constructor() {
    super();

    this.map = {
      Relationship: new RelationshipXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: RelationshipModel[] | null): void {
    model = model || this._values;
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);
    xmlStream.openNode('Relationships', RelationshipsXform.RELATIONSHIPS_ATTRIBUTES);

    model!.forEach(relationship => {
      this.map.Relationship.render(xmlStream, relationship);
    });

    xmlStream.closeNode();
  }

  static RELATIONSHIPS_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/package/2006/relationships',
  };
}

export default RelationshipsXform;
export {RelationshipsXform};
