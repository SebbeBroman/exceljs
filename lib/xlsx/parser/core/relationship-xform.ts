import type {RelationshipModel} from '../../xform/core/relationship-xform.js';
export type {RelationshipModel} from '../../xform/core/relationship-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class RelationshipXform extends BaseXform<RelationshipModel> {
  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'Relationship':
        this.model = node.attributes;
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default RelationshipXform;
export {RelationshipXform};
