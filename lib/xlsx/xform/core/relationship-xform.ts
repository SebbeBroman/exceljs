import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export type RelationshipModel = Record<string, string>;

class RelationshipXform extends BaseXform<RelationshipModel> {
  override render(xmlStream: XmlStreamLike, model?: RelationshipModel | null): void {
    xmlStream.leafNode('Relationship', model!);
  }

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
