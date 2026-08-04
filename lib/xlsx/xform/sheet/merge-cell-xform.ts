import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class MergeCellXform extends BaseXform<string> {
  override tag = 'mergeCell';

  override render(xmlStream: XmlStreamLike, model?: string | null): void {
    xmlStream.leafNode('mergeCell', {ref: model as string});
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'mergeCell') {
      this.model = node.attributes.ref;
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default MergeCellXform;
export {MergeCellXform};
