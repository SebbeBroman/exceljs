import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class AppHeadingPairsXform extends BaseXform<unknown[]> {
  override render(xmlStream: XmlStreamLike, model?: unknown[] | null): void {
    xmlStream.openNode('HeadingPairs');
    xmlStream.openNode('vt:vector', {size: 2, baseType: 'variant'});

    xmlStream.openNode('vt:variant');
    xmlStream.leafNode('vt:lpstr', undefined, 'Worksheets');
    xmlStream.closeNode();

    xmlStream.openNode('vt:variant');
    xmlStream.leafNode('vt:i4', undefined, model!.length);
    xmlStream.closeNode();

    xmlStream.closeNode();
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    // no parsing
    return node.name === 'HeadingPairs';
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    return name !== 'HeadingPairs';
  }
}

export default AppHeadingPairsXform;
export {AppHeadingPairsXform};
