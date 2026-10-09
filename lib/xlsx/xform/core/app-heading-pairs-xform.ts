import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default AppHeadingPairsXform;
export {AppHeadingPairsXform};
