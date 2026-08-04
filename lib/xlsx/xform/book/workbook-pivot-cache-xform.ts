import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface WorkbookPivotCacheModel {
  cacheId: string;
  rId: string;
}

class WorkbookPivotCacheXform extends BaseXform<WorkbookPivotCacheModel> {
  override render(xmlStream: XmlStreamLike, model?: WorkbookPivotCacheModel | null): void {
    xmlStream.leafNode('pivotCache', {
      cacheId: model!.cacheId,
      'r:id': model!.rId,
    });
  }

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'pivotCache') {
      this.model = {
        cacheId: node.attributes.cacheId,
        rId: node.attributes['r:id'],
      };
      return true;
    }
    return false;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default WorkbookPivotCacheXform;
export {WorkbookPivotCacheXform};
