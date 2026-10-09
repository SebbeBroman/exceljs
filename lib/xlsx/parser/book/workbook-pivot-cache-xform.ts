import type {WorkbookPivotCacheModel} from '../../xform/book/workbook-pivot-cache-xform.js';
export type {WorkbookPivotCacheModel} from '../../xform/book/workbook-pivot-cache-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class WorkbookPivotCacheXform extends BaseXform<WorkbookPivotCacheModel> {
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
