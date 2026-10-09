import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

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
}

export default WorkbookPivotCacheXform;
export {WorkbookPivotCacheXform};
