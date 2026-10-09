import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface PageBreakModel {
  id?: number;
  max?: number;
  min?: number;
  man?: number;
  ref?: string;
  [key: string]: unknown;
}

class PageBreaksXform extends BaseXform<PageBreakModel | string> {
  override tag = 'brk';

  override render(xmlStream: XmlStreamLike, model?: PageBreakModel | string | null): void {
    xmlStream.leafNode('brk', model as unknown as Record<string, unknown>);
  }
}

export default PageBreaksXform;
export {PageBreaksXform};
