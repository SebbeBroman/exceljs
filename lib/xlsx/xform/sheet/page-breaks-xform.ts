import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (node.name === 'brk') {
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

export default PageBreaksXform;
export {PageBreaksXform};
