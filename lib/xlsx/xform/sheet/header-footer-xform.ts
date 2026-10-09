import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface HeaderFooterModel {
  differentFirst?: boolean;
  differentOddEven?: boolean;
  oddHeader?: string;
  oddFooter?: string;
  evenHeader?: string;
  evenFooter?: string;
  firstHeader?: string;
  firstFooter?: string;
}

export type HeaderFooterNodeName =
  | 'oddHeader'
  | 'oddFooter'
  | 'evenHeader'
  | 'evenFooter'
  | 'firstHeader'
  | 'firstFooter';

class HeaderFooterXform extends BaseXform<HeaderFooterModel> {
  currentNode?: HeaderFooterNodeName;

  override tag = 'headerFooter';

  override render(xmlStream: XmlStreamLike, model?: HeaderFooterModel | null): void {
    if (model) {
      xmlStream.addRollback();

      let createTag = false;

      xmlStream.openNode('headerFooter');
      if (model.differentFirst) {
        xmlStream.addAttribute('differentFirst', '1');
        createTag = true;
      }
      if (model.differentOddEven) {
        xmlStream.addAttribute('differentOddEven', '1');
        createTag = true;
      }
      if (model.oddHeader && typeof model.oddHeader === 'string') {
        xmlStream.leafNode('oddHeader', null as unknown as undefined, model.oddHeader);
        createTag = true;
      }
      if (model.oddFooter && typeof model.oddFooter === 'string') {
        xmlStream.leafNode('oddFooter', null as unknown as undefined, model.oddFooter);
        createTag = true;
      }
      if (model.evenHeader && typeof model.evenHeader === 'string') {
        xmlStream.leafNode('evenHeader', null as unknown as undefined, model.evenHeader);
        createTag = true;
      }
      if (model.evenFooter && typeof model.evenFooter === 'string') {
        xmlStream.leafNode('evenFooter', null as unknown as undefined, model.evenFooter);
        createTag = true;
      }
      if (model.firstHeader && typeof model.firstHeader === 'string') {
        xmlStream.leafNode('firstHeader', null as unknown as undefined, model.firstHeader);
        createTag = true;
      }
      if (model.firstFooter && typeof model.firstFooter === 'string') {
        xmlStream.leafNode('firstFooter', null as unknown as undefined, model.firstFooter);
        createTag = true;
      }

      if (createTag) {
        xmlStream.closeNode();
        xmlStream.commit();
      } else {
        xmlStream.rollback();
      }
    }
  }
}

export default HeaderFooterXform;
export {HeaderFooterXform};
