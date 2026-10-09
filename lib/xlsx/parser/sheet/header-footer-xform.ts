import type {
  HeaderFooterNodeName,
  HeaderFooterModel,
} from '../../xform/sheet/header-footer-xform.js';
export type {
  HeaderFooterNodeName,
  HeaderFooterModel,
} from '../../xform/sheet/header-footer-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class HeaderFooterXform extends BaseXform<HeaderFooterModel> {
  currentNode?: HeaderFooterNodeName;

  override tag = 'headerFooter';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case 'headerFooter':
        this.model = {};
        if (node.attributes.differentFirst) {
          this.model.differentFirst = parseInt(node.attributes.differentFirst, 0) === 1;
        }
        if (node.attributes.differentOddEven) {
          this.model.differentOddEven = parseInt(node.attributes.differentOddEven, 0) === 1;
        }
        return true;

      case 'oddHeader':
        this.currentNode = 'oddHeader';
        return true;

      case 'oddFooter':
        this.currentNode = 'oddFooter';
        return true;

      case 'evenHeader':
        this.currentNode = 'evenHeader';
        return true;

      case 'evenFooter':
        this.currentNode = 'evenFooter';
        return true;

      case 'firstHeader':
        this.currentNode = 'firstHeader';
        return true;

      case 'firstFooter':
        this.currentNode = 'firstFooter';
        return true;

      default:
        return false;
    }
  }

  override parseText(text: string): void {
    switch (this.currentNode) {
      case 'oddHeader':
        this.model!.oddHeader = text;
        break;

      case 'oddFooter':
        this.model!.oddFooter = text;
        break;

      case 'evenHeader':
        this.model!.evenHeader = text;
        break;

      case 'evenFooter':
        this.model!.evenFooter = text;
        break;

      case 'firstHeader':
        this.model!.firstHeader = text;
        break;

      case 'firstFooter':
        this.model!.firstFooter = text;
        break;

      default:
        break;
    }
  }

  override parseClose(): boolean {
    switch (this.currentNode) {
      case 'oddHeader':
      case 'oddFooter':
      case 'evenHeader':
      case 'evenFooter':
      case 'firstHeader':
      case 'firstFooter':
        this.currentNode = undefined;
        return true;

      default:
        return false;
    }
  }
}

export default HeaderFooterXform;
export {HeaderFooterXform};
