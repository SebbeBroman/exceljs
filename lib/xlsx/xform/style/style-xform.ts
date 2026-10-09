import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import AlignmentXform from './alignment-xform.js';
import type {AlignmentModel} from './alignment-xform.js';
import ProtectionXform from './protection-xform.js';
import type {ProtectionModel} from './protection-xform.js';

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
//   Optional <protection>
// </xf>

export interface StyleModel {
  numFmtId?: number;
  fontId?: number;
  fillId?: number;
  borderId?: number;
  xfId?: number;
  alignment?: AlignmentModel | null;
  protection?: ProtectionModel | null;
}

// Style assists translation from style model to/from xlsx
class StyleXform extends BaseXform<StyleModel> {
  xfId: boolean;
  declare map: {
    alignment: AlignmentXform;
    protection: ProtectionXform;
  };

  constructor(options?: {xfId?: boolean}) {
    super();

    this.xfId = !!(options && options.xfId);
    this.map = {
      alignment: new AlignmentXform(),
      protection: new ProtectionXform(),
    };
  }

  override tag = 'xf';

  override render(xmlStream: XmlStreamLike, model?: StyleModel | null): void {
    xmlStream.openNode('xf', {
      numFmtId: model!.numFmtId || 0,
      fontId: model!.fontId || 0,
      fillId: model!.fillId || 0,
      borderId: model!.borderId || 0,
    });
    if (this.xfId) {
      xmlStream.addAttribute('xfId', model!.xfId || 0);
    }

    if (model!.numFmtId) {
      xmlStream.addAttribute('applyNumberFormat', '1');
    }
    if (model!.fontId) {
      xmlStream.addAttribute('applyFont', '1');
    }
    if (model!.fillId) {
      xmlStream.addAttribute('applyFill', '1');
    }
    if (model!.borderId) {
      xmlStream.addAttribute('applyBorder', '1');
    }
    if (model!.alignment) {
      xmlStream.addAttribute('applyAlignment', '1');
    }
    if (model!.protection) {
      xmlStream.addAttribute('applyProtection', '1');
    }

    /**
     * Rendering tags causes close of XML stream.
     * Therefore adding attributes must be done before rendering tags.
     */

    if (model!.alignment) {
      this.map.alignment.render(xmlStream, model!.alignment);
    }
    if (model!.protection) {
      this.map.protection.render(xmlStream, model!.protection);
    }

    xmlStream.closeNode();
  }
}

export default StyleXform;
export {StyleXform};
