import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface VmlTextboxModel {
  inset?: number[];
}

export interface VmlTextboxRenderModel {
  note?: {
    margins?: {
      inset?: number[] | string;
      insetmode?: string;
    };
  };
}

class VmlTextboxXform extends BaseXform<VmlTextboxModel> {
  override tag = 'v:textbox';

  conversionUnit(value: number | string, multiple: number, unit: string): string {
    // Preserve original operator behaviour: multiple.toFixed(2) is coerced via *
    return `${parseFloat(String(value)) * Number(multiple.toFixed(2))}${unit}`;
  }

  override render(xmlStream: XmlStreamLike, model?: VmlTextboxModel | null): void {
    const attributes: Record<string, unknown> = {
      style: 'mso-direction-alt:auto',
    };
    const renderModel = model as unknown as VmlTextboxRenderModel | null | undefined;
    if (renderModel && renderModel.note) {
      let {inset} = (renderModel.note && renderModel.note.margins) as {
        inset?: number[] | string;
      };
      if (Array.isArray(inset)) {
        inset = inset
          .map(margin => {
            return this.conversionUnit(margin, 10, 'mm');
          })
          .join(',');
      }
      if (inset) {
        attributes.inset = inset;
      }
    }
    xmlStream.openNode('v:textbox', attributes);
    xmlStream.leafNode('div', {style: 'text-align:left'});
    xmlStream.closeNode();
  }
}

export default VmlTextboxXform;
export {VmlTextboxXform};
