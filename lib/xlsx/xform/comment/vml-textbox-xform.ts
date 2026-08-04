import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

  reverseConversionUnit(inset: string | undefined): number[] {
    return (inset || '').split(',').map(margin => {
      return Number(parseFloat(this.conversionUnit(parseFloat(margin), 0.1, '')).toFixed(2));
    });
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

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          inset: this.reverseConversionUnit(node.attributes.inset),
        };
        return true;
      default:
        return true;
    }
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    switch (name) {
      case this.tag:
        return false;
      default:
        return true;
    }
  }
}

export default VmlTextboxXform;
export {VmlTextboxXform};
