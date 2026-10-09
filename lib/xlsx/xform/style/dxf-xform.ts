import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import AlignmentXform from './alignment-xform.js';
import BorderXform from './border-xform.js';
import FillXform from './fill-xform.js';
import FontXform from './font-xform.js';
import NumFmtXform from './numfmt-xform.js';
import ProtectionXform from './protection-xform.js';

// <xf numFmtId="[numFmtId]" fontId="[fontId]" fillId="[fillId]" borderId="[xf.borderId]" xfId="[xfId]">
//   Optional <alignment>
//   Optional <protection>
// </xf>

export interface DxfModel {
  font?: unknown;
  numFmt?: string;
  numFmtId?: number;
  fill?: unknown;
  alignment?: unknown;
  border?: unknown;
  protection?: unknown;
}

// Style assists translation from style model to/from xlsx
class DxfXform extends BaseXform<DxfModel> {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  declare map: any;

  constructor() {
    super();

    this.map = {
      alignment: new AlignmentXform(),
      border: new BorderXform(),
      fill: new FillXform(),
      font: new FontXform(),
      numFmt: new NumFmtXform(),
      protection: new ProtectionXform(),
    };
  }

  override tag = 'dxf';

  // how do we generate dxfid?

  override render(xmlStream: XmlStreamLike, model?: DxfModel | null): void {
    xmlStream.openNode(this.tag);

    if (model!.font) {
      this.map.font.render(xmlStream, model!.font as Record<string, unknown>);
    }
    if (model!.numFmt && model!.numFmtId) {
      const numFmtModel = {id: model!.numFmtId, formatCode: model!.numFmt};
      this.map.numFmt.render(xmlStream, numFmtModel);
    }
    if (model!.fill) {
      this.map.fill.render(xmlStream, model!.fill as never);
    }
    if (model!.alignment) {
      this.map.alignment.render(xmlStream, model!.alignment as never);
    }
    if (model!.border) {
      this.map.border.render(xmlStream, model!.border as never);
    }
    if (model!.protection) {
      this.map.protection.render(xmlStream, model!.protection as never);
    }

    xmlStream.closeNode();
  }
}

export default DxfXform;
export {DxfXform};
