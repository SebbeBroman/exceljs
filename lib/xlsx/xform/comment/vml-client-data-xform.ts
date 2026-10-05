import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import VmlAnchorXform from './vml-anchor-xform.js';
import type {VmlAnchorRenderModel} from './vml-anchor-xform.js';
import VmlProtectionXform from './style/vml-protection-xform.js';
import VmlPositionXform from './style/vml-position-xform.js';

const POSITION_TYPE = ['twoCells', 'oneCells', 'absolute'] as const;

export interface VmlClientDataModel {
  anchor?: string | number[];
  protection: {
    locked?: string;
    lockText?: string;
  };
  editAs?: string;
}

export interface VmlClientDataRenderModel extends VmlAnchorRenderModel {
  note: {
    protection: {
      locked?: string;
      lockText?: string;
    };
    editAs?: string;
    margins?: {
      inset?: number[] | string;
      insetmode?: string;
    };
  };
  refAddress: {
    row: number;
    col: number;
  };
}

class VmlClientDataXform extends BaseXform<VmlClientDataModel> {
  override tag = 'x:ClientData';
  declare map: {
    'x:Anchor': VmlAnchorXform;
    'x:Locked': VmlProtectionXform;
    'x:LockText': VmlProtectionXform;
    'x:SizeWithCells': VmlPositionXform;
    'x:MoveWithCells': VmlPositionXform;
  };

  constructor() {
    super();
    this.map = {
      'x:Anchor': new VmlAnchorXform(),
      'x:Locked': new VmlProtectionXform({tag: 'x:Locked'}),
      'x:LockText': new VmlProtectionXform({tag: 'x:LockText'}),
      'x:SizeWithCells': new VmlPositionXform({tag: 'x:SizeWithCells'}),
      'x:MoveWithCells': new VmlPositionXform({tag: 'x:MoveWithCells'}),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: VmlClientDataModel | null): void {
    const renderModel = model as unknown as VmlClientDataRenderModel;
    const {protection, editAs} = renderModel.note;
    xmlStream.openNode(this.tag, {ObjectType: 'Note'});
    this.map['x:MoveWithCells'].render(xmlStream, editAs, POSITION_TYPE as unknown as number);
    this.map['x:SizeWithCells'].render(xmlStream, editAs, POSITION_TYPE as unknown as number);
    this.map['x:Anchor'].render(xmlStream, renderModel);
    this.map['x:Locked'].render(xmlStream, protection.locked);
    xmlStream.leafNode('x:AutoFill', undefined, 'False');
    this.map['x:LockText'].render(xmlStream, protection.lockText);
    xmlStream.leafNode('x:Row', undefined, renderModel.refAddress.row - 1);
    xmlStream.leafNode('x:Column', undefined, renderModel.refAddress.col - 1);
    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.reset();
        this.model = {
          anchor: [],
          protection: {},
          editAs: '',
        };
        break;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        break;
    }
    return true;
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case this.tag:
        this.normalizeModel();
        return false;
      default:
        return true;
    }
  }

  normalizeModel(): void {
    const position = Object.assign(
      {},
      this.map['x:MoveWithCells'].model,
      this.map['x:SizeWithCells'].model,
    );
    const len = Object.keys(position).length;
    this.model!.editAs = POSITION_TYPE[len];
    this.model!.anchor = this.map['x:Anchor'].text;
    this.model!.protection.locked = this.map['x:Locked'].text;
    this.model!.protection.lockText = this.map['x:LockText'].text;
  }
}

export default VmlClientDataXform;
export {VmlClientDataXform};
