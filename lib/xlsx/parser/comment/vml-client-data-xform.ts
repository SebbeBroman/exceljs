import type {VmlClientDataModel} from '../../xform/comment/vml-client-data-xform.js';
export type {
  VmlClientDataRenderModel,
  VmlClientDataModel,
} from '../../xform/comment/vml-client-data-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import VmlAnchorXform from './vml-anchor-xform.js';
import VmlProtectionXform from './style/vml-protection-xform.js';
import VmlPositionXform from './style/vml-position-xform.js';

const POSITION_TYPE = ['twoCells', 'oneCells', 'absolute'] as const;

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
