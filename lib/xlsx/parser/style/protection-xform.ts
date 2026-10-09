import type {ProtectionModel} from '../../xform/style/protection-xform.js';
export type {ProtectionModel} from '../../xform/style/protection-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

// Protection encapsulates translation from style.protection model to/from xlsx
class ProtectionXform extends BaseXform<ProtectionModel | null> {
  override tag = 'protection';

  override parseOpen(node: XmlNode): void {
    const model = {
      locked: !(node.attributes.locked === '0'),
      hidden: node.attributes.hidden === '1',
    };

    // only want to record models that differ from defaults
    const isSignificant = !model.locked || model.hidden;

    this.model = isSignificant ? model : null;
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default ProtectionXform;
export {ProtectionXform};
