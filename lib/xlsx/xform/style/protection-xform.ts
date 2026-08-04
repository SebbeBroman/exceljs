import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

export interface ProtectionModel {
  locked?: boolean;
  hidden?: boolean;
}

const validation = {
  boolean(value: boolean | undefined, dflt: boolean): boolean {
    if (value === undefined) {
      return dflt;
    }
    return value;
  },
};

// Protection encapsulates translation from style.protection model to/from xlsx
class ProtectionXform extends BaseXform<ProtectionModel | null> {
  override tag = 'protection';

  override render(xmlStream: XmlStreamLike, model?: ProtectionModel | null): void {
    xmlStream.addRollback();
    xmlStream.openNode('protection');

    let isValid = false;
    function add(name: string, value: string | undefined): void {
      if (value !== undefined) {
        xmlStream.addAttribute(name, value);
        isValid = true;
      }
    }
    add('locked', validation.boolean(model!.locked, true) ? undefined : '0');
    add('hidden', validation.boolean(model!.hidden, false) ? '1' : undefined);

    xmlStream.closeNode();

    if (isValid) {
      xmlStream.commit();
    } else {
      xmlStream.rollback();
    }
  }

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
