import type {SheetPropertiesModel} from '../../xform/sheet/sheet-properties-xform.js';
export type {SheetPropertiesModel} from '../../xform/sheet/sheet-properties-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import ColorXform from '../style/color-xform.js';
import PageSetupPropertiesXform from './page-setup-properties-xform.js';
import OutlinePropertiesXform from './outline-properties-xform.js';

class SheetPropertiesXform extends BaseXform<SheetPropertiesModel> {
  override map: {
    tabColor: ColorXform;
    pageSetUpPr: PageSetupPropertiesXform;
    outlinePr: OutlinePropertiesXform;
  };

  constructor() {
    super();

    this.map = {
      tabColor: new ColorXform('tabColor'),
      pageSetUpPr: new PageSetupPropertiesXform(),
      outlinePr: new OutlinePropertiesXform(),
    };
  }

  override tag = 'sheetPr';

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    if (node.name === this.tag) {
      this.reset();
      return true;
    }
    const child = this.map[node.name as keyof typeof this.map] as BaseXform | undefined;
    if (child) {
      this.parser = child;
      this.parser.parseOpen(node);
      return true;
    }
    return false;
  }

  override parseText(text: string): boolean {
    if (this.parser) {
      this.parser.parseText(text);
      return true;
    }
    return false;
  }

  override parseClose(_name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(_name)) {
        this.parser = undefined;
      }
      return true;
    }
    if (this.map.tabColor.model || this.map.pageSetUpPr.model || this.map.outlinePr.model) {
      this.model = {};
      if (this.map.tabColor.model) {
        this.model.tabColor = this.map.tabColor.model;
      }
      if (this.map.pageSetUpPr.model) {
        this.model.pageSetup = this.map.pageSetUpPr.model;
      }
      if (this.map.outlinePr.model) {
        this.model.outlineProperties = this.map.outlinePr.model;
      }
    } else {
      this.model = null;
    }
    return false;
  }
}

export default SheetPropertiesXform;
export {SheetPropertiesXform};
