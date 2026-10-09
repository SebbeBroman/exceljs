import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import ColorXform from '../style/color-xform.js';
import type {ColorModel} from '../style/color-xform.js';
import PageSetupPropertiesXform from './page-setup-properties-xform.js';
import type {PageSetupPropertiesModel} from './page-setup-properties-xform.js';
import OutlinePropertiesXform from './outline-properties-xform.js';
import type {OutlinePropertiesModel} from './outline-properties-xform.js';

export interface SheetPropertiesModel {
  tabColor?: ColorModel;
  pageSetup?: PageSetupPropertiesModel;
  outlineProperties?: OutlinePropertiesModel;
}

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

  override render(xmlStream: XmlStreamLike, model?: SheetPropertiesModel | null): void {
    if (model) {
      xmlStream.addRollback();
      xmlStream.openNode('sheetPr');

      let inner = false;
      // ColorXform / property xforms return boolean at runtime (base type is void)
      inner = (this.map.tabColor.render(xmlStream, model.tabColor) as unknown as boolean) || inner;
      inner =
        (this.map.pageSetUpPr.render(xmlStream, model.pageSetup) as unknown as boolean) || inner;
      inner =
        (this.map.outlinePr.render(xmlStream, model.outlineProperties) as unknown as boolean) ||
        inner;

      if (inner) {
        xmlStream.closeNode();
        xmlStream.commit();
      } else {
        xmlStream.rollback();
      }
    }
  }
}

export default SheetPropertiesXform;
export {SheetPropertiesXform};
