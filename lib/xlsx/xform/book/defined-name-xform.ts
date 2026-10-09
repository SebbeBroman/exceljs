import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface DefinedNameModel {
  name: string;
  ranges: string[];
  localSheetId?: number;
}

class DefinedNamesXform extends BaseXform<DefinedNameModel> {
  override render(xmlStream: XmlStreamLike, model?: DefinedNameModel | null): void {
    // <definedNames>
    //   <definedName name="name">name.ranges.join(',')</definedName>
    //   <definedName name="_xlnm.Print_Area" localSheetId="0">name.ranges.join(',')</definedName>
    // </definedNames>
    xmlStream.openNode('definedName', {
      name: model!.name,
      localSheetId: model!.localSheetId,
    });
    xmlStream.writeText(model!.ranges.join(','));
    xmlStream.closeNode();
  }
}

export default DefinedNamesXform;
export {DefinedNamesXform};
