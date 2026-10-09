import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';

export interface SheetProtectionModel {
  sheet?: boolean;
  objects?: boolean;
  scenarios?: boolean;
  selectLockedCells?: boolean;
  selectUnlockedCells?: boolean;
  formatCells?: boolean;
  formatColumns?: boolean;
  formatRows?: boolean;
  insertColumns?: boolean;
  insertRows?: boolean;
  insertHyperlinks?: boolean;
  deleteColumns?: boolean;
  deleteRows?: boolean;
  sort?: boolean;
  autoFilter?: boolean;
  pivotTables?: boolean;
  algorithmName?: string;
  hashValue?: string;
  saltValue?: string;
  spinCount?: number;
}

function booleanToXml(model: boolean | undefined, value: string): string | undefined {
  return model ? value : undefined;
}

class SheetProtectionXform extends BaseXform<SheetProtectionModel> {
  override tag = 'sheetProtection';

  override render(xmlStream: XmlStreamLike, model?: SheetProtectionModel | null): void {
    if (model) {
      const attributes: Record<string, unknown> = {
        sheet: booleanToXml(model.sheet, '1'),
        selectLockedCells: model.selectLockedCells === false ? '1' : undefined,
        selectUnlockedCells: model.selectUnlockedCells === false ? '1' : undefined,
        formatCells: booleanToXml(model.formatCells, '0'),
        formatColumns: booleanToXml(model.formatColumns, '0'),
        formatRows: booleanToXml(model.formatRows, '0'),
        insertColumns: booleanToXml(model.insertColumns, '0'),
        insertRows: booleanToXml(model.insertRows, '0'),
        insertHyperlinks: booleanToXml(model.insertHyperlinks, '0'),
        deleteColumns: booleanToXml(model.deleteColumns, '0'),
        deleteRows: booleanToXml(model.deleteRows, '0'),
        sort: booleanToXml(model.sort, '0'),
        autoFilter: booleanToXml(model.autoFilter, '0'),
        pivotTables: booleanToXml(model.pivotTables, '0'),
      };
      if (model.sheet) {
        attributes.algorithmName = model.algorithmName;
        attributes.hashValue = model.hashValue;
        attributes.saltValue = model.saltValue;
        attributes.spinCount = model.spinCount;
        attributes.objects = booleanToXml(model.objects === false, '1');
        attributes.scenarios = booleanToXml(model.scenarios === false, '1');
      }
      if (Object.values(attributes).some(value => value !== undefined)) {
        xmlStream.leafNode(this.tag, attributes);
      }
    }
  }
}

export default SheetProtectionXform;
export {SheetProtectionXform};
