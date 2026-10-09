import type {SheetProtectionModel} from '../../xform/sheet/sheet-protection-xform.js';
export type {SheetProtectionModel} from '../../xform/sheet/sheet-protection-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

function xmlToBoolean(value: string | undefined, equals: string): true | undefined {
  return value === equals ? true : undefined;
}

class SheetProtectionXform extends BaseXform<SheetProtectionModel> {
  override tag = 'sheetProtection';

  override parseOpen(node: XmlNode): boolean {
    switch (node.name) {
      case this.tag:
        this.model = {
          sheet: xmlToBoolean(node.attributes.sheet, '1'),
          objects: node.attributes.objects === '1' ? false : undefined,
          scenarios: node.attributes.scenarios === '1' ? false : undefined,
          selectLockedCells: node.attributes.selectLockedCells === '1' ? false : undefined,
          selectUnlockedCells: node.attributes.selectUnlockedCells === '1' ? false : undefined,
          formatCells: xmlToBoolean(node.attributes.formatCells, '0'),
          formatColumns: xmlToBoolean(node.attributes.formatColumns, '0'),
          formatRows: xmlToBoolean(node.attributes.formatRows, '0'),
          insertColumns: xmlToBoolean(node.attributes.insertColumns, '0'),
          insertRows: xmlToBoolean(node.attributes.insertRows, '0'),
          insertHyperlinks: xmlToBoolean(node.attributes.insertHyperlinks, '0'),
          deleteColumns: xmlToBoolean(node.attributes.deleteColumns, '0'),
          deleteRows: xmlToBoolean(node.attributes.deleteRows, '0'),
          sort: xmlToBoolean(node.attributes.sort, '0'),
          autoFilter: xmlToBoolean(node.attributes.autoFilter, '0'),
          pivotTables: xmlToBoolean(node.attributes.pivotTables, '0'),
        };
        if (node.attributes.algorithmName) {
          this.model.algorithmName = node.attributes.algorithmName;
          this.model.hashValue = node.attributes.hashValue;
          this.model.saltValue = node.attributes.saltValue;
          this.model.spinCount = parseInt(node.attributes.spinCount, 10);
        }
        return true;
      default:
        return false;
    }
  }

  override parseText(): void {}

  override parseClose(): boolean {
    return false;
  }
}

export default SheetProtectionXform;
export {SheetProtectionXform};
