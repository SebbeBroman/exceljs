import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

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

function xmlToBoolean(value: string | undefined, equals: string): true | undefined {
  return value === equals ? true : undefined;
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
