import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import type {CacheFieldData} from './cache-field.js';

export interface PivotCacheRecordsModel {
  sourceSheet: {
    getSheetValues(): unknown[][];
  };
  cacheFields: CacheFieldData[];
}

class PivotCacheRecordsXform extends BaseXform<PivotCacheRecordsModel> {
  declare map: Record<string, never>;

  constructor() {
    super();

    this.map = {};
  }

  override prepare(_model?: PivotCacheRecordsModel | null): void {
    // TK
  }

  override tag = 'pivotCacheRecords';

  override render(xmlStream: XmlStreamLike, model?: PivotCacheRecordsModel | null): void {
    const {sourceSheet, cacheFields} = model!;
    const sourceBodyRows = sourceSheet.getSheetValues().slice(2);

    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheRecordsXform.PIVOT_CACHE_RECORDS_ATTRIBUTES,
      count: sourceBodyRows.length,
    });
    xmlStream.writeXml(renderTable());
    xmlStream.closeNode();

    // Helpers

    function renderTable(): string {
      const rowsInXML = sourceBodyRows.map(row => {
        const realRow = (row as unknown[]).slice(1);
        return [...renderRowLines(realRow)].join('');
      });
      return rowsInXML.join('');
    }

    function* renderRowLines(row: unknown[]): Generator<string> {
      // PivotCache Record: http://www.datypic.com/sc/ooxml/e-ssml_r-1.html
      // Note: pretty-printing this for now to ease debugging.
      yield '\n  <r>';
      for (const [index, cellValue] of row.entries()) {
        yield '\n    ';
        yield renderCell(cellValue, cacheFields[index].sharedItems);
      }
      yield '\n  </r>';
    }

    function renderCell(value: unknown, sharedItems: string[] | null): string {
      // no shared items
      // --------------------------------------------------
      if (sharedItems === null) {
        if (Number.isFinite(value)) {
          // Numeric value: http://www.datypic.com/sc/ooxml/e-ssml_n-2.html
          return `<n v="${value}" />`;
        }
        // Character Value: http://www.datypic.com/sc/ooxml/e-ssml_s-2.html
        return `<s v="${value}" />`;
      }

      // shared items
      // --------------------------------------------------
      const sharedItemsIndex = sharedItems.indexOf(value as string);
      if (sharedItemsIndex < 0) {
        throw new Error(
          `${JSON.stringify(value)} not in sharedItems ${JSON.stringify(sharedItems)}`,
        );
      }
      // Shared Items Index: http://www.datypic.com/sc/ooxml/e-ssml_x-9.html
      return `<x v="${sharedItemsIndex}" />`;
    }
  }

  static PIVOT_CACHE_RECORDS_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'xr',
    'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  };
}

export default PivotCacheRecordsXform;
export {PivotCacheRecordsXform};
