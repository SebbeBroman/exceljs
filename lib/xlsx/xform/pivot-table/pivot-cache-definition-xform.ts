import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import CacheField from './cache-field.js';
import type {CacheFieldData} from './cache-field.js';
import XmlStream from '../../../utils/xml-stream.js';

export interface PivotCacheDefinitionModel {
  sourceSheet: {
    dimensions: {shortRange: string};
    name: string;
  };
  cacheFields: CacheFieldData[];
}

class PivotCacheDefinitionXform extends BaseXform<PivotCacheDefinitionModel> {
  declare map: Record<string, never>;

  constructor() {
    super();

    this.map = {};
  }

  override prepare(_model?: PivotCacheDefinitionModel | null): void {
    // TK
  }

  override tag = 'pivotCacheDefinition';

  override render(xmlStream: XmlStreamLike, model?: PivotCacheDefinitionModel | null): void {
    const {sourceSheet, cacheFields} = model!;

    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);
    xmlStream.openNode(this.tag, {
      ...PivotCacheDefinitionXform.PIVOT_CACHE_DEFINITION_ATTRIBUTES,
      'r:id': 'rId1',
      refreshOnLoad: '1', // important for our implementation to work
      refreshedBy: 'Author',
      refreshedDate: '45125.026046874998',
      createdVersion: '8',
      refreshedVersion: '8',
      minRefreshableVersion: '3',
      recordCount: cacheFields.length + 1,
    });

    xmlStream.openNode('cacheSource', {type: 'worksheet'});
    xmlStream.leafNode('worksheetSource', {
      ref: sourceSheet.dimensions.shortRange,
      sheet: sourceSheet.name,
    });
    xmlStream.closeNode();

    xmlStream.openNode('cacheFields', {count: cacheFields.length});
    // Note: keeping this pretty-printed for now to ease debugging.
    xmlStream.writeXml(
      cacheFields.map(cacheField => new CacheField(cacheField).render()).join('\n    '),
    );
    xmlStream.closeNode();

    xmlStream.closeNode();
  }

  override parseOpen(_node?: unknown): void {
    // TK
  }

  override parseText(_text?: string): void {
    // TK
  }

  override parseClose(_name?: string): void {
    // TK
  }

  override reconcile(_model?: PivotCacheDefinitionModel | null, _options?: unknown): void {
    // TK
  }

  static PIVOT_CACHE_DEFINITION_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'xr',
    'xmlns:xr': 'http://schemas.microsoft.com/office/spreadsheetml/2014/revision',
  };
}

export default PivotCacheDefinitionXform;
export {PivotCacheDefinitionXform};
