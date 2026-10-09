import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike} from '../base-xform.js';
import StringXform from '../simple/string-xform.js';
import AppHeadingPairsXform from './app-heading-pairs-xform.js';
import AppTitleOfPartsXform from './app-titles-of-parts-xform.js';

export interface AppModel {
  worksheets?: {name: string}[];
  company?: string;
  manager?: string;
}

class AppXform extends BaseXform<AppModel> {
  declare map: {
    Company: StringXform;
    Manager: StringXform;
    HeadingPairs: AppHeadingPairsXform;
    TitleOfParts: AppTitleOfPartsXform;
  };

  constructor() {
    super();

    this.map = {
      Company: new StringXform({tag: 'Company'}),
      Manager: new StringXform({tag: 'Manager'}),
      HeadingPairs: new AppHeadingPairsXform(),
      TitleOfParts: new AppTitleOfPartsXform(),
    };
  }

  override render(xmlStream: XmlStreamLike, model?: AppModel | null): void {
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);

    xmlStream.openNode('Properties', AppXform.PROPERTY_ATTRIBUTES);

    xmlStream.leafNode('Application', undefined, 'Microsoft Excel');
    xmlStream.leafNode('DocSecurity', undefined, '0');
    xmlStream.leafNode('ScaleCrop', undefined, 'false');

    this.map.HeadingPairs.render(xmlStream, model!.worksheets as never);
    this.map.TitleOfParts.render(xmlStream, model!.worksheets as never);
    this.map.Company.render(xmlStream, model!.company || '');
    this.map.Manager.render(xmlStream, model!.manager);

    xmlStream.leafNode('LinksUpToDate', undefined, 'false');
    xmlStream.leafNode('SharedDoc', undefined, 'false');
    xmlStream.leafNode('HyperlinksChanged', undefined, 'false');
    xmlStream.leafNode('AppVersion', undefined, '16.0300');

    xmlStream.closeNode();
  }

  static DateFormat(dt: Date): string {
    return dt.toISOString().replace(/[.]\d{3,6}/, '');
  }

  static DateAttrs = {'xsi:type': 'dcterms:W3CDTF'};

  static PROPERTY_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/officeDocument/2006/extended-properties',
    'xmlns:vt': 'http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes',
  };
}

export default AppXform;
export {AppXform};
