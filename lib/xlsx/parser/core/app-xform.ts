import type {AppModel} from '../../xform/core/app-xform.js';
export type {AppModel} from '../../xform/core/app-xform.js';
import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import StringXform from '../simple/string-xform.js';
import AppHeadingPairsXform from './app-heading-pairs-xform.js';
import AppTitleOfPartsXform from './app-titles-of-parts-xform.js';

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

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'Properties':
        return true;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }

        // there's a lot we don't bother to parse
        return false;
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'Properties':
        this.model = {
          worksheets: this.map.TitleOfParts.model as {name: string}[] | undefined,
          company: this.map.Company.model as string | undefined,
          manager: this.map.Manager.model as string | undefined,
        };
        return false;
      default:
        return true;
    }
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
