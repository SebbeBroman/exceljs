/** Fixed default stylesheet for exports without cell styling. */
import Enums from '../../../model/enums.js';
import type {DxfModel} from './dxf-xform.js';

const DEFAULT_XML =
  '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac x16r2" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac" xmlns:x16r2="http://schemas.microsoft.com/office/spreadsheetml/2015/02/main"><fonts x14ac:knownFonts="1" count="1"><font><color theme="1"/><family val="2"/><scheme val="minor"/><sz val="11"/><name val="Calibri"/></font></fonts><fills count="2"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles><dxfs count="0"/><tableStyles count="0" defaultTableStyle="TableStyleMedium2" defaultPivotStyle="PivotStyleLight16"/><extLst><ext uri="{EB79DEF2-80B8-43e5-95BD-54CBDDF9020C}" xmlns:x14="http://schemas.microsoft.com/office/spreadsheetml/2009/9/main"><x14:slicerStyles defaultSlicerStyle="SlicerStyleLight1"/></ext><ext uri="{9260A510-F301-46a8-8635-F512D64BE5F5}" xmlns:x15="http://schemas.microsoft.com/office/spreadsheetml/2010/11/main"><x15:timelineStyles defaultTimelineStyle="TimeSlicerStyleLight1"/></ext></extLst></styleSheet>';

export default class MinimalStyles {
  private hasDates = false;
  private dxfs: DxfModel[] = [];

  addStyleModel(_style: unknown, cellType?: number): number {
    if (cellType === Enums.ValueType.Date) {
      this.hasDates = true;
      return 1;
    }
    return 0;
  }

  getStyleModel(): Record<string, unknown> {
    return {};
  }
  addDxfStyle(style: DxfModel): number {
    return this.dxfs.push(style) - 1;
  }
  getDxfStyle(id: number): DxfModel {
    return this.dxfs[id];
  }

  toXml(): string | Promise<string> {
    let xml = DEFAULT_XML;
    if (this.hasDates)
      xml = xml
        .replace('<cellXfs count="1">', '<cellXfs count="2">')
        .replace(
          '</cellXfs>',
          '<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs>',
        );
    if (!this.dxfs.length) return xml;
    return this.renderDxfs(xml);
  }

  private async renderDxfs(xml: string): Promise<string> {
    // Differential formats remain automatic even when cell styling is disabled.
    const [{default: DxfXform}, {default: NumFmtXform}] = await Promise.all([
      import('./dxf-xform.js'),
      import('./numfmt-xform.js'),
    ]);
    const formats = new Map<string, number>();
    const numFmtXform = new NumFmtXform();
    const dxfXform = new DxfXform();
    const dxfs = this.dxfs.map(style => {
      if (style.numFmt) {
        let id = NumFmtXform.getDefaultFmtId(style.numFmt);
        if (id === undefined) {
          id = formats.get(style.numFmt);
          if (id === undefined) {
            id = 164 + formats.size;
            formats.set(style.numFmt, id);
          }
        }
        style.numFmtId = id;
      }
      return dxfXform.toXml(style);
    });
    if (formats.size)
      xml = xml.replace(
        '<fonts ',
        `<numFmts count="${formats.size}">${[...formats].map(([formatCode, id]) => numFmtXform.toXml({id, formatCode})).join('')}</numFmts><fonts `,
      );
    return xml.replace('<dxfs count="0"/>', `<dxfs count="${dxfs.length}">${dxfs.join('')}</dxfs>`);
  }
}
