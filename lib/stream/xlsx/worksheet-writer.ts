import {columnsToModel, rowToModel} from '../../model/row-model.js';
import type {ColumnInput, RowInput} from '../../model/types.js';
import StringBuf from '../../utils/string-buf.js';
import SheetRelsWriter from './sheet-rels-writer.js';
import ListXform from '../../xlsx/xform/list-xform.js';
import SheetPropertiesXform from '../../xlsx/xform/sheet/sheet-properties-xform.js';
import SheetFormatPropertiesXform from '../../xlsx/xform/sheet/sheet-format-properties-xform.js';
import ColXform from '../../xlsx/xform/sheet/col-xform.js';
import RowXform from '../../xlsx/xform/sheet/row-xform.js';
import HyperlinkXform from '../../xlsx/xform/sheet/hyperlink-xform.js';
import SheetViewXform from '../../xlsx/xform/sheet/sheet-view-xform.js';
import PageMarginsXform from '../../xlsx/xform/sheet/page-margins-xform.js';
import PageSetupXform from '../../xlsx/xform/sheet/page-setup-xform.js';
import AutoFilterXform from '../../xlsx/xform/sheet/auto-filter-xform.js';
import HeaderFooterXform from '../../xlsx/xform/sheet/header-footer-xform.js';

const xmlBuffer = new StringBuf();

// ============================================================================================
// Xforms

// since prepare and render are functional, we can use singletons
const xform = {
  sheetProperties: new SheetPropertiesXform(),
  sheetFormatProperties: new SheetFormatPropertiesXform(),
  columns: new ListXform({tag: 'cols', count: false, childXform: new ColXform()}),
  row: new RowXform(),
  hyperlinks: new ListXform({tag: 'hyperlinks', count: false, childXform: new HyperlinkXform()}),
  sheetViews: new ListXform({tag: 'sheetViews', count: false, childXform: new SheetViewXform()}),
  pageMargins: new PageMarginsXform(),
  pageSetup: new PageSetupXform(),
  autoFilter: new AutoFilterXform(),
  headerFooter: new HeaderFooterXform(),
};

// ============================================================================================

export interface WorksheetWriterOptions {
  id: number | string;
  name?: string;
  state?: string;
  useSharedStrings?: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbook: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  properties?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pageSetup?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  views?: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  autoFilter?: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  headerFooter?: any;
}

class WorksheetWriter {
  id: number | string;
  name: string;
  state: string;
  rId?: string;
  columns: ColumnInput[] = [];
  private nextRow = 1;
  _sheetRelsWriter: SheetRelsWriter;
  committed: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _formulae: Record<string, any>;
  _siFormulae: number;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  properties: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  headerFooter: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  pageSetup: any;
  useSharedStrings: boolean;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _workbook: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _views: any[];
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  autoFilter: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _stream?: any;
  startedData: boolean;

  constructor(options: WorksheetWriterOptions) {
    // in a workbook, each sheet will have a number
    this.id = options.id;

    // and a name
    this.name = options.name || `Sheet${this.id}`;

    // add a state
    this.state = options.state || 'visible';

    this._sheetRelsWriter = new SheetRelsWriter(options);
    this.committed = false;
    // for sharing formulae
    this._formulae = {};
    this._siFormulae = 0;

    // for default row height, outline levels, etc
    this.properties = Object.assign(
      {},
      {
        defaultRowHeight: 15,
        dyDescent: 55,
        outlineLevelCol: 0,
        outlineLevelRow: 0,
      },
      options.properties,
    );

    this.headerFooter = Object.assign(
      {},
      {
        differentFirst: false,
        differentOddEven: false,
        oddHeader: null,
        oddFooter: null,
        evenHeader: null,
        evenFooter: null,
        firstHeader: null,
        firstFooter: null,
      },
      options.headerFooter,
    );

    // for all things printing
    this.pageSetup = Object.assign(
      {},
      {
        margins: {left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3},
        orientation: 'portrait',
        horizontalDpi: 4294967295,
        verticalDpi: 4294967295,
        fitToPage: !!(
          options.pageSetup &&
          (options.pageSetup.fitToWidth || options.pageSetup.fitToHeight) &&
          !options.pageSetup.scale
        ),
        pageOrder: 'downThenOver',
        blackAndWhite: false,
        draft: false,
        cellComments: 'None',
        errors: 'displayed',
        scale: 100,
        fitToWidth: 1,
        fitToHeight: 1,
        paperSize: undefined,
        showRowColHeaders: false,
        showGridLines: false,
        horizontalCentered: false,
        verticalCentered: false,
        rowBreaks: null,
        colBreaks: null,
      },
      options.pageSetup,
    );

    // using shared strings creates a smaller xlsx file but may use more memory
    this.useSharedStrings = options.useSharedStrings || false;

    this._workbook = options.workbook;

    // views
    this._views = options.views || [];

    // auto filter
    this.autoFilter = options.autoFilter || null;

    // start writing to stream now
    this._writeOpenWorksheet();

    this.startedData = false;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  get workbook(): any {
    return this._workbook;
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  get stream(): any {
    if (!this._stream) {
      this._stream = this._workbook._openStream(`/xl/worksheets/sheet${this.id}.xml`);

      // pause stream to prevent 'data' events
      this._stream.pause();
    }
    return this._stream;
  }

  commit(): void {
    if (this.committed) {
      return;
    }
    if (!this.startedData) {
      this._writeOpenSheetData();
    }
    this._writeCloseSheetData();
    this._writeAutoFilter();

    // for some reason, Excel can't handle dimensions at the bottom of the file
    // this._writeDimensions();

    this._writeHyperlinks();
    this._writePageMargins();
    this._writePageSetup();
    this._writeHeaderFooter();

    this._writeCloseWorksheet();
    // signal end of stream to workbook
    this.stream.end();

    // also commit the hyperlinks if any
    this._sheetRelsWriter.commit();

    this.committed = true;
  }

  get views(): any[] {
    return this._views;
  }

  writeRow(values: RowInput): void {
    if (this.committed) throw new Error('Cannot write to a committed worksheet');
    const model = rowToModel(this.nextRow++, values, this.columns);
    this._writeRow(model);
  }

  _write(text: string): void {
    xmlBuffer.reset();
    xmlBuffer.addText(text);
    this.stream.write(xmlBuffer);
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _writeSheetProperties(xmlBuf: any, properties: any, pageSetup: any): void {
    const sheetPropertiesModel = {
      outlineProperties: properties && properties.outlineProperties,
      tabColor: properties && properties.tabColor,
      pageSetup:
        pageSetup && pageSetup.fitToPage
          ? {
              fitToPage: pageSetup.fitToPage,
            }
          : undefined,
    };

    xmlBuf.addText(xform.sheetProperties.toXml(sheetPropertiesModel));
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _writeSheetFormatProperties(xmlBuf: any, properties: any): void {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const sheetFormatPropertiesModel: any = properties
      ? {
          defaultRowHeight: properties.defaultRowHeight,
          dyDescent: properties.dyDescent,
          outlineLevelCol: properties.outlineLevelCol,
          outlineLevelRow: properties.outlineLevelRow,
        }
      : undefined;
    if (properties.defaultColWidth) {
      sheetFormatPropertiesModel.defaultColWidth = properties.defaultColWidth;
    }

    xmlBuf.addText(xform.sheetFormatProperties.toXml(sheetFormatPropertiesModel));
  }

  _writeOpenWorksheet(): void {
    xmlBuffer.reset();

    xmlBuffer.addText('<?xml version="1.0" encoding="UTF-8" standalone="yes"?>');
    xmlBuffer.addText(
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"' +
        ' xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"' +
        ' xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"' +
        ' mc:Ignorable="x14ac"' +
        ' xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac">',
    );

    this._writeSheetProperties(xmlBuffer, this.properties, this.pageSetup);

    xmlBuffer.addText(xform.sheetViews.toXml(this.views));

    this._writeSheetFormatProperties(xmlBuffer, this.properties);

    this.stream.write(xmlBuffer);
  }

  _writeColumns(): void {
    const cols = columnsToModel(this.columns, this.properties.outlineLevelCol);
    if (cols) {
      xform.columns.prepare(cols, {styles: this._workbook.styles});
      this.stream.write(xform.columns.toXml(cols));
    }
  }

  _writeOpenSheetData(): void {
    this._write('<sheetData>');
  }

  _writeRow(model: ReturnType<typeof rowToModel>): void {
    if (!this.startedData) {
      this._writeColumns();
      this._writeOpenSheetData();
      this.startedData = true;
    }

    if (model.cells.some(cell => cell.type !== 0)) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const options: any = {
        styles: this._workbook.styles,
        sharedStrings: this.useSharedStrings ? this._workbook.sharedStrings : undefined,
        hyperlinks: this._sheetRelsWriter.hyperlinksProxy,
        merges: {add() {}},
        formulae: this._formulae,
        siFormulae: this._siFormulae,
        comments: [],
      };
      xform.row.prepare(model as any, options);
      this.stream.write(xform.row.toXml(model as any));

      this._siFormulae = options.siFormulae;
    }
  }

  _writeCloseSheetData(): void {
    this._write('</sheetData>');
  }

  _writeHyperlinks(): void {
    this.stream.write(xform.hyperlinks.toXml(this._sheetRelsWriter._hyperlinks));
  }

  _writePageMargins(): void {
    this.stream.write(xform.pageMargins.toXml(this.pageSetup.margins));
  }

  _writePageSetup(): void {
    this.stream.write(xform.pageSetup.toXml(this.pageSetup));
  }

  _writeHeaderFooter(): void {
    this.stream.write(xform.headerFooter.toXml(this.headerFooter));
  }

  _writeAutoFilter(): void {
    this.stream.write(xform.autoFilter.toXml(this.autoFilter));
  }

  _writeCloseWorksheet(): void {
    this._write('</worksheet>');
  }
}

export default WorksheetWriter;
export {WorksheetWriter};
