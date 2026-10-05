import colCache from '../../../utils/col-cache.js';
import XmlStream from '../../../utils/xml-stream.js';
import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';
import StaticXform from '../static-xform.js';
import ListXform from '../list-xform.js';
import DefinedNameXform from './defined-name-xform.js';
import SheetXform from './sheet-xform.js';
import WorkbookViewXform from './workbook-view-xform.js';
import WorkbookPropertiesXform from './workbook-properties-xform.js';
import WorkbookCalcPropertiesXform from './workbook-calc-properties-xform.js';
import WorkbookPivotCacheXform from './workbook-pivot-cache-xform.js';

export interface WorkbookModel {
  worksheets?: unknown[];
  sheets?: unknown[];
  definedNames?: unknown[];
  media?: {type: string; name?: string; index?: number}[];
  properties?: unknown;
  views?: unknown[];
  calcProperties?: unknown;
  pivotTables?: unknown[];
  workbookRels?: {Id: string; Target: string}[];
  worksheetHash?: Record<string, Record<string, unknown>>;
  [key: string]: unknown;
}

class WorkbookXform extends BaseXform<WorkbookModel> {
  declare map: Record<string, BaseXform | StaticXform | ListXform>;

  constructor() {
    super();

    this.map = {
      fileVersion: WorkbookXform.STATIC_XFORMS.fileVersion,
      workbookPr: new WorkbookPropertiesXform(),
      bookViews: new ListXform({
        tag: 'bookViews',
        count: false,
        childXform: new WorkbookViewXform(),
      }),
      sheets: new ListXform({tag: 'sheets', count: false, childXform: new SheetXform()}),
      definedNames: new ListXform({
        tag: 'definedNames',
        count: false,
        childXform: new DefinedNameXform(),
      }),
      calcPr: new WorkbookCalcPropertiesXform(),
      pivotCaches: new ListXform({
        tag: 'pivotCaches',
        count: false,
        childXform: new WorkbookPivotCacheXform(),
      }),
    };
  }

  override prepare(model?: WorkbookModel | null): void {
    const m = model!;
    m.sheets = m.worksheets;

    // collate all the print areas from all of the sheets and add them to the defined names
    const printAreas: unknown[] = [];
    let index = 0; // sheets is sparse array - calc index manually
    (m.sheets as Record<string, unknown>[]).forEach(sheet => {
      const pageSetup = sheet.pageSetup as
        | {printArea?: string; printTitlesRow?: string; printTitlesColumn?: string}
        | undefined;
      if (pageSetup && pageSetup.printArea) {
        pageSetup.printArea.split('&&').forEach(printArea => {
          const printAreaComponents = printArea.split(':');
          const definedName = {
            name: '_xlnm.Print_Area',
            ranges: [`'${sheet.name}'!$${printAreaComponents[0]}:$${printAreaComponents[1]}`],
            localSheetId: index,
          };
          printAreas.push(definedName);
        });
      }

      if (pageSetup && (pageSetup.printTitlesRow || pageSetup.printTitlesColumn)) {
        const ranges: string[] = [];

        if (pageSetup.printTitlesColumn) {
          const titlesColumns = pageSetup.printTitlesColumn.split(':');
          ranges.push(`'${sheet.name}'!$${titlesColumns[0]}:$${titlesColumns[1]}`);
        }

        if (pageSetup.printTitlesRow) {
          const titlesRows = pageSetup.printTitlesRow.split(':');
          ranges.push(`'${sheet.name}'!$${titlesRows[0]}:$${titlesRows[1]}`);
        }

        const definedName = {
          name: '_xlnm.Print_Titles',
          ranges,
          localSheetId: index,
        };

        printAreas.push(definedName);
      }
      index++;
    });
    if (printAreas.length) {
      m.definedNames = (m.definedNames || []).concat(printAreas);
    }

    (m.media || []).forEach((medium, i) => {
      // assign name
      medium.name = medium.type + (i + 1);
    });
  }

  override render(xmlStream: XmlStreamLike, model?: WorkbookModel | null): void {
    xmlStream.openXml((XmlStream as {StdDocAttributes: Record<string, string>}).StdDocAttributes);
    xmlStream.openNode('workbook', WorkbookXform.WORKBOOK_ATTRIBUTES);

    this.map.fileVersion.render(xmlStream);
    this.map.workbookPr.render(xmlStream, model!.properties as never);
    this.map.bookViews.render(xmlStream, model!.views as never);
    this.map.sheets.render(xmlStream, model!.sheets as never);
    this.map.definedNames.render(xmlStream, model!.definedNames as never);
    this.map.calcPr.render(xmlStream, model!.calcProperties as never);
    this.map.pivotCaches.render(xmlStream, model!.pivotTables as never);

    xmlStream.closeNode();
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'workbook':
        return true;
      default:
        this.parser = this.map[node.name] as BaseXform;
        if (this.parser) {
          this.parser.parseOpen(node);
        }
        return true;
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
      case 'workbook':
        this.model = {
          sheets: this.map.sheets.model as unknown[],
          properties: (this.map.workbookPr.model as object) || {},
          views: this.map.bookViews.model as unknown[],
          calcProperties: {},
        };
        if (this.map.definedNames.model) {
          this.model.definedNames = this.map.definedNames.model as unknown[];
        }

        return false;
      default:
        // not quite sure how we get here!
        return true;
    }
  }

  override reconcile(model?: WorkbookModel | null): void {
    const m = model!;
    const rels = ((m.workbookRels || []) as {Id: string; Target: string}[]).reduce(
      (map: Record<string, {Id: string; Target: string}>, rel) => {
        map[rel.Id] = rel;
        return map;
      },
      {},
    );

    // reconcile sheet ids, rIds and names
    const worksheets: Record<string, unknown>[] = [];
    let worksheet: Record<string, unknown> | undefined;
    let index = 0;

    ((m.sheets || []) as {rId: string; name: string; id: number; state?: string}[]).forEach(
      sheet => {
        const rel = rels[sheet.rId];
        if (!rel) {
          return;
        }
        // if rel.Target start with `[space]/xl/` or `/xl/` , then it will be replaced with `''` and spliced behind `xl/`,
        // otherwise it will be spliced directly behind `xl/`. i.g.
        worksheet = m.worksheetHash![`xl/${rel.Target.replace(/^(\s|\/xl\/)+/, '')}`];
        // If there are "chartsheets" in the file, rel.Target will
        // come out as chartsheets/sheet1.xml or similar here, and
        // that won't be in model.worksheetHash.
        // As we don't have the infrastructure to support chartsheets,
        // we will ignore them for now:
        if (worksheet) {
          worksheet.name = sheet.name;
          worksheet.id = sheet.id;
          worksheet.state = sheet.state;
          worksheets[index++] = worksheet;
        }
      },
    );

    // reconcile print areas
    const definedNames: unknown[] = [];
    const definedNameList =
      (m.definedNames as {name: string; localSheetId: number; ranges: string[]}[] | undefined) ||
      [];
    for (const definedName of definedNameList) {
      if (definedName.name === '_xlnm.Print_Area') {
        worksheet = worksheets[definedName.localSheetId];
        if (worksheet) {
          if (!worksheet.pageSetup) {
            worksheet.pageSetup = {};
          }
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const range = colCache.decodeEx(definedName.ranges[0]) as any;
          const pageSetup = worksheet.pageSetup as {printArea?: string};
          pageSetup.printArea = pageSetup.printArea
            ? `${pageSetup.printArea}&&${range.dimensions}`
            : range.dimensions;
        }
      } else if (definedName.name === '_xlnm.Print_Titles') {
        worksheet = worksheets[definedName.localSheetId];
        if (worksheet) {
          if (!worksheet.pageSetup) {
            worksheet.pageSetup = {};
          }

          const rangeString = definedName.ranges.join(',');

          const dollarRegex = /\$/g;

          const rowRangeRegex = /\$\d+:\$\d+/;
          const rowRangeMatches = rangeString.match(rowRangeRegex);

          const pageSetup = worksheet.pageSetup as {
            printTitlesRow?: string;
            printTitlesColumn?: string;
          };

          if (rowRangeMatches && rowRangeMatches.length) {
            const range = rowRangeMatches[0];
            pageSetup.printTitlesRow = range.replace(dollarRegex, '');
          }

          const columnRangeRegex = /\$[A-Z]+:\$[A-Z]+/;
          const columnRangeMatches = rangeString.match(columnRangeRegex);

          if (columnRangeMatches && columnRangeMatches.length) {
            const range = columnRangeMatches[0];
            pageSetup.printTitlesColumn = range.replace(dollarRegex, '');
          }
        }
      } else {
        definedNames.push(definedName);
      }
    }
    m.definedNames = definedNames;

    // used by sheets to build their image models
    (m.media || []).forEach((media, i) => {
      media.index = i;
    });
  }

  static WORKBOOK_ATTRIBUTES: Record<string, string> = {
    xmlns: 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
    'xmlns:r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'xmlns:mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'mc:Ignorable': 'x15',
    'xmlns:x15': 'http://schemas.microsoft.com/office/spreadsheetml/2010/11/main',
  };

  static STATIC_XFORMS = {
    fileVersion: new StaticXform({
      tag: 'fileVersion',
      $: {appName: 'xl', lastEdited: 5, lowestEdited: 5, rupBuild: 9303},
    }),
  };
}

export default WorkbookXform;
export {WorkbookXform};
