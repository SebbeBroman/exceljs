/**
 * Convert an internal DocWorkbook (post-load or materialize) into a plain Workbook snapshot.
 * Not part of the public API.
 */

import type {
  CellValue,
  ColumnInput,
  ConditionalFormattingOptions,
  DataValidation,
  DefinedNameEntry,
  HeaderFooter,
  MediaImage,
  NoteValue,
  PageSetup,
  SheetCell,
  SheetModel,
  SheetRow,
  Style,
  TableProperties,
  Workbook,
  WorkbookMeta,
  WorksheetViewInput,
} from '../model/types.js';
import type DocWorkbook from '../doc/workbook.js';
import Enums from '../doc/enums.js';
import colCache from '../utils/col-cache.js';

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyWs = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyRow = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyCell = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyCol = any;

function pickStyle(style: Record<string, unknown> | undefined | null): Style | undefined {
  if (!style || typeof style !== 'object') return undefined;
  const out: Style = {};
  if (style.numFmt != null) out.numFmt = style.numFmt as string;
  if (style.font) out.font = style.font as Style['font'];
  if (style.alignment) out.alignment = style.alignment as Style['alignment'];
  if (style.protection) out.protection = style.protection as Style['protection'];
  if (style.border) out.border = style.border as Style['border'];
  if (style.fill) out.fill = style.fill as Style['fill'];
  return Object.keys(out).length ? out : undefined;
}

function styleHasContent(style: Style | undefined): boolean {
  return Boolean(style && Object.keys(style).length);
}

function cellValueFromDoc(cell: AnyCell): CellValue {
  // Merge slaves share the master; omit them from the sparse plain model.
  if (cell.type === Enums.ValueType.Merge) {
    return null;
  }
  const v = cell.value as CellValue;
  if (v === undefined) return null;
  return v;
}

function columnsFromSheet(ws: AnyWs): ColumnInput[] | undefined {
  const cols: AnyCol[] | null | undefined = ws.columns;
  if (!cols || !cols.length) return undefined;

  let maxMeaningful = 0;
  const out: ColumnInput[] = [];
  for (let i = 0; i < cols.length; i++) {
    const col = cols[i];
    if (!col || col.isDefault) {
      out.push({});
      continue;
    }
    maxMeaningful = i + 1;
    const entry: ColumnInput = {};
    if (col.header != null) entry.header = col.header;
    if (col.key != null) entry.key = col.key;
    if (col.width != null && col.isCustomWidth) entry.width = col.width;
    if (col.hidden) entry.hidden = true;
    if (col.outlineLevel) entry.outlineLevel = col.outlineLevel;
    const st = pickStyle(col.style);
    if (st) entry.style = st;
    out.push(entry);
  }
  if (!maxMeaningful) return undefined;
  // trim trailing empty column defs
  return out.slice(0, maxMeaningful);
}

function mergesFromSheet(ws: AnyWs): string[] | undefined {
  const merges: Record<string, {range?: string; shortRange?: string}> | undefined = ws._merges;
  if (!merges) return undefined;
  const list = Object.values(merges)
    .map(m => m.range || m.shortRange)
    .filter((r): r is string => Boolean(r));
  return list.length ? list : undefined;
}

function metaFromDoc(wb: InstanceType<typeof DocWorkbook>): WorkbookMeta {
  const meta: WorkbookMeta = {};
  if (wb.creator != null && wb.creator !== '') meta.creator = wb.creator;
  if (wb.lastModifiedBy != null && wb.lastModifiedBy !== '') meta.lastModifiedBy = wb.lastModifiedBy;
  if (wb.created) meta.created = wb.created;
  if (wb.modified) meta.modified = wb.modified;
  if (wb.company) meta.company = wb.company;
  if (wb.manager) meta.manager = wb.manager;
  if (wb.title) meta.title = wb.title;
  if (wb.subject) meta.subject = wb.subject;
  if (wb.keywords) meta.keywords = wb.keywords;
  if (wb.category) meta.category = wb.category;
  if (wb.description) meta.description = wb.description;
  if (wb.language) meta.language = wb.language;
  if (wb.revision != null) meta.revision = wb.revision;
  if (wb.contentStatus) meta.contentStatus = wb.contentStatus;
  if (wb.properties && Object.keys(wb.properties).length) {
    meta.properties = wb.properties as WorkbookMeta['properties'];
  }
  if (wb.views && wb.views.length) meta.views = wb.views;
  return meta;
}

function dataValidationsFromSheet(ws: AnyWs): Record<string, DataValidation> | undefined {
  const model = ws.dataValidations?.model as Record<string, DataValidation | undefined> | undefined;
  if (!model) return undefined;
  const out: Record<string, DataValidation> = {};
  for (const [addr, rules] of Object.entries(model)) {
    if (rules) out[addr] = rules;
  }
  return Object.keys(out).length ? out : undefined;
}

function notesFromSheet(ws: AnyWs): Record<string, NoteValue> | undefined {
  const notes: Record<string, NoteValue> = {};
  (ws as AnyWs).eachRow({includeEmpty: false}, (row: AnyRow) => {
    row.eachCell({includeEmpty: false}, (cell: AnyCell) => {
      if (cell.note != null && cell.note !== '') {
        const address =
          cell.address ||
          `${colCache.n2l(cell.col as number)}${cell.row as number}`;
        notes[address] = cell.note as NoteValue;
      }
    });
  });
  return Object.keys(notes).length ? notes : undefined;
}

function tablesFromSheet(ws: AnyWs): TableProperties[] | undefined {
  const tables = ws.getTables?.() as Array<{model?: TableProperties} & TableProperties> | undefined;
  if (!tables?.length) return undefined;
  return tables.map(t => {
    const m = (t.model || t) as TableProperties;
    return {
      name: m.name,
      displayName: m.displayName,
      ref: m.ref,
      headerRow: m.headerRow,
      totalsRow: m.totalsRow,
      style: m.style,
      columns: m.columns,
      rows: m.rows,
    };
  });
}

function definedNamesFromDoc(wb: InstanceType<typeof DocWorkbook>): DefinedNameEntry[] | undefined {
  const model = wb.definedNames?.model;
  if (!model?.length) return undefined;
  const out: DefinedNameEntry[] = [];
  for (const dn of model) {
    if (!dn.name || !dn.ranges?.length) continue;
    for (const range of dn.ranges) {
      out.push({name: dn.name, refersTo: range});
    }
  }
  return out.length ? out : undefined;
}

function mediaFromDoc(wb: InstanceType<typeof DocWorkbook>): MediaImage[] | undefined {
  const media = wb.media;
  if (!media?.length) return undefined;
  const out: MediaImage[] = [];
  for (const m of media) {
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    const item = m as any;
    if (item.type && item.type !== 'image') continue;
    out.push({
      extension: item.extension,
      base64: item.base64,
      filename: item.filename,
      buffer: item.buffer,
    });
  }
  return out.length ? out : undefined;
}

/**
 * Project DocWorkbook → plain Workbook (cells, styles, merges, columns, meta,
 * plus Phase 5 sheet features when present).
 *
 * Tables / images: best-effort (write-oriented re-encode). Pivot not projected.
 */
export function docWorkbookToPlain(wb: InstanceType<typeof DocWorkbook>): Workbook {
  const sheets: SheetModel[] = [];

  for (const ws of wb.worksheets) {
    const rows: SheetRow[] = [];

    (ws as AnyWs).eachRow((row: AnyRow, rowNumber: number) => {
      const cells: Record<number, SheetCell> = {};
      let hasCell = false;

      row.eachCell((cell: AnyCell, colNumber: number) => {
        if (cell.type === Enums.ValueType.Merge) return;
        if (cell.type === Enums.ValueType.Null) return;

        const value = cellValueFromDoc(cell);
        // Skip pure nulls without style
        const style = pickStyle(cell.style);
        if ((value === null || value === undefined) && !styleHasContent(style)) return;

        cells[colNumber] = styleHasContent(style)
          ? {value: value ?? null, style}
          : {value: value ?? null};
        hasCell = true;
      });

      if (!hasCell && !row.height && !row.hidden) return;

      const sheetRow: SheetRow = {number: rowNumber, cells};
      if (row.height != null) sheetRow.height = row.height;
      if (row.hidden) sheetRow.hidden = true;
      const rowStyle = pickStyle(row.style);
      if (rowStyle) sheetRow.style = rowStyle;
      rows.push(sheetRow);
    });

    const sheet: SheetModel = {
      id: ws.id,
      name: ws.name,
      rows,
    };

    const columns = columnsFromSheet(ws);
    if (columns) sheet.columns = columns;

    const merges = mergesFromSheet(ws);
    if (merges) sheet.merges = merges;

    if (ws.views?.length) {
      sheet.views = ws.views as WorksheetViewInput[];
    }
    if (ws.pageSetup && Object.keys(ws.pageSetup).length) {
      sheet.pageSetup = ws.pageSetup as Partial<PageSetup>;
    }
    if (ws.headerFooter && Object.keys(ws.headerFooter).length) {
      // Drop empty defaults where possible
      sheet.headerFooter = ws.headerFooter as Partial<HeaderFooter>;
    }

    const dvs = dataValidationsFromSheet(ws);
    if (dvs) sheet.dataValidations = dvs;

    if (ws.conditionalFormattings?.length) {
      sheet.conditionalFormattings = ws.conditionalFormattings as ConditionalFormattingOptions[];
    }

    const notes = notesFromSheet(ws);
    if (notes) sheet.notes = notes;

    if (ws.sheetProtection) {
      sheet.sheetProtection = {...ws.sheetProtection} as Record<string, unknown>;
    }

    const tables = tablesFromSheet(ws);
    if (tables) sheet.tables = tables;

    // Sheet images: store placements if media exists (range may be opaque model)
    const images = (ws as AnyWs).getImages?.() as Array<{imageId: number | string; range: unknown}> | undefined;
    if (images?.length) {
      sheet.images = images.map(img => ({
        imageId: typeof img.imageId === 'string' ? Number(img.imageId) : img.imageId,
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        range: img.range as any,
      }));
    }

    sheets.push(sheet);
  }

  const result: Workbook = {
    meta: metaFromDoc(wb),
    sheets,
  };

  const media = mediaFromDoc(wb);
  if (media) result.media = media;

  const definedNames = definedNamesFromDoc(wb);
  if (definedNames) result.definedNames = definedNames;

  return result;
}
