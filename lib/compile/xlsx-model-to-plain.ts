/** Project reconciled OOXML models directly into the public snapshot. */
import type {
  Workbook,
  SheetModel,
  SheetRow,
  SheetCell,
  ColumnInput,
  NoteValue,
  MediaImage,
  TableProperties,
} from '../model/types.js';
import type {XlsxWorkbookModel} from './ops-to-xlsx-model.js';
import {pickStyle, metaFromModel} from '../model/snapshot-helpers.js';
import {valueFromCellModel} from '../model/cell-model.js';
import Enums from '../model/enums.js';
import DefinedNames from '../model/defined-names.js';
import Note from '../model/note.js';
import Range from '../model/range.js';
import colCache from '../utils/col-cache.js';
import Image from '../model/image.js';
import type {ImageModelInput} from '../model/image.js';
import type {CellValueModel} from '../model/cell-model.js';

export function xlsxModelToPlain(model: XlsxWorkbookModel): Workbook {
  const sheets: SheetModel[] = [];
  for (const worksheet of model.worksheets) {
    const merges = (worksheet.mergeCells ?? []).map(range => new Range(range));
    const rows: SheetRow[] = [];
    const notes: Record<string, NoteValue> = {};
    for (const input of worksheet.rows ?? []) {
      const cells: Record<number, SheetCell> = {};
      let previousColumn = 0;
      let hasValues = merges.some(
        m =>
          input.number >= m.top &&
          input.number <= m.bottom &&
          (input.number > m.top || m.right > m.left),
      );
      const rowNotes: Record<string, NoteValue> = {};
      for (const raw of input.cells) {
        const cell = raw as CellValueModel;
        if (cell.type === Enums.ValueType.Merge) continue;
        if (cell.type !== Enums.ValueType.Null) hasValues = true;
        const address = cell.address
          ? colCache.decodeAddress(cell.address)
          : colCache.decodeAddress(colCache.encodeAddress(input.number, previousColumn + 1));
        previousColumn = address.col!;
        // Hydration previously replaced merge slaves, even if the XML gave them values.
        const slave = merges.some(
          m =>
            address.row! >= m.top &&
            address.row! <= m.bottom &&
            address.col! >= m.left &&
            address.col! <= m.right &&
            (address.row !== m.top || address.col !== m.left),
        );
        if (slave) continue;
        if (cell.comment) {
          const note = Note.fromModel(cell.comment as never).note;
          if (note != null && note !== '') rowNotes[address.address] = note as NoteValue;
        }
        // Match the existing public projection: null cells are omitted, including styled nulls.
        if (cell.type === Enums.ValueType.Null) continue;
        const value = valueFromCellModel(cell) as SheetCell['value'];
        const style = pickStyle(cell.style);
        if (value == null && !style) continue;
        cells[address.col!] = style ? {value: value ?? null, style} : {value: value ?? null};
      }
      if (!hasValues) continue;
      Object.assign(notes, rowNotes);
      if (!Object.keys(cells).length && !input.height && !input.hidden) continue;
      const row: SheetRow = {number: input.number, cells};
      if (input.height != null) row.height = input.height;
      if (input.hidden) row.hidden = true;
      const style = pickStyle(input.style);
      if (style) row.style = style;
      rows.push(row);
    }
    const sheet: SheetModel = {id: worksheet.id!, name: worksheet.name!, rows};
    const columns: ColumnInput[] = [];
    let meaningful = 0;
    for (const col of worksheet.cols ?? []) {
      const style = pickStyle(col.style);
      const customWidth = col.width != null && col.width !== 9;
      for (let number = col.min; number <= col.max; number++) {
        while (columns.length < number) columns.push({});
        const entry: ColumnInput = {};
        if (customWidth) entry.width = col.width;
        if (col.hidden) entry.hidden = true;
        if (col.outlineLevel) entry.outlineLevel = col.outlineLevel;
        if (style) entry.style = style;
        columns[number - 1] = entry;
        if (Object.keys(entry).length) meaningful = number;
      }
    }
    if (meaningful) sheet.columns = columns.slice(0, meaningful);
    if (merges.length) sheet.merges = merges.map(m => m.range);
    if (worksheet.views?.length) sheet.views = worksheet.views;
    if (worksheet.pageSetup && Object.keys(worksheet.pageSetup).length)
      sheet.pageSetup = worksheet.pageSetup;
    if (worksheet.headerFooter && Object.keys(worksheet.headerFooter).length)
      sheet.headerFooter = worksheet.headerFooter;
    const validations = Object.fromEntries(
      Object.entries(worksheet.dataValidations ?? {}).filter(([, rules]) => rules),
    );
    if (Object.keys(validations).length) sheet.dataValidations = validations as never;
    if (worksheet.conditionalFormattings?.length)
      sheet.conditionalFormattings = worksheet.conditionalFormattings;
    if (Object.keys(notes).length) sheet.notes = notes;
    if (worksheet.sheetProtection) sheet.sheetProtection = {...worksheet.sheetProtection};
    if (worksheet.tables?.length)
      sheet.tables = worksheet.tables.map(table => {
        const m = table as unknown as TableProperties;
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
    const images = (worksheet.media ?? []).filter(
      (m): m is ImageModelInput => !!m && (m as {type?: string}).type === 'image',
    );
    if (images.length) {
      // Preserve the existing anchor coordinate accessors with a dimension-only
      // host. Anchors no longer keep an entire document workbook alive.
      const widths = new Map<number, number>();
      for (const col of worksheet.cols ?? []) {
        if (col.width != null && col.width !== 9) {
          for (let number = col.min; number <= col.max; number++) widths.set(number, col.width);
        }
      }
      const heights = new Map(
        (worksheet.rows ?? []).filter(row => row.height).map(row => [row.number, row.height]),
      );
      const host = {
        getColumn: (col: number) => ({width: widths.get(col), isCustomWidth: widths.has(col)}),
        getRow: (row: number) => ({height: heights.get(row)}),
      };
      sheet.images = images.map(model => ({
        imageId: model.imageId,
        range: new Image(host, model).range as never,
      }));
    }
    sheets.push(sheet);
  }
  const result: Workbook = {meta: metaFromModel(model), sheets};
  const names = new DefinedNames();
  names.model = model.definedNames as never;
  const entries = names.model.flatMap(d => d.ranges.map(refersTo => ({name: d.name, refersTo})));
  if (entries.length) result.definedNames = entries;
  const media = ((model.media as Array<MediaImage & {type?: string}>) ?? [])
    .filter(m => !m.type || m.type === 'image')
    .map(m => ({extension: m.extension, base64: m.base64, filename: m.filename, buffer: m.buffer}));
  if (media.length) result.media = media;
  return result;
}
