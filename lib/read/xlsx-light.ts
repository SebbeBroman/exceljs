/**
 * Values-only / string-grid xlsx reader with early-exit row ranges.
 * Core engine for lazy viewWorkbook / readRows:
 * unzip once → workbook + sharedStrings → SAX parse sheet on demand.
 */

import {Parser} from 'saxen';
import type {CellValue} from '../model/types.js';
import colCache from '../utils/col-cache.js';
import utils from '../utils/utils.js';
import {entryToString, unzipToFiles} from '../utils/zip-reader.js';
import {cellToDisplayString} from './cells.js';

export interface LightSheetInfo {
  /** 0-based index in workbook sheet order */
  index: number;
  name: string;
  /** sheetId from workbook.xml (often 1-based) */
  id: number;
  /** Zip path e.g. "xl/worksheets/sheet1.xml" */
  path: string;
  state?: string;
}

/** Alias used by viewWorkbook lazy path. */
export type LightSheetRef = LightSheetInfo;

export interface LightPackage {
  /** Worksheet names only — chartsheets have no grid and are skipped. */
  sheetNames: string[];
  sheets: LightSheetInfo[];
  /** Shared string table; index → plain string (rich text flattened) */
  sharedStrings: string[];
  /** Raw zip entry bytes by path (normalized, no leading /) */
  files: Record<string, Uint8Array>;
}

export interface LightParseSheetOptions {
  /** 1-based inclusive start row. Default 1. */
  start?: number;
  /** 1-based inclusive end row. When set, SAX must stop after this row (early exit). */
  end?: number;
  /**
   * Column filter:
   * - `{ start, end }` 1-based inclusive contiguous
   * - `number[]` 1-based columns
   * - `string[]` column letters (accepted for view integrator)
   * - omit = all columns seen in parsed rows
   */
  cols?: {start?: number; end?: number} | number[] | string[];
  /** 'string' (default) → string[][]; 'cell' → CellValue[][] (no style) */
  values?: 'string' | 'cell';
  blankrows?: boolean; // default false = skip fully empty rows in range
  trim?: boolean; // default true for string mode
  defval?: string; // default ''
}

export interface LightSheetGrid {
  name: string;
  id: number;
  index: number;
  /** Dense matrix for the requested slice (not full sheet if end set) */
  rows: string[][] | CellValue[][];
  /** Max column index seen in this parse (1-based), useful for extent */
  maxCol: number;
  /** Last row number included (1-based), or 0 if empty */
  lastRow: number;
}

/** Thrown from SAX callbacks to abort the rest of the sheet XML. */
class StopParse extends Error {
  constructor() {
    super('STOP_PARSE');
    this.name = 'StopParse';
  }
}

function isStopParse(err: unknown): boolean {
  return err instanceof StopParse || (err as {name?: string})?.name === 'StopParse';
}

function normalizePath(p: string): string {
  return p.replace(/^\/+/, '');
}

function attrOf(attrs: string, name: string): string | undefined {
  // Escape regex metacharacters in attribute name (e.g. r:id is fine; keep general).
  const esc = name.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const re = new RegExp(`(?:^|\\s)${esc}\\s*=\\s*(?:"([^"]*)"|'([^']*)')`, 'i');
  const m = attrs.match(re);
  if (!m) return undefined;
  return m[1] ?? m[2];
}

/** Resolve workbook.xml.rels Target → zip path under xl/. Chartsheets → null. */
function resolveWorksheetPath(target: string): string | null {
  let t = target.trim().replace(/\\/g, '/');
  t = t.replace(/^\/+/, '');
  t = t.replace(/^(\.\/)+/, '');
  if (t.startsWith('xl/')) t = t.slice(3);
  if (t.startsWith('chartsheets/')) return null;
  if (!t.startsWith('worksheets/')) return null;
  return `xl/${t}`;
}

/** Scan for `<tag ...>` open tags, respecting quoted `>` inside attribute values. */
function scanTagAttrs(xml: string, tag: string): string[] {
  const out: string[] = [];
  const lower = xml.toLowerCase();
  const needle = `<${tag.toLowerCase()}`;
  let pos = 0;
  while (true) {
    const start = lower.indexOf(needle, pos);
    if (start === -1) break;
    // Avoid `</tag`, `<tagData` etc: next char must terminate the tag name.
    const after = xml[start + tag.length + 1] ?? '';
    if (after === '/' || /[A-Za-z0-9_:.\-]/.test(after)) {
      pos = start + needle.length;
      continue;
    }
    // Walk to the unquoted `>`.
    let i = start + needle.length;
    let quote: string | null = null;
    let end = -1;
    for (; i < xml.length; i++) {
      const ch = xml[i]!;
      if (quote) {
        if (ch === quote) quote = null;
      } else if (ch === '"' || ch === "'") {
        quote = ch;
      } else if (ch === '>') {
        end = i;
        break;
      }
    }
    if (end === -1) break;
    out.push(xml.slice(start + needle.length, end));
    pos = end + 1;
  }
  return out;
}

function parseWorkbookSheets(xml: string): Array<{
  name: string;
  id: number;
  rId: string;
  state?: string;
}> {
  const out: Array<{name: string; id: number; rId: string; state?: string}> = [];
  // Quote-aware scan: sheet names may legally contain `>` which breaks `[^>]*`.
  for (const attrs of scanTagAttrs(xml, 'sheet')) {
    const nameRaw = attrOf(attrs, 'name');
    const rId = attrOf(attrs, 'r:id') ?? attrOf(attrs, 'id');
    if (!nameRaw || !rId) continue;
    const sheetId = attrOf(attrs, 'sheetId');
    const state = attrOf(attrs, 'state');
    const entry: {name: string; id: number; rId: string; state?: string} = {
      name: utils.xmlDecode(nameRaw),
      id: sheetId ? parseInt(sheetId, 10) : out.length + 1,
      rId,
    };
    if (state) entry.state = state;
    out.push(entry);
  }
  return out;
}

function parseWorkbookRels(xml: string): Map<string, string> {
  const map = new Map<string, string>();
  for (const attrs of scanTagAttrs(xml, 'Relationship')) {
    const id = attrOf(attrs, 'Id');
    const target = attrOf(attrs, 'Target');
    if (id && target) map.set(id, target);
  }
  return map;
}

/** Flatten sharedStrings.xml → string[] (join all <t>; ignore rPh). */
function parseSharedStringsXml(xml: string): string[] {
  const values: string[] = [];
  const parser = new Parser();
  let error: Error | undefined;
  let inSi = false;
  let inT = false;
  let inRPh = false;
  let buf = '';

  const fail = (err: unknown): void => {
    if (!error) error = err instanceof Error ? err : new Error(String(err));
  };
  parser.on('error', fail);
  parser.on('warn', fail);

  parser.on('openTag', name => {
    if (name === 'si') {
      inSi = true;
      buf = '';
      inRPh = false;
      inT = false;
    } else if (name === 'rPh') {
      inRPh = true;
    } else if (name === 't' && inSi && !inRPh) {
      inT = true;
    }
  });

  parser.on('text', (value, decodeEntities) => {
    if (inT && inSi && !inRPh) {
      buf += value.indexOf('&') === -1 ? value : decodeEntities(value);
    }
  });

  parser.on('closeTag', name => {
    if (name === 't') {
      inT = false;
    } else if (name === 'rPh') {
      inRPh = false;
    } else if (name === 'si') {
      values.push(buf);
      buf = '';
      inSi = false;
      inT = false;
      inRPh = false;
    }
  });

  parser.write(xml);
  parser.end();
  if (error) throw error;
  return values;
}

/** Unzip + parse workbook.xml + workbook rels + sharedStrings. Do NOT parse any sheet XML. */
export async function openLightPackage(
  data: Uint8Array | ArrayBuffer | ArrayBufferView,
): Promise<LightPackage> {
  const raw = await unzipToFiles(data);
  const files: Record<string, Uint8Array> = Object.create(null);
  for (const key of Object.keys(raw)) {
    if (!key || key.endsWith('/')) continue;
    files[normalizePath(key)] = raw[key]!;
  }

  const wbBytes = files['xl/workbook.xml'];
  if (!wbBytes) {
    throw new Error('Invalid xlsx: missing xl/workbook.xml');
  }

  const sheetMetas = parseWorkbookSheets(entryToString(wbBytes));
  const relsBytes = files['xl/_rels/workbook.xml.rels'];
  const rels = relsBytes ? parseWorkbookRels(entryToString(relsBytes)) : new Map<string, string>();

  const sheets: LightSheetInfo[] = [];
  for (const meta of sheetMetas) {
    const target = rels.get(meta.rId);
    let path: string | null = null;
    if (target) {
      path = resolveWorksheetPath(target);
    } else {
      // Fallback when rels missing/incomplete
      path = `xl/worksheets/sheet${sheets.length + 1}.xml`;
      if (!files[path]) path = null;
    }
    if (!path) continue; // chartsheet or unresolved
    if (!files[path]) continue;
    const info: LightSheetInfo = {
      index: sheets.length,
      name: meta.name,
      id: meta.id,
      path,
    };
    if (meta.state) info.state = meta.state;
    sheets.push(info);
  }

  let sharedStrings: string[] = [];
  const sst = files['xl/sharedStrings.xml'];
  if (sst) {
    sharedStrings = parseSharedStringsXml(entryToString(sst));
  }

  return {
    sheetNames: sheets.map(s => s.name),
    sheets,
    sharedStrings,
    files,
  };
}

function colFromAddress(address: string): number {
  let col = 0;
  for (let i = 0; i < address.length; i++) {
    const ch = address.charCodeAt(i);
    if (ch >= 65 && ch <= 90) {
      col = col * 26 + (ch - 64);
    } else if (ch >= 97 && ch <= 122) {
      col = col * 26 + (ch - 96);
    } else {
      break;
    }
  }
  if (col > 0) return col;
  return (colCache.decodeAddress(address).col as number) || 0;
}

function normalizeColList(cols: LightParseSheetOptions['cols']): number[] | {start?: number; end?: number} | null {
  if (cols == null) return null;
  if (Array.isArray(cols)) {
    if (cols.length === 0) return [];
    if (typeof cols[0] === 'string') {
      return (cols as string[])
        .map(letter => {
          const s = letter.trim().toUpperCase();
          if (/^\d+$/.test(s)) return Number(s);
          return colCache.l2n(s);
        })
        .filter(c => c >= 1);
    }
    return (cols as number[]).filter(c => c >= 1);
  }
  return cols;
}

function resolveOutCols(
  cols: LightParseSheetOptions['cols'],
  maxCol: number,
): number[] {
  const normalized = normalizeColList(cols);
  if (normalized == null) {
    const out: number[] = [];
    for (let c = 1; c <= maxCol; c++) out.push(c);
    return out;
  }
  if (Array.isArray(normalized)) {
    return normalized;
  }
  const start = normalized.start ?? 1;
  const end = normalized.end ?? maxCol;
  const out: number[] = [];
  for (let c = start; c <= end; c++) out.push(c);
  return out;
}

function colAllowed(col: number, cols: LightParseSheetOptions['cols']): boolean {
  const normalized = normalizeColList(cols);
  if (normalized == null) return true;
  if (Array.isArray(normalized)) return normalized.includes(col);
  const start = normalized.start ?? 1;
  const end = normalized.end ?? Number.POSITIVE_INFINITY;
  return col >= start && col <= end;
}

interface SparseRow {
  number: number;
  cells: Map<number, CellValue>;
}

/** Parse one sheet from an already-opened package (lazy). */
export function parseLightSheet(
  pkg: LightPackage,
  sheet: string | number,
  options?: LightParseSheetOptions,
): LightSheetGrid {
  const opts = options ?? {};
  const asString = opts.values !== 'cell';
  const trim = opts.trim !== false && asString;
  const defval = opts.defval ?? '';
  const skipBlank = opts.blankrows !== true;
  const rowStart = opts.start ?? 1;
  const rowEnd = opts.end; // undefined = unlimited
  const colsOpt = opts.cols;

  let index: number;
  let info: LightSheetInfo | undefined;
  if (typeof sheet === 'number') {
    index = sheet;
    info = pkg.sheets[sheet];
  } else {
    index = pkg.sheets.findIndex(s => s.name === sheet);
    info = index >= 0 ? pkg.sheets[index] : undefined;
  }
  if (!info) {
    throw new Error(`Sheet not found: ${String(sheet)}`);
  }

  const bytes = pkg.files[info.path];
  if (!bytes) {
    throw new Error(`Missing sheet part: ${info.path}`);
  }

  const sharedStrings = pkg.sharedStrings;
  const sparseRows: SparseRow[] = [];
  let maxCol = 0;
  let lastRow = 0;

  // Pre-seed maxCol when cols filter has a known end / discrete list
  {
    const n = normalizeColList(colsOpt);
    if (n && Array.isArray(n)) {
      for (const c of n) {
        if (c > maxCol) maxCol = c;
      }
    } else if (n && !Array.isArray(n) && n.end != null) {
      maxCol = n.end;
    }
  }

  let inSheetData = false;
  let currentRow: SparseRow | null = null;
  let skipRow = false;
  let prevRowNum = 0;

  let cellCol = 0;
  let cellType = '';
  let cellText = '';
  let inV = false;
  let inIs = false;
  let inT = false;
  let hasValue = false;

  const parser = new Parser();
  let error: Error | undefined;
  const fail = (err: unknown): void => {
    if (isStopParse(err)) return;
    if (!error) error = err instanceof Error ? err : new Error(String(err));
  };
  parser.on('error', fail);
  parser.on('warn', fail);

  const finishCell = (): void => {
    if (!currentRow || skipRow || cellCol < 1) return;
    if (!colAllowed(cellCol, colsOpt)) return;

    // Values path: use cached <v> / inlineStr only (ignore formula text).
    let value: CellValue = null;
    if (hasValue || cellType === 'inlineStr') {
      switch (cellType) {
        case 's': {
          const idx = parseInt(cellText, 10);
          value = Number.isFinite(idx) ? (sharedStrings[idx] ?? '') : '';
          break;
        }
        case 'inlineStr':
        case 'str':
          value = cellText;
          break;
        case 'b':
          value = cellText === '1' || cellText === 'true';
          break;
        case 'e':
          value = {
            error: cellText as
              | '#N/A'
              | '#REF!'
              | '#NAME?'
              | '#DIV/0!'
              | '#NULL!'
              | '#VALUE!'
              | '#NUM!',
          };
          break;
        default: {
          if (cellText === '') {
            value = null;
          } else {
            const n = parseFloat(cellText);
            value = Number.isFinite(n) ? n : cellText;
          }
          break;
        }
      }
    } else {
      return; // empty / formula-only with no cached value
    }

    if (value === null && cellType !== 'b') return;

    if (cellCol > maxCol) maxCol = cellCol;
    currentRow.cells.set(cellCol, value);
  };

  parser.on('openTag', (name, getAttrs) => {
    if (name === 'sheetData') {
      inSheetData = true;
      return;
    }
    if (!inSheetData) return;

    if (name === 'row') {
      const attrs = getAttrs();
      let r: number;
      if (attrs.r != null && attrs.r !== '') {
        r = parseInt(attrs.r, 10) || prevRowNum + 1;
      } else {
        r = prevRowNum + 1;
      }
      prevRowNum = r;

      if (rowEnd != null && r > rowEnd) {
        throw new StopParse();
      }

      skipRow = r < rowStart;
      if (skipRow) {
        currentRow = null;
      } else {
        currentRow = {number: r, cells: new Map()};
        if (r > lastRow) lastRow = r;
      }
      return;
    }

    if (name === 'c' && currentRow && !skipRow) {
      const attrs = getAttrs();
      cellType = attrs.t || '';
      cellText = '';
      hasValue = false;
      inV = false;
      inIs = false;
      inT = false;
      if (attrs.r) {
        cellCol = colFromAddress(attrs.r);
      } else {
        cellCol = (cellCol || 0) + 1;
      }
      return;
    }

    if (!currentRow || skipRow) return;

    if (name === 'v') {
      inV = true;
      cellText = '';
      return;
    }
    if (name === 'is') {
      inIs = true;
      cellType = cellType || 'inlineStr';
      cellText = '';
      return;
    }
    if (name === 't' && inIs) {
      inT = true;
      return;
    }
  });

  parser.on('text', (value, decodeEntities) => {
    if (skipRow || !currentRow) return;
    const text = value.indexOf('&') === -1 ? value : decodeEntities(value);
    if (inV) {
      cellText += text;
      hasValue = true;
    } else if (inT && inIs) {
      cellText += text;
      hasValue = true;
    }
  });

  parser.on('closeTag', name => {
    if (name === 'sheetData') {
      inSheetData = false;
      // No more rows — stop so we don't scan pageMargins / drawings etc.
      throw new StopParse();
    }
    if (!inSheetData) return;

    if (name === 'v') {
      inV = false;
      return;
    }
    if (name === 't') {
      inT = false;
      return;
    }
    if (name === 'is') {
      inIs = false;
      return;
    }
    if (name === 'c') {
      finishCell();
      inV = false;
      inIs = false;
      inT = false;
      return;
    }
    if (name === 'row') {
      if (currentRow && !skipRow) {
        sparseRows.push(currentRow);
      }
      const finished = currentRow?.number ?? prevRowNum;
      currentRow = null;
      skipRow = false;
      // Early exit once the last requested row is fully closed.
      if (rowEnd != null && finished >= rowEnd) {
        throw new StopParse();
      }
    }
  });

  const xml = entryToString(bytes);
  try {
    parser.write(xml);
    parser.end();
  } catch (err) {
    if (!isStopParse(err)) throw err;
  }
  if (error) throw error;

  const outCols = resolveOutCols(colsOpt, maxCol);
  if (outCols.length === 0 || (sparseRows.length === 0 && maxCol < 1)) {
    return {
      name: info.name,
      id: info.id,
      index: info.index,
      rows: [],
      maxCol: 0,
      lastRow: 0,
    };
  }

  const byNumber = new Map(sparseRows.map(r => [r.number, r]));
  const emptyStringLine = (): string[] => outCols.map(() => defval);
  const emptyCellLine = (): CellValue[] => outCols.map(() => null);

  const buildStringLine = (sr: SparseRow | undefined): string[] => {
    if (!sr) return emptyStringLine();
    const line: string[] = new Array(outCols.length);
    for (let i = 0; i < outCols.length; i++) {
      const raw = sr.cells.get(outCols[i]!);
      let s = raw !== undefined ? cellToDisplayString(raw, defval) : defval;
      if (trim) s = s.trim();
      line[i] = s;
    }
    return line;
  };

  const buildCellLine = (sr: SparseRow | undefined): CellValue[] => {
    if (!sr) return emptyCellLine();
    const line: CellValue[] = new Array(outCols.length);
    for (let i = 0; i < outCols.length; i++) {
      line[i] = sr.cells.has(outCols[i]!) ? (sr.cells.get(outCols[i]!) ?? null) : null;
    }
    return line;
  };

  let emittedLast = 0;
  const reportMaxCol = maxCol || (outCols.length ? outCols[outCols.length - 1]! : 0);

  if (skipBlank) {
    // Only rows present in sheet XML (and not fully empty after densify).
    if (asString) {
      const rows: string[][] = [];
      for (const sr of sparseRows) {
        if (sr.number < rowStart) continue;
        if (rowEnd != null && sr.number > rowEnd) break;
        const line = buildStringLine(sr);
        if (line.every(v => v === defval)) continue;
        rows.push(line);
        emittedLast = sr.number;
      }
      return {
        name: info.name,
        id: info.id,
        index: info.index,
        rows,
        maxCol: reportMaxCol,
        lastRow: emittedLast,
      };
    }
    const rows: CellValue[][] = [];
    for (const sr of sparseRows) {
      if (sr.number < rowStart) continue;
      if (rowEnd != null && sr.number > rowEnd) break;
      const line = buildCellLine(sr);
      if (line.every(v => v == null || v === '')) continue;
      rows.push(line);
      emittedLast = sr.number;
    }
    return {
      name: info.name,
      id: info.id,
      index: info.index,
      rows,
      maxCol: reportMaxCol,
      lastRow: emittedLast,
    };
  }

  // blankrows: true — gap-fill absolute row numbers in [rowStart, last].
  const rangeLast =
    rowEnd != null ? Math.min(rowEnd, lastRow || rowEnd) : lastRow;
  if (rangeLast < rowStart) {
    return {
      name: info.name,
      id: info.id,
      index: info.index,
      rows: [],
      maxCol: reportMaxCol,
      lastRow: 0,
    };
  }

  if (asString) {
    const rows: string[][] = [];
    for (let r = rowStart; r <= rangeLast; r++) {
      rows.push(buildStringLine(byNumber.get(r)));
    }
    return {
      name: info.name,
      id: info.id,
      index: info.index,
      rows,
      maxCol: reportMaxCol,
      lastRow: rangeLast,
    };
  }

  const rows: CellValue[][] = [];
  for (let r = rowStart; r <= rangeLast; r++) {
    rows.push(buildCellLine(byNumber.get(r)));
  }
  return {
    name: info.name,
    id: info.id,
    index: info.index,
    rows,
    maxCol: reportMaxCol,
    lastRow: rangeLast,
  };
}

/** Convenience: open + parse one sheet. */
export async function readLightSheet(
  data: Uint8Array | ArrayBuffer | ArrayBufferView,
  sheet: string | number = 0,
  options?: LightParseSheetOptions,
): Promise<LightSheetGrid> {
  const pkg = await openLightPackage(data);
  return parseLightSheet(pkg, sheet, options);
}
