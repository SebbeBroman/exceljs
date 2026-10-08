/**
 * Fused fast path for worksheet `<sheetData>` on load.
 *
 * The classic pipeline parses sheet XML into per-cell xform models
 * (parseSax event objects + CellXformModel allocs), reconciles every cell
 * (style lookup, shared-string resolve, hyperlink/comment probes), then
 * projects plain snapshot rows (another per-cell model copy). For dense sheets
 * that is the bulk of `load()` time.
 *
 * This module fuses all three stages, emitting *reconciled* row models —
 * same shapes and semantics as RowXform/CellXform parse + reconcile
 * (style/date/shared-string/formula/hyperlink/comment handling) — which
 * hydrate through the unchanged `Row.model` setter. The caller still parses
 * sheet head/tail (cols, merges, validations, CF, page setup, …) with
 * WorksheetXform; only `<sheetData>` content is fused, and
 * WorksheetXform.reconcile skips `sheetData` for flagged models. (An
 * indexOf-driven Tier 1 was tried and dropped: per-tag dispatch overhead
 * beat none of saxen's single pass. The saxen tier below is the fast path.)
 *
 * Fidelity rule: anything not implemented forces the classic path via
 * `canUseFastSheetData()`. When in doubt, fall back.
 */

import {Parser} from 'saxen';
import Enums from '../../../model/enums.js';
import colCache from '../../../utils/col-cache.js';
import {excelToDate, isDateFmt, parseBoolean, xmlDecode} from '../../../utils/utils.js';
import type {CellXformModel} from './cell-xform.js';
import type {RowXformModel} from './row-xform.js';

/** Test/diagnostic kill-switch (default on). Not public API. */
let fastEnabled = true;
export function setFastSheetDataEnabled(value: boolean): void {
  fastEnabled = value;
}
export function isFastSheetDataEnabled(): boolean {
  return fastEnabled;
}

export interface FastSheetContext {
  /** Parsed shared-strings table (may be undefined when the part is absent). */
  sharedStrings?: {getString(index: number): unknown};
  /** Style model lookup; must behave like StylesXform.getStyleModel. */
  getStyleModel(id: number): Record<string, unknown> | null | undefined;
  /** Row-style lookup; must behave like row reconcile (throws without styles). */
  getRowStyleModel(id: number): unknown;
  date1904?: boolean;
  hyperlinkMap?: Record<string, string>;
  commentsMap?: Record<string, unknown>;
  formulae?: Record<string | number, unknown>;
}

/** Split sheet XML at `<sheetData>`; null when the element is missing/malformed. */
export function splitSheetData(xml: string): {head: string; content: string; tail: string} | null {
  const tag = '<sheetData';
  const start = xml.indexOf(tag);
  if (start === -1) return null;
  // The match must terminate the tag name (not e.g. `<sheetDatas…`).
  const after = xml[start + tag.length] ?? '';
  if (after !== '' && /[A-Za-z0-9_:.-]/.test(after)) return null;
  const openEnd = findTagEnd(xml, start + tag.length);
  if (openEnd === -1) return null;
  const closeTag = '</sheetData>';
  const closeStart = xml.indexOf(closeTag, openEnd + 1);
  if (closeStart === -1) return null;
  const content = xml.slice(openEnd + 1, closeStart);
  // Nested sheetData would break the split — let the classic path handle it.
  if (content.includes(tag)) return null;
  return {
    head: xml.slice(0, start),
    content,
    tail: xml.slice(closeStart + closeTag.length),
  };
}

/** Offset of the unquoted `>` terminating the tag opened before `from`. */
function findTagEnd(xml: string, from: number): number {
  let quote: string | null = null;
  for (let i = from; i < xml.length; i++) {
    const ch = xml[i]!;
    if (quote) {
      if (ch === quote) quote = null;
    } else if (ch === '"' || ch === "'") {
      quote = ch;
    } else if (ch === '>') {
      return i;
    }
  }
  return -1;
}

/**
 * True when sheetData content uses only constructs the fused parsers
 * implement. Denylisted: run properties (per-run fonts), phonetic runs,
 * and extension elements — all vanishingly rare inside sheetData.
 *
 * Single regex pass with early exit (four separate `includes` scans cost
 * ~5ms on a 1.4MB sheet).
 */
const fastDeny = /<rPr|<rPh|<extLst|<AlternateContent/;
export function canUseFastSheetData(content: string): boolean {
  return !fastDeny.test(content);
}

/** Same `_xNNNN_` control-char unescape as TextXform.model. */
const controlEscape = /_x([0-9A-F]{4})_/g;
function decodeControlEscapes(text: string): string {
  return text.indexOf('_x') === -1
    ? text
    : text.replace(controlEscape, ($0, $1: string) => String.fromCharCode(parseInt($1, 16)));
}

/**
 * Column number for plain `AB12` addresses (the overwhelmingly common case).
 * Returns undefined for anything else so the caller falls back to
 * colCache.decodeAddress — preserving its quirks exactly (lowercase →
 * undefined col, >16384 → throw).
 */
function fastColOf(raw: string): number | undefined {
  let col = 0;
  let i = 0;
  for (; i < raw.length; i++) {
    const code = raw.charCodeAt(i);
    if (code < 65 || code > 90) break;
    col = col * 26 + (code - 64);
    if (col > 16384) return undefined;
  }
  if (col === 0 || i === raw.length) return undefined;
  for (let j = i; j < raw.length; j++) {
    const code = raw.charCodeAt(j);
    if (code < 48 || code > 57) return undefined;
  }
  return col;
}

// ---------------------------------------------------------------------------
// Shared finalizers (used by both fused tiers)
// ---------------------------------------------------------------------------

interface SharedFinalize {
  ctx: FastSheetContext;
  styleCache: Array<Record<string, unknown> | null | undefined>;
  styleSeen: boolean[];
  dateCache: boolean[];
  formulae: Record<string | number, unknown>;
  needRelMaps: boolean;
}

function createShared(ctx: FastSheetContext): SharedFinalize {
  const needRelMaps = Boolean(
    (ctx.hyperlinkMap && Object.keys(ctx.hyperlinkMap).length > 0) ||
    (ctx.commentsMap && Object.keys(ctx.commentsMap).length > 0),
  );
  return {
    ctx,
    styleCache: [],
    styleSeen: [],
    dateCache: [],
    formulae: ctx.formulae ?? {},
    needRelMaps,
  };
}

function cachedStyle(
  shared: SharedFinalize,
  id: number,
): Record<string, unknown> | null | undefined {
  if (!shared.styleSeen[id]) {
    shared.styleSeen[id] = true;
    shared.styleCache[id] = shared.ctx.getStyleModel(id);
  }
  return shared.styleCache[id];
}

function cachedIsDate(
  shared: SharedFinalize,
  id: number,
  style: Record<string, unknown> | null | undefined,
): boolean {
  if (shared.dateCache[id] === undefined) {
    shared.dateCache[id] = !!style && isDateFmt((style as {numFmt?: string}).numFmt);
  }
  return shared.dateCache[id]!;
}

interface MutableRow {
  number: number;
  min?: number;
  max?: number;
  cells: CellXformModel[];
  height?: number;
  hidden?: boolean;
  bestFit?: boolean;
  styleId?: number;
  outlineLevel?: number;
  collapsed?: boolean;
  lastCol: number;
}

/** Plain cell fields; both tiers build this, then finalize into `out`. */
interface FlatCell {
  raw?: string;
  t: string;
  styleId?: number;
  formula?: string;
  si?: string;
  shareType?: string;
  ref?: string;
  /** Single value slot mirroring CellXformModel.value (string | {richText} | undefined). */
  val: string | {richText: Array<{text?: string}>} | undefined;
}

/**
 * Reconciled cell finalizer — mirrors CellXform.parseClose + reconcile.
 * Merge fillers (no value, no style) return false and are skipped by hydrate.
 */
function finalizeCell(row: MutableRow, cell: FlatCell, shared: SharedFinalize): boolean {
  const {ctx, formulae} = shared;
  const {raw, t, styleId} = cell;
  // Resolve the column: explicit `r` wins, else previous+1 within the row.
  // (Row.model infers the same way; a missing `r` on the first cell infers
  // col 1 where the classic path would throw — strictly more robust.)
  let col: number;
  if (raw !== undefined) {
    const fast = fastColOf(raw);
    col = fast ?? (colCache.decodeAddress(raw).col as number);
    if (!col) col = row.lastCol + 1;
  } else {
    col = row.lastCol + 1;
  }
  row.lastCol = col;
  const address = raw ?? colCache.encodeAddress(row.number, col);

  const out: CellXformModel = {address} as CellXformModel;
  const isFormula = cell.formula !== undefined || cell.shareType !== undefined;
  if (isFormula) {
    out.type = Enums.ValueType.Formula;
    if (cell.formula !== undefined) out.formula = cell.formula;
    if (cell.shareType !== undefined) out.shareType = cell.shareType;
    if (cell.si !== undefined) out.si = cell.si;
    if (cell.ref !== undefined) out.ref = cell.ref;
    if (cell.val !== undefined) {
      const v = cell.val as string;
      if (t === 'str') out.result = xmlDecode(v);
      else if (t === 'b') out.result = parseInt(v, 10) !== 0;
      else if (t === 'e') out.result = {error: v};
      else out.result = parseFloat(v);
    }
    // Shared-formula master/slave resolution (mirrors cell reconcile).
    if (out.shareType === 'shared') {
      if (out.ref) {
        formulae[out.si as string] = raw;
      } else {
        out.sharedFormula = formulae[out.si as string] as string;
        delete out.shareType;
        delete out.si;
      }
    }
  } else if (cell.val !== undefined) {
    const v = cell.val;
    switch (t) {
      case 's':
        out.type = Enums.ValueType.String;
        out.value = parseInt(v as string, 10);
        break;
      case 'str':
        out.type = Enums.ValueType.String;
        out.value = xmlDecode(v as string);
        break;
      case 'inlineStr':
        out.type = Enums.ValueType.String;
        out.value = v;
        break;
      case 'b':
        out.type = Enums.ValueType.Boolean;
        out.value = parseInt(v as string, 10) !== 0;
        break;
      case 'e':
        out.type = Enums.ValueType.Error;
        out.value = {error: v};
        break;
      default:
        out.type = Enums.ValueType.Number;
        out.value = parseFloat(v as string);
        break;
    }
  } else if (styleId !== undefined) {
    out.type = Enums.ValueType.Null;
  } else {
    return false;
  }

  // --- fused reconcile: style ---
  let style: Record<string, unknown> | null | undefined;
  if (styleId !== undefined) {
    style = cachedStyle(shared, styleId);
    if (style) out.style = style;
  }
  // --- shared strings / rich text ---
  const {sharedStrings, date1904, hyperlinkMap, commentsMap} = ctx;
  if (out.type === Enums.ValueType.String && typeof out.value === 'number') {
    if (sharedStrings) out.value = sharedStrings.getString(out.value);
  }
  if (out.value && typeof out.value === 'object' && (out.value as {richText?: unknown}).richText) {
    out.type = Enums.ValueType.RichText;
  }
  // --- dates ---
  if (
    out.type === Enums.ValueType.Number &&
    style &&
    cachedIsDate(shared, styleId as number, style)
  ) {
    out.type = Enums.ValueType.Date;
    out.value = excelToDate(out.value as number, date1904);
  }
  if (out.type === Enums.ValueType.Formula && out.result !== undefined && style) {
    if (cachedIsDate(shared, styleId as number, style)) {
      out.result = excelToDate(out.result as number, date1904);
    }
  }
  // --- hyperlinks / comments (raw address, exactly like cell reconcile) ---
  if (shared.needRelMaps) {
    const hyperlink = hyperlinkMap?.[raw as string];
    if (hyperlink) {
      if (out.type === Enums.ValueType.Formula) {
        out.text = out.result;
        out.result = undefined;
      } else {
        out.text = out.value;
        out.value = undefined;
      }
      out.type = Enums.ValueType.Hyperlink;
      out.hyperlink = hyperlink;
    }
    const comment = commentsMap?.[raw as string];
    if (comment) out.comment = comment;
  }

  row.cells.push(out);
  return true;
}

/** Row finalizer — mirrors RowXform parse + reconcile (style rule included). */
function finalizeRow(
  number: number,
  min: number | undefined,
  max: number | undefined,
  cells: CellXformModel[],
  opts: {
    height?: number;
    hidden?: boolean;
    bestFit?: boolean;
    styleId?: number;
    outlineLevel?: number;
    collapsed?: boolean;
  },
  ctx: FastSheetContext,
): RowXformModel {
  // Row style rule mirrors row reconcile: truthy styleId resolves, else {}.
  // (Like the classic path, this throws when styles are absent but used.)
  const style = opts.styleId ? ctx.getRowStyleModel(opts.styleId) : {};
  return {
    number,
    min,
    max,
    cells,
    ...(opts.height !== undefined ? {height: opts.height} : {}),
    ...(opts.hidden ? {hidden: true} : {}),
    ...(opts.bestFit ? {bestFit: true} : {}),
    style,
    ...(opts.outlineLevel !== undefined ? {outlineLevel: opts.outlineLevel} : {}),
    ...(opts.collapsed ? {collapsed: true} : {}),
  } as RowXformModel;
}

// ---------------------------------------------------------------------------
// Direct saxen pass (formulas, inline strings, rich text, row/cell styles, …)
// ---------------------------------------------------------------------------

type DecodeEntities = (s: string) => string;

interface PendingCell extends FlatCell {
  /** Mirrors CellXform.currentNode for f/v/t routing. */
  cur?: 'f' | 'v' | 't';
  inIs: boolean;
  inRun: boolean;
  inRunT: boolean;
  runText?: string;
  runSeen: boolean;
  runs?: Array<{text?: string}>;
}

/**
 * Parse + reconcile sheetData content into reconciled row models.
 * Output feeds the unchanged `Row.model` setter (hydrate).
 */
export function parseFastSheetData(content: string, ctx: FastSheetContext): RowXformModel[] {
  const shared = createShared(ctx);

  const rows: RowXformModel[] = [];
  let row: MutableRow | null = null;
  let cell: PendingCell | null = null;
  let error: Error | undefined;
  let stopped = false;

  const fail = (err: unknown): void => {
    if (!error) error = err instanceof Error ? err : new Error(String(err));
    stopped = true;
  };

  // Attribute values arrive raw from saxen; parseSax decodes entities on
  // demand, so do the same (no-op fast path unless `&` is present).
  const attr = (
    attrs: Record<string, string>,
    key: string,
    decodeEntities: DecodeEntities,
  ): string | undefined => {
    const v = attrs[key];
    if (v === undefined) return undefined;
    return v.indexOf('&') === -1 ? v : decodeEntities(v);
  };

  const appendValueText = (text: string): void => {
    if (!cell) return;
    const v = cell.val;
    if (v && typeof v === 'object' && (v as {richText?: unknown}).richText) {
      // Pathological mixed content — mirrors the classic quirk exactly.
      const rt = (v as {richText: Array<{text?: string}> & {text?: string}}).richText;
      rt.text = rt.text ? rt.text + text : text;
    } else {
      cell.val = v ? (v as string) + text : text;
    }
  };

  const finishRow = (): void => {
    if (!row) return;
    rows.push(finalizeRow(row.number, row.min, row.max, row.cells, row, ctx));
    row = null;
  };

  const parser = new Parser();
  parser.on('error', fail);
  parser.on('warn', fail);

  parser.on(
    'openTag',
    (name: string, getAttrs: () => Record<string, string>, decodeEntities: DecodeEntities) => {
      if (stopped) return;
      if (name === 'row') {
        const attrs = getAttrs();
        const spans = attrs.spans
          ? attr(attrs, 'spans', decodeEntities)!
              .split(':')
              .map(s => parseInt(s, 10))
          : [undefined, undefined];
        const s = attr(attrs, 's', decodeEntities);
        const next: MutableRow = {
          number: parseInt(attr(attrs, 'r', decodeEntities) as string, 10),
          min: spans[0],
          max: spans[1],
          cells: [],
          lastCol: 0,
        };
        if (s !== undefined) next.styleId = parseInt(s, 10);
        if (parseBoolean(attr(attrs, 'hidden', decodeEntities))) next.hidden = true;
        if (parseBoolean(attr(attrs, 'bestFit', decodeEntities))) next.bestFit = true;
        const ht = attr(attrs, 'ht', decodeEntities);
        if (ht !== undefined) next.height = parseFloat(ht);
        const outlineLevel = attr(attrs, 'outlineLevel', decodeEntities);
        if (outlineLevel !== undefined) next.outlineLevel = parseInt(outlineLevel, 10);
        if (parseBoolean(attr(attrs, 'collapsed', decodeEntities))) next.collapsed = true;
        row = next;
        cell = null;
        return;
      }
      if (!row) return;
      if (name === 'c') {
        // `<v>`, `<t>`, `<is>` and run tags carry no attributes we read —
        // skip getAttrs() for them (saxen parses the tag text per call).
        const attrs = getAttrs();
        const s = attr(attrs, 's', decodeEntities);
        cell = {
          raw: attr(attrs, 'r', decodeEntities),
          t: attr(attrs, 't', decodeEntities) ?? '',
          styleId: s !== undefined ? parseInt(s, 10) : undefined,
          val: undefined,
          inIs: false,
          inRun: false,
          inRunT: false,
          runSeen: false,
        };
        return;
      }
      if (!cell) return;
      if (name === 'f' && !cell.inRun) {
        const attrs = getAttrs();
        cell.cur = 'f';
        const si = attr(attrs, 'si', decodeEntities);
        const fType = attr(attrs, 't', decodeEntities);
        const ref = attr(attrs, 'ref', decodeEntities);
        if (si !== undefined) cell.si = si;
        if (fType !== undefined) cell.shareType = fType;
        if (ref !== undefined) cell.ref = ref;
        return;
      }
      if (name === 'v' && !cell.inRun) {
        cell.cur = 'v';
        return;
      }
      if (name === 'is') {
        cell.inIs = true;
        return;
      }
      if (name === 'r') {
        // Rich-text run (inside or outside `<is>` — mirrors classic tolerance).
        // Nested `<r>` is ignored like the classic path (already inside a run).
        if (!cell.inRun) {
          cell.inRun = true;
          cell.inRunT = false;
          cell.runText = undefined;
          cell.runSeen = false;
        }
        return;
      }
      if (name === 't') {
        if (cell.inRun) cell.inRunT = true;
        else cell.cur = 't';
      }
    },
  );

  parser.on('text', (value: string, decodeEntities: DecodeEntities) => {
    if (stopped || !cell) return;
    const text = value.indexOf('&') === -1 ? value : decodeEntities(value);
    if (cell.inRun) {
      // Inside a run only `<t>` text counts (mirrors the rich-text parser);
      // anything after the run's `</t>` is ignored.
      if (cell.inRunT) {
        cell.runText = cell.runText ? cell.runText + text : text;
        cell.runSeen = true;
      }
      return;
    }
    if (cell.cur === 'f') {
      cell.formula = cell.formula ? cell.formula + text : text;
    } else if (cell.cur === 'v' || cell.cur === 't') {
      appendValueText(text);
    }
  });

  parser.on('closeTag', (name: string) => {
    if (stopped) return;
    if (name === 't') {
      // A run consumes its own `</t>` (rich parser stays active, currentNode
      // untouched); otherwise the current node clears.
      if (cell?.inRun) cell.inRunT = false;
      else if (cell) cell.cur = undefined;
      return;
    }
    if (name === 'r') {
      if (cell?.inRun) {
        const runs = cell.runs ?? (cell.runs = []);
        runs.push(cell.runSeen ? {text: decodeControlEscapes(cell.runText as string)} : {});
        // Mirror cell parseClose: value becomes {richText} on first run close.
        if (cell.val === undefined || typeof cell.val === 'string') {
          const prev = typeof cell.val === 'string' ? cell.val : undefined;
          const holder: {richText: Array<{text?: string}>} = {richText: runs};
          if (prev !== undefined) {
            // Pathological mixed content — mirrors the classic quirk.
            (holder.richText as unknown as {text?: string}).text = prev;
          }
          cell.val = holder;
        }
        cell.inRun = false;
      }
      return;
    }
    if (name === 'is') {
      if (cell) {
        cell.inIs = false;
        cell.cur = undefined;
      }
      return;
    }
    if (name === 'v' || name === 'f') {
      if (cell) cell.cur = undefined;
      return;
    }
    if (name === 'c') {
      if (!stopped && row && cell) {
        try {
          finalizeCell(row, cell, shared);
        } catch (err) {
          fail(err);
        }
      }
      cell = null;
      return;
    }
    if (name === 'row') {
      if (!stopped) {
        try {
          finishRow();
        } catch (err) {
          fail(err);
        }
      }
    }
  });

  try {
    parser.write(`<sheetData>${content}</sheetData>`);
    parser.end();
  } catch (err) {
    fail(err);
  }
  if (error) throw error;
  return rows;
}
