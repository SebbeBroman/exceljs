/**
 * Lazy loaders for optional OOXML xform trees.
 *
 * These features are not needed for plain cell read/write. Using dynamic
 * import() lets bundlers with code-splitting (Vite/Rollup/esbuild splitting)
 * keep them out of the initial Workbook chunk until drawings, tables,
 * comments, or pivot tables are actually used.
 *
 * Module-level caches share one in-flight import across concurrent callers.
 */

let drawingXform;
let tableXform;
let commentsXform;
let vmlNotesXform;
let pivotXforms;

export async function loadDrawingXform() {
  if (!drawingXform) {
    drawingXform = (await import('./xform/drawing/drawing-xform.js')).default;
  }
  return drawingXform;
}

export async function loadTableXform() {
  if (!tableXform) {
    tableXform = (await import('./xform/table/table-xform.js')).default;
  }
  return tableXform;
}

export async function loadCommentsXform() {
  if (!commentsXform) {
    commentsXform = (await import('./xform/comment/comments-xform.js')).default;
  }
  return commentsXform;
}

export async function loadVmlNotesXform() {
  if (!vmlNotesXform) {
    vmlNotesXform = (await import('./xform/comment/vml-notes-xform.js')).default;
  }
  return vmlNotesXform;
}

export async function loadPivotXforms() {
  if (!pivotXforms) {
    const [records, definition, table] = await Promise.all([
      import('./xform/pivot-table/pivot-cache-records-xform.js'),
      import('./xform/pivot-table/pivot-cache-definition-xform.js'),
      import('./xform/pivot-table/pivot-table-xform.js'),
    ]);
    pivotXforms = {
      PivotCacheRecordsXform: records.default,
      PivotCacheDefinitionXform: definition.default,
      PivotTableXform: table.default,
    };
  }
  return pivotXforms;
}

let cfXforms;

/**
 * Conditional formatting + extLst (x14) tree (~14KB min).
 * Write paths without CF never load this.
 */
export async function loadCfXforms() {
  if (!cfXforms) {
    const [cf, ext] = await Promise.all([
      import('./xform/sheet/cf/conditional-formattings-xform.js'),
      import('./xform/sheet/ext-lst-xform.js'),
    ]);
    cfXforms = {
      ConditionalFormattingsXform: cf.default,
      ExtLstXform: ext.default,
    };
  }
  return cfXforms;
}
