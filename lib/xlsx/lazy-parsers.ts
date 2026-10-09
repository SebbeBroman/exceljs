/**
 * Lazy loaders for optional OOXML parser trees.
 *
 * These features are not needed for plain cell reads. Using dynamic
 * import() lets bundlers with code-splitting (Vite/Rollup/esbuild splitting)
 * keep them out of the initial load chunk until drawings, tables,
 * or comments are encountered.
 *
 * Module-level caches reuse the constructors after their first import.
 */

// Optional parser families have heterogeneous constructor signatures.
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type XformConstructor = new (...args: any[]) => any;

let drawingXform: XformConstructor | undefined;
let tableXform: XformConstructor | undefined;
let commentsXform: XformConstructor | undefined;
let vmlNotesXform: XformConstructor | undefined;
export async function loadDrawingXform(): Promise<XformConstructor> {
  if (!drawingXform) {
    drawingXform = (await import('./parser/drawing/drawing-xform.js')).default;
  }
  return drawingXform;
}

export async function loadTableXform(): Promise<XformConstructor> {
  if (!tableXform) {
    tableXform = (await import('./parser/table/table-xform.js')).default;
  }
  return tableXform;
}

export async function loadCommentsXform(): Promise<XformConstructor> {
  if (!commentsXform) {
    commentsXform = (await import('./parser/comment/comments-xform.js')).default;
  }
  return commentsXform;
}

export async function loadVmlNotesXform(): Promise<XformConstructor> {
  if (!vmlNotesXform) {
    vmlNotesXform = (await import('./parser/comment/vml-notes-xform.js')).default;
  }
  return vmlNotesXform;
}

let cfXforms:
  | {
      ConditionalFormattingsXform: XformConstructor;
      ExtLstXform: XformConstructor;
    }
  | undefined;

/**
 * Conditional formatting + extLst (x14) parser tree.
 * Readers install these automatically when parsing worksheets.
 */
export async function loadCfXforms(): Promise<{
  ConditionalFormattingsXform: XformConstructor;
  ExtLstXform: XformConstructor;
}> {
  if (!cfXforms) {
    const [cf, ext] = await Promise.all([
      import('./parser/sheet/cf/conditional-formattings-xform.js'),
      import('./parser/sheet/ext-lst-xform.js'),
    ]);
    cfXforms = {
      ConditionalFormattingsXform: cf.default,
      ExtLstXform: ext.default,
    };
  }
  return cfXforms;
}
