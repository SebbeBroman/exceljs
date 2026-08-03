/**
 * Lazy loaders for optional worksheet document features:
 * tables, images (anchors), pivot tables.
 *
 * Dynamic import() keeps these out of the initial Workbook chunk for apps that
 * only use plain cells. Method signatures stay sync (addTable, addImage, …).
 *
 * - Node: first use sync-loads via createRequire (tests + scripts).
 * - Browser: call `await ensureDocFeatures()` once before tables/images/pivots,
 *   or rely on xlsx load/write which ensure automatically when needed.
 * - Vitest setup awaits ensureDocFeatures() so the suite stays sync-compatible.
 */

import {createRequire} from 'node:module';

const cache = Object.create(null);
let loading;

const nodeRequire =
  typeof process !== 'undefined' && process.versions && process.versions.node
    ? createRequire(import.meta.url)
    : null;

export async function ensureDocFeatures() {
  if (cache.Table) {
    return cache;
  }
  if (!loading) {
    loading = Promise.all([
      import('./table.js'),
      import('./image.js'),
      import('./pivot-table.js'),
    ]).then(([table, image, pivot]) => {
      cache.Table = table.default;
      cache.Image = image.default;
      cache.makePivotTable = pivot.makePivotTable;
      return cache;
    });
  }
  return loading;
}

function syncLoadNode() {
  if (cache.Table || !nodeRequire) {
    return Boolean(cache.Table);
  }
  try {
    // Build paths at runtime so bundlers (esbuild/Vite) do not statically
    // include these modules in the main chunk via createRequire analysis.
    const t = `./${'table'}.js`;
    const i = `./${'image'}.js`;
    const p = `./${'pivot-table'}.js`;
    cache.Table = nodeRequire(t).default;
    cache.Image = nodeRequire(i).default;
    cache.makePivotTable = nodeRequire(p).makePivotTable;
    return true;
  } catch {
    return false;
  }
}

function notReady(name) {
  return new Error(
    `ExcelJS: ${name} support is not loaded yet. ` +
      `Await ensureDocFeatures() once after importing exceljs before using tables, images, or pivots ` +
      `(xlsx load/write will ensure this automatically when needed).`,
  );
}

export function getTable() {
  if (!cache.Table) syncLoadNode();
  if (!cache.Table) throw notReady('Table');
  return cache.Table;
}

export function getImage() {
  if (!cache.Image) syncLoadNode();
  if (!cache.Image) throw notReady('Image');
  return cache.Image;
}

export function getMakePivotTable() {
  if (!cache.makePivotTable) syncLoadNode();
  if (!cache.makePivotTable) throw notReady('PivotTable');
  return cache.makePivotTable;
}

export function docFeaturesLoaded() {
  return Boolean(cache.Table);
}
