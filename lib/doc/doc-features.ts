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
import type Table from './table.js';
import type Image from './image.js';
import type {makePivotTable as MakePivotTableFn} from './pivot-table.js';

export interface DocFeaturesCache {
  Table?: typeof Table;
  Image?: typeof Image;
  makePivotTable?: typeof MakePivotTableFn;
}

const cache: DocFeaturesCache = Object.create(null) as DocFeaturesCache;
let loading: Promise<DocFeaturesCache> | undefined;

const nodeRequire =
  typeof process !== 'undefined' && process.versions && process.versions.node
    ? createRequire(import.meta.url)
    : null;

export async function ensureDocFeatures(): Promise<DocFeaturesCache> {
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

function syncLoadNode(): boolean {
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

function notReady(name: string): Error {
  return new Error(
    `@sebbebroman/exceljs: ${name} support is not loaded yet. ` +
      `Await ensureDocFeatures() once after importing the package before using tables, images, or pivots ` +
      `(xlsx load/write will ensure this automatically when needed).`,
  );
}

export function getTable(): typeof Table {
  if (!cache.Table) syncLoadNode();
  if (!cache.Table) throw notReady('Table');
  return cache.Table;
}

export function getImage(): typeof Image {
  if (!cache.Image) syncLoadNode();
  if (!cache.Image) throw notReady('Image');
  return cache.Image;
}

export function getMakePivotTable(): typeof MakePivotTableFn {
  if (!cache.makePivotTable) syncLoadNode();
  if (!cache.makePivotTable) throw notReady('PivotTable');
  return cache.makePivotTable;
}

export function docFeaturesLoaded(): boolean {
  return Boolean(cache.Table);
}
