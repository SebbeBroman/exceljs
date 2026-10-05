/**
 * Node-only types for `@sebbebroman/exceljs/node`.
 *
 * Re-exports the full browser-safe API from `./excel.js` plus Node-only
 * file/stream helpers. Do not import this entry from browser bundles
 * (it pulls `node:fs`).
 */
export * from './excel.js';
