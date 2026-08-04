# @sebbebroman/excel-ts

Read, manipulate, and write Excel workbooks (`.xlsx`) and CSV files.

ESM + TypeScript fork of [ExcelJS](https://github.com/exceljs/exceljs), aimed at modern Node and bundlers (Vite, SvelteKit, Rollup, esbuild) with tree-shakeable entry points.

## Install

```bash
npm install @sebbebroman/excel-ts
# or
pnpm add @sebbebroman/excel-ts
```

**Requirements:** Node.js `>= 22`.

## Quick start

```ts
import { Workbook } from '@sebbebroman/excel-ts';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['name', 'value']);
sheet.addRow(['alpha', 1]);

await workbook.xlsx.writeFile('out.xlsx');
// or: const buffer = await workbook.xlsx.writeBuffer();
```

Namespace-style default (closest to classic ExcelJS):

```ts
import ExcelJS from '@sebbebroman/excel-ts';

const workbook = new ExcelJS.Workbook();
```

## Entry points

| Import | Purpose |
|--------|---------|
| `@sebbebroman/excel-ts` | Core workbook API (xlsx read/write) |
| `@sebbebroman/excel-ts/csv` | Optional CSV (`workbook.csv`) |
| `@sebbebroman/excel-ts/stream/xlsx` | Streaming reader/writer |

### CSV (optional)

CSV is not loaded by default so apps that never need it can drop `fast-csv` from the dependency graph when tree-shaking allows.

```ts
import { Workbook } from '@sebbebroman/excel-ts';
import '@sebbebroman/excel-ts/csv'; // enables workbook.csv

const workbook = new Workbook();
await workbook.csv.writeFile('out.csv');
```

### Streaming XLSX

```ts
import { WorkbookWriter, WorkbookReader } from '@sebbebroman/excel-ts/stream/xlsx';
```

## Browser / bundlers

The core `Workbook` + `xlsx.load` / `xlsx.writeBuffer` path does not need Node polyfills (`readable-stream`, `buffer`, bare `process`). Zip uses [fflate](https://github.com/101arrowz/fflate). Sheet password protection uses [`@noble/hashes`](https://github.com/paulmillr/noble-hashes) and Web Crypto for salt.

**Browser-friendly:** `writeBuffer` / `load`, cell styles, protection, most xlsx features.

**Node-only:** `readFile` / `writeFile`, `@sebbebroman/excel-ts/stream/xlsx`, path-based CSV.

```ts
import { Workbook } from '@sebbebroman/excel-ts';

const workbook = new Workbook();
workbook.addWorksheet('Sheet1').addRow(['a', 1]);
const buffer = await workbook.xlsx.writeBuffer();
// download via Blob…
```

Optional Vite notes:

```js
// vite.config.js
export default defineConfig({
  define: { global: 'globalThis' }, // some deps check `global`
  ssr: { noExternal: ['@sebbebroman/excel-ts'] },
});
```

### Tree-shaking

Heavy pieces (drawings, tables, comments, pivot, conditional formatting, SAX) are loaded via dynamic `import()` when used. With Rollup/Vite code-splitting, write-only plain-cell apps keep a smaller initial chunk.

In the browser, tables / images / pivots need either:

- `await ensureDocFeatures()` once after import, or
- an async `xlsx.load` / `writeBuffer` path that ensures them automatically.

## API

Public types live in [`index.d.ts`](./index.d.ts). The workbook/worksheet surface is largely compatible with ExcelJS.

For the full historical API walkthrough (styles, tables, images, validations, streaming options, and so on), see the [upstream ExcelJS README](https://github.com/exceljs/exceljs#interface). Where this package differs:

- ESM only (`"type": "module"`)
- Named exports preferred for tree-shaking
- CSV and streaming xlsx are separate entry points
- Node `>= 22`; no legacy browserify / ES5 publish path

## Attribution

This package is derived from **ExcelJS**, originally created by [Guyon Roche](https://github.com/guyonroche) and maintained by [exceljs/exceljs](https://github.com/exceljs/exceljs) contributors.

Substantial changes in this fork include: native ESM packaging, strict TypeScript sources, optional CSV/stream entry points, fflate-based zip, pure-JS sheet protection, and tree-shake oriented lazy loading.

## License

[MIT](./LICENSE)

Copyright (c) 2014–2019 Guyon Roche  
Copyright (c) 2025–2026 SebbeBroman and contributors
