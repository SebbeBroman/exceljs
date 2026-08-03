# ESM / tree-shaking (SvelteKit / Vite)

This fork converts the library to native **ES modules** aimed at modern bundlers (Vite, SvelteKit, Rollup, esbuild).

## Install (local fork)

```bash
# in your SvelteKit app
npm install ../path/to/exceljs
# or
npm install github:you/exceljs#branch
```

## Usage

### Preferred (tree-shakeable)

```js
import { Workbook } from 'exceljs';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['name', 'value']);
sheet.addRow(['alpha', 1]);

// SvelteKit +server.js / +server.ts
const buffer = await workbook.xlsx.writeBuffer();
return new Response(buffer, {
  headers: {
    'Content-Type': 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    'Content-Disposition': 'attachment; filename="export.xlsx"',
  },
});
```

### Namespace default (closest to classic API)

```js
import ExcelJS from 'exceljs';

const workbook = new ExcelJS.Workbook();
```

### Optional CSV

CSV is **not** loaded by default (so `fast-csv` / dayjs CSV path can be dropped).

```js
import { Workbook } from 'exceljs';
import 'exceljs/csv'; // side-effect: enables workbook.csv

const workbook = new Workbook();
await workbook.csv.writeFile('out.csv');
```

### Streaming XLSX

```js
import { WorkbookWriter } from 'exceljs/stream/xlsx';
```

## What was changed

| Area | Change |
|------|--------|
| Module system | All of `lib/` converted CJS → ESM (`import` / `export`) |
| Package | `"type": "module"`, `exports` map, `sideEffects` |
| Tree-shaking | Named exports; CSV optional; stream as separate entry |
| Node | Engines ≥ 18; drop ES5/browserify publish path for this fork |

## SvelteKit notes

### Client-side (browser)

The core `Workbook` + `xlsx.load` / `xlsx.writeBuffer` path is **process-free**:
it does not use `readable-stream`, the npm `buffer` polyfill, or bare `process`
(no `vite-plugin-node-polyfills` needed). Binary data is `Uint8Array` (wrapped as
Node `Buffer` when available). Zip is **[fflate](https://github.com/101arrowz/fflate)** (not JSZip).

```js
import { Workbook } from 'exceljs';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['a', 1]);

const buffer = await workbook.xlsx.writeBuffer();
// download via Blob…

const wb2 = new Workbook();
await wb2.xlsx.load(await file.arrayBuffer());
```

**Do not use in the browser:** `readFile` / `writeFile`, `exceljs/stream/xlsx`, path-based CSV,
`worksheet.protect()` (Node crypto).

`package.json` `browser` maps `fs` / `crypto` to shims so Vite can resolve those imports without Node.

Optional Vite bits:

```js
// vite.config.js
export default defineConfig({
  define: { global: 'globalThis' }, // some deps check `global`
  ssr: { noExternal: ['exceljs'] },
});
```

### Server-side

`+server.ts` / `*.server.js` can use the full API including `readFile` / `writeFile`.

## Smoke test

```bash
npm install
npm run test:esm
```
