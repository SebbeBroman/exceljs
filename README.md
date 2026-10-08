# @sebbebroman/exceljs

Read, manipulate, and write Excel workbooks (`.xlsx`) with a **builder-first** ESM API.

Fork of [ExcelJS](https://github.com/exceljs/exceljs) aimed at modern Node and bundlers (Vite, SvelteKit, Rollup, esbuild).

> **0.2.0** — builder-first public API (`.`, `./node`, and optional `./protection`). Write, load, CSV, advanced sheet features, Node streaming. Breaking changes: [CHANGELOG.md](./CHANGELOG.md). See [MIGRATION.md](./MIGRATION.md) and [ARCHITECTURE.md](./ARCHITECTURE.md).

## Install

```bash
npm install @sebbebroman/exceljs
# or
pnpm add @sebbebroman/exceljs
```

**Requirements:** Node.js `>= 22`.

## Quick start

```ts
import { workbook } from '@sebbebroman/exceljs';

const buffer = await workbook({ creator: 'Reports' })
  .sheet('Data')
  .row(['name', 'value'])
  .row(['alpha', 1])
  .style('A1:B1', { font: { bold: true } })
  .writeBuffer();
```

### Node file write

```ts
import { workbook, writeFile } from '@sebbebroman/exceljs/node';

await writeFile(
  'out.xlsx',
  workbook().sheet('Data').rows([
    ['name', 'value'],
    ['alpha', 1],
  ]),
);
```

### Nested sheets

```ts
const buffer = await workbook()
  .sheet('A', s => s.row([1]).row([2]))
  .sheet('B', s => s.cell('A1', 'hello'))
  .writeBuffer();
```

### Title row + table

`.columns([{ header }])` appends a header row after any rows already written, so a title-then-table chain is safe. Prefer `SheetInit.title` (or fluent `.title()`) for the common case:

```ts
const buffer = await workbook().sheet('Report', {
  title: {
    text: 'Q1 Revenue',
    style: { font: { bold: true, size: 16 } },
    merge: 'A1:B1',
  },
  columns: [
    { header: 'Product', key: 'p', width: 20 },
    { header: 'Revenue', key: 'r', width: 12 },
  ],
  rows: [
    { p: 'Widgets', r: 12000 },
    { p: 'Gadgets', r: 8000 },
  ],
}).writeBuffer();

// Equivalent fluent form:
// workbook().sheet('Report').title({ text: 'Q1 Revenue', merge: 'A1:B1' })
//   .columns([...]).rows([...])
```

### Plain snapshot (no I/O)

```ts
const data = workbook().sheet('S').rows([[1, 2]]).build();
// { meta, sheets: [...] }
```

### Load + edit

```ts
import { workbook, load } from '@sebbebroman/exceljs';

const data = await load(buffer); // plain { meta, sheets }
const out = await workbook(data)
  .sheet('Data')
  .cell('A1', 'updated')
  .writeBuffer();
```

### Import CSV / xlsx — view + rows (recommended)

```ts
import { viewWorkbook } from '@sebbebroman/exceljs';
// writes are a separate import so pure readers tree-shake better:
import { workbook, writeBuffer } from '@sebbebroman/exceljs';

const data = reader.result as ArrayBuffer;
const view = await viewWorkbook(data, {
  format: 'auto',
  filename: file.name, // optional sniff (.csv / .xlsx / .xlsm)
});

view.sheetNames; // ['Sheet1', ...]
const sheet = view.sheet(0);

// Dense string grid with optional row/col slice (1-based inclusive)
const rows = sheet.rows({
  start: 1,
  end: 100,
  cols: { start: 1, end: 4 }, // or cols: ['A', 'D']
  // values: 'cell' for CellValue[][]
});

// Header → objects
const records = sheet.records({ header: true });

// Opt-in write from the same view
const buf = await writeBuffer(workbook(view));
```

One-liner sugar (same as first sheet `.rows({ values: 'string' })`):

```ts
import { readRows } from '@sebbebroman/exceljs';
const rows = await readRows(data, { filename: file.name, start: 1, end: 50 });
```

Supports **CSV** and **OOXML** (`.xlsx` / `.xlsm` / …). Not `.xls` / `.xlsb`.

> **Read path choice:** `load()` / `readFile()` parse workbook models (values,
> styles, merges, tables, …) for re-encode, with the round-trip limitations
> listed below. `viewWorkbook()` / `readRows()`
> are **values-only** (no styles, merges, formulas, hyperlinks, dates stay
> numeric) and must not be used for round-trip writes expecting fidelity —
> `workbook(view).writeBuffer()` keeps values only.

### Node file read

```ts
import { readFile, writeFile, workbook } from '@sebbebroman/exceljs/node';

const data = await readFile('in.xlsx');
await writeFile('out.xlsx', workbook(data).sheet('Sheet1').cell('A1', 'x'));
```

### Node streaming write (large row counts)

Prefer `streamWrite` when rows are produced incrementally or the sheet is huge — rows are committed as they are written (bounded memory). Not available from the browser entry.

```ts
import { streamWrite, streamRead } from '@sebbebroman/exceljs/node';

// Declarative: async iterable / array / generator
await streamWrite('big.xlsx', {
  useSharedStrings: true,
  sheets: [
    {
      name: 'Data',
      columns: [
        { header: 'Id', key: 'id', width: 10 },
        { header: 'Name', key: 'name', width: 24 },
      ],
      rows: (async function* () {
        for (let i = 1; i <= 100_000; i++) {
          yield { id: i, name: `row-${i}` };
        }
      })(),
    },
  ],
});

// Callback: full control over multi-sheet flow
await streamWrite('report.xlsx', async w => {
  const sheet = w.sheet('Events', {
    columns: [{ header: 'Ts' }, { header: 'Msg' }],
  });
  await sheet.rows(eventSource());
});

// Optional: stream rows back without a full plain model
for await (const { sheetName, rowNumber, values } of streamRead('big.xlsx')) {
  // values[1] is column A (one-based sparse layout)
  console.log(sheetName, rowNumber, values[1]);
}
```

### CSV

```ts
import { workbook, csv } from '@sebbebroman/exceljs';

// Parse text → sheet → xlsx
const init = await csv.parse('name,value\nalpha,1');
const buffer = await workbook().sheet('Data', init).writeBuffer();

// Stringify active sheet
const text = await workbook()
  .sheet('Data')
  .rows([
    ['name', 'value'],
    ['alpha', 1],
  ])
  .csv();
```

Node:

```ts
import { readCsvFile, writeCsvFile } from '@sebbebroman/exceljs/node';

const data = await readCsvFile('in.csv'); // plain Workbook (one sheet)
await writeCsvFile('out.csv', data);
```

## Entry points

| Import | Purpose |
|--------|---------|
| `@sebbebroman/exceljs` | Builder, `writeBuffer`, `load`, `csv`, enums (browser-safe) |
| `@sebbebroman/exceljs/protection` | Optional synchronous `sheetProtection(password, options)` factory |
| `@sebbebroman/exceljs/node` | + `writeFile` / `readFile` / `streamWrite` / `streamRead` / `readCsvFile` / `writeCsvFile` |

The supported runtime entry points are the main entry, `/node`, and `/protection`; `./package.json` is also exported. There is no `./csv`, `./stream/xlsx`, or default `ExcelJS` class. Types resolve to [`excel.d.ts`](./excel.d.ts) and [`node.d.ts`](./node.d.ts). The mutable document API has been removed; stream coordinators are internal.

Password hashing is opt-in, so unused protection crypto can be tree-shaken out even in single-file browser bundles:

```ts
import {workbook} from '@sebbebroman/exceljs';
import {sheetProtection} from '@sebbebroman/exceljs/protection';

const bytes = await workbook().sheet('Data').row([1, 2])
  .protect(sheetProtection('secret'))
  .writeBuffer();
```

`.protect()` now takes a prepared protection model; the previous password/options overload has been removed. For passwordless protection, use `.protect({sheet: true})`. Hashing runs synchronously when the optional factory is called. Loaded hashes can be re-encoded through the core without importing crypto.

Pipeline overview: [ARCHITECTURE.md](./ARCHITECTURE.md) (op-log → materialize → XLSX).

## Browser / bundlers

Core path uses [fflate](https://github.com/101arrowz/fflate) for zip. No `readable-stream` / npm `buffer` polyfills required for `writeBuffer`.

```ts
import { workbook } from '@sebbebroman/exceljs';

const buffer = await workbook()
  .sheet('Sheet1')
  .row(['a', 1])
  .writeBuffer();
// download via Blob…
```

### Tree-shaking

- Named exports only (`workbook`, `writeBuffer`, `load`, `csv`, …).
- `"sideEffects": false` (importing the package does not extend `dayjs`; CSV helpers extend it on first use).
- Import `@sebbebroman/exceljs/node` only in Node code paths.
- `writeBuffer` and `load` are separate modules (read does not pull write).
- Builder `.csv()` imports the CSV helper module dynamically. Bundlers can retain CSV chunks even in write-only clients because the builder exposes `.csv()`.

Heavy OOXML features still load with the current encoder bridge; later milestones split more of the encoder. Optional drawings/tables/comments/pivots already use dynamic `import()`.

### Dense export optimizations

When the builder op-log is **rectangular only** (`.sheet` / `.columns` / `.row` / `.rows` — no random `.cell` patches, styles, or merges), materialize uses a bulk `addRows` path and keeps compact cell storage (no full Cell class graph per value). Consecutive `.row()` calls are fused into one `.rows` op; same-address `.cell` writes are last-write-wins. If the builder never applied styles, `writeBuffer` defaults `useStyles: false` (override with `{ useStyles: true }`).

### Benchmarks (vs `exceljs@4`)

`exceljs` is a **devDependency** for head-to-head e2e benches (write / read / round-trip) via [mitata](https://github.com/evanwashere/mitata):

```bash
pnpm bench                 # Node: full suite (default 5000×8)
pnpm bench:write
pnpm bench:read
pnpm bench:roundtrip
pnpm bench -- --rows 20000
pnpm bench -- --filter write --rows 1000
pnpm bench:dense           # internal dense-write microbench
pnpm bench:browser         # Chrome headless: size + write/read vs exceljs browser build
pnpm bench:browser -- --rows 1000 --runs 5
```

Contenders (Node): **@sebbebroman/exceljs builder**, **exceljs@4** (npm).

Contenders (browser): **@sebbebroman/exceljs** esbuild browser bundle vs **exceljs** official `dist/exceljs.min.js` (their browser field). Node often favors exceljs; browser compares the polyfill-heavy UMD build against the ESM/fflate path. Needs Chrome (`CHROME_PATH` override supported).

## Status (0.2.0)

| Phase / feature | Status |
|-----------------|--------|
| Phase 1–2: builder + `writeBuffer` / Node `writeFile` | ✅ |
| Phase 3: dense write path / size-minded materialize | ✅ |
| Phase 4: `load` / `readFile` + edit loop | ✅ |
| Phase 5: views, pageSetup, headerFooter, validations, CF, notes, protect, tables, images, defined names | ✅ |
| Phase 6: named `csv`, Node streamWrite / streamRead | ✅ |
| Phase 7: public API cleanup (builder-only exports + docs) | ✅ |
| Builder write (rows, cells, styles, merges, columns) | ✅ |
| CSV (`csv.parse` / `.csv()` / Node file helpers) | ✅ |
| Streaming (`streamWrite` / `streamRead` on `./node`) | ✅ |
| Sheet protection | ✅ write; load re-encodes hash (password not recoverable) |
| Tables / images / defined names | ✅ write; load best-effort |
| Pivot builder API | later (no `.pivot()` yet — read/write of pivot tables is best-effort via load) |
| Drop internal DocWorkbook bridge | ✅ buffered and streaming paths; mutable document API removed |

> **Column headers:** `ColumnInput.header` as an array uses only the first
> line — multi-row headers are not supported and extra lines are dropped
> consistently across `build()` / `writeBuffer()` / `streamWrite()`.

### Bundle size (indicative — write-only fixture)

Measured by `pnpm test:browser-bundle` (esbuild minify, write-only builder path, CSV enabled, 2-cell fixture):

| Build | Size |
|-------|------|
| Single-file minified | ~268 KiB (gzip ~77 KiB) |
| Code-split entry | ~142 KiB (excludes shared/async chunks; total ~266 KiB) |

Not representative of `load`/styles/tables/comments/CSV builds. Quote with fixture + flags + commit hash.

Re-run after encoder changes. Details: [ARCHITECTURE.md](./ARCHITECTURE.md).

## API

Public types: [`excel.d.ts`](./excel.d.ts). Migration from 4.x: [MIGRATION.md](./MIGRATION.md). Architecture: [ARCHITECTURE.md](./ARCHITECTURE.md).

## Attribution

Derived from **ExcelJS**, originally created by [Guyon Roche](https://github.com/guyonroche) and maintained by [exceljs/exceljs](https://github.com/exceljs/exceljs) contributors.

## License

[MIT](./LICENSE)

Copyright (c) 2014–2019 Guyon Roche.

ExcelJS contributions by the ExcelJS team and contributors.

Fork modifications copyright (c) 2026 SebbeBroman and contributors.
