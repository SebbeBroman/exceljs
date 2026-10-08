# Migration from ExcelJS 4.x to @sebbebroman/exceljs 0.2

`@sebbebroman/exceljs` 0.2 is a **breaking** redesign of the ExcelJS 4.x API. The public API is a fluent **builder** plus three package entry points. The mutable ExcelJS-style `Workbook` / `Worksheet` classes are no longer exported.

For the 0.1.x protection API changes, see [the upgrade section](#upgrading-from-01x-to-020-opt-in-protection-hashing) and [CHANGELOG.md](./CHANGELOG.md).

## Entry points

| Import                            | Role                                                                              |
| --------------------------------- | --------------------------------------------------------------------------------- |
| `@sebbebroman/exceljs`            | Builder, `writeBuffer`, `load`, `csv`, types, enums (browser-safe)                |
| `@sebbebroman/exceljs/protection` | Optional synchronous `sheetProtection(password, options)` factory                 |
| `@sebbebroman/exceljs/node`       | Same + `writeFile` / `readFile` / `streamWrite` / `streamRead` / CSV file helpers |

Differences from ExcelJS 4.x:

- `@sebbebroman/exceljs/csv` (side-effect import) — use named `csv` from the main entry instead
- `@sebbebroman/exceljs/stream/xlsx` — use `streamWrite` / `streamRead` on `./node` (no public `WorkbookWriter` / `WorkbookReader`)
- Default export `import ExcelJS from '…'`

## API map

| 4.x (mutable)                                                                               | 0.2.x (builder)                                                                        |
| ------------------------------------------------------------------------------------------- | -------------------------------------------------------------------------------------- |
| `new Workbook()`                                                                            | `workbook()`                                                                           |
| `wb.creator = '…'`                                                                          | `workbook({ creator: '…' })` or `.props({ creator: '…' })`                             |
| `wb.addWorksheet('S')`                                                                      | `.sheet('S')`                                                                          |
| `ws.getCell('A1').value = x`                                                                | `.cell('A1', x)`                                                                       |
| `ws.addRow([…])` / `addRows`                                                                | `.row([…])` / `.rows([…])`                                                             |
| `cell.font = { bold: true }`                                                                | `.cell('A1', v, { font: { bold: true } })` or `.style('A1', { font: { bold: true } })` |
| `ws.mergeCells('A1:B2')`                                                                    | `.merge('A1:B2')`                                                                      |
| `ws.columns = […]`                                                                          | `.columns([…])`                                                                        |
| `ws.views = […]`                                                                            | `.views([…])`                                                                          |
| `ws.pageSetup = {…}`                                                                        | `.pageSetup({…})`                                                                      |
| `ws.headerFooter = {…}`                                                                     | `.headerFooter({…})`                                                                   |
| `cell.dataValidation = {…}`                                                                 | `.dataValidation(address, rules)`                                                      |
| `ws.addConditionalFormatting(cf)`                                                           | `.conditionalFormatting(cf)`                                                           |
| `cell.note = '…'`                                                                           | `.note(address, text \| note)`                                                         |
| `await ws.protect(pw, opts)`                                                                | `.protect(sheetProtection(pw, opts))` (opt-in `/protection` import)                    |
| `ws.addTable(props)`                                                                        | `.table(props)`                                                                        |
| `wb.addImage` + `ws.addImage`                                                               | `const id = b.image(def); b.image(id, range)`                                          |
| `wb.definedNames.add(range, name)`                                                          | `.definedName(name, refersTo)`                                                         |
| `await wb.xlsx.writeBuffer()`                                                               | `await b.writeBuffer()` or `await writeBuffer(b)`                                      |
| `await wb.xlsx.writeFile(p)`                                                                | `await writeFile(p, b)` from `@sebbebroman/exceljs/node`                               |
| `await wb.xlsx.load(buf)`                                                                   | `await load(buf)` → plain `{ meta, sheets }`                                           |
| `await wb.xlsx.readFile(p)`                                                                 | `await readFile(p)` from `@sebbebroman/exceljs/node`                                   |
| edit after load                                                                             | `workbook(await load(buf)).sheet(…).cell(…).writeBuffer()`                             |
| `import '…/csv'` + `wb.csv.read/write`                                                      | `csv.parse` / `csv.stringify` / `.csv()`; Node `readCsvFile` / `writeCsvFile`          |
| `new ExcelJS.stream.xlsx.WorkbookWriter({filename})` + `addWorksheet` / `addRow` / `commit` | `streamWrite(path, spec)` on `@sebbebroman/exceljs/node`                               |
| `new ExcelJS.stream.xlsx.WorkbookReader(path)` async iterate                                | `streamRead(path)` on `./node` (row-oriented); or `readFile` for full plain model      |

## Examples

### Write buffer (browser or Node)

```ts
import {workbook} from '@sebbebroman/exceljs';

const buffer = await workbook({creator: 'Reports'})
  .sheet('Q1')
  .rows([
    ['Product', 'Revenue'],
    ['Widgets', 12000],
  ])
  .style('A1:B1', {font: {bold: true}})
  .writeBuffer();
```

### Write file (Node)

```ts
import {workbook, writeFile} from '@sebbebroman/exceljs/node';

await writeFile('out.xlsx', workbook().sheet('Data').row(['a', 1]).row(['b', 2]));
```

### Nested sheet callback

```ts
workbook()
  .sheet('A', s => s.row([1]).row([2]))
  .sheet('B', s => s.cell('A1', 'hello'))
  .writeBuffer();
```

### Plain snapshot (no I/O)

```ts
const data = workbook()
  .sheet('S')
  .rows([[1, 2]])
  .build();
// data: { meta, sheets: [...] }
```

### Load + edit loop

```ts
import {workbook, load} from '@sebbebroman/exceljs';

const bytes = await workbook().sheet('Sheet1').cell('A1', 'hi').writeBuffer();
const data = await load(bytes); // plain { meta, sheets } — not a mutable class

const out = await workbook(data).sheet('Sheet1').cell('A1', 'updated').writeBuffer();
```

### Read file (Node)

```ts
import {readFile, writeFile, workbook} from '@sebbebroman/exceljs/node';

const data = await readFile('in.xlsx');
await writeFile('out.xlsx', workbook(data).sheet('Sheet1').cell('A1', 'x'));
```

### CSV (named API — no side-effect import)

```ts
import {workbook, csv} from '@sebbebroman/exceljs';

// Parse → sheet init → xlsx
const init = await csv.parse('name,value\nalpha,1');
const xlsx = await workbook().sheet('Data', init).writeBuffer();

// Stringify active sheet
const text = await workbook()
  .sheet('Data')
  .rows([
    ['name', 'value'],
    ['alpha', 1],
  ])
  .csv();

// Or free helpers
const again = await csv.stringify(workbook().sheet('Data', init));
```

Node file helpers:

```ts
import {readCsvFile, writeCsvFile, workbook} from '@sebbebroman/exceljs/node';

const data = await readCsvFile('in.csv'); // plain Workbook, one sheet
await writeCsvFile('out.csv', data);
await writeCsvFile('sheet-b.csv', data, {sheetName: 'B'});
```

**ExcelJS 4.x → fork 0.2.x:** replace `wb.csv.read/write` with `csv.parse` / `csv.stringify` or builder `.csv()`. The fork has no `./csv` entry or mutable Workbook class.

### Advanced sheet features (Phase 5)

```ts
import {workbook} from '@sebbebroman/exceljs';
import {sheetProtection} from '@sebbebroman/exceljs/protection';

const buf = await workbook()
  .sheet('Report')
  .rows([
    ['Name', 'Score'],
    ['Ada', 95],
    ['Bob', 80],
  ])
  // freeze header row
  .views([{state: 'frozen', ySplit: 1, topLeftCell: 'A2'}])
  .pageSetup({orientation: 'landscape', fitToPage: true, fitToWidth: 1})
  .headerFooter({oddHeader: 'Q1 Report'})
  .dataValidation('B2:B100', {
    type: 'whole',
    operator: 'between',
    formulae: [0, 100],
    allowBlank: true,
  })
  .conditionalFormatting({
    ref: 'B2:B100',
    rules: [
      {
        type: 'cellIs',
        operator: 'greaterThan',
        formulae: [90],
        priority: 1,
        style: {font: {bold: true}},
      },
    ],
  })
  .note('A2', 'Top performer')
  .protect(sheetProtection('secret', {spinCount: 1000})) // import from /protection; hashes now
  .table({
    name: 'Scores',
    ref: 'A1:B3',
    columns: [{name: 'Name'}, {name: 'Score'}],
    rows: [
      ['Ada', 95],
      ['Bob', 80],
    ],
  })
  .definedName('ScoresRange', 'Report!$B$2:$B$3')
  .writeBuffer();

// Images: register media (returns id), then place on the active sheet
const b = workbook().sheet('Img');
const imgId = b.image({extension: 'png', buffer: pngBytes});
await b.image(imgId, 'B2:D6').writeBuffer();
```

### Streaming (Node) — WorkbookWriter → streamWrite

```ts
// 4.x
const wb = new ExcelJS.stream.xlsx.WorkbookWriter({filename: 'out.xlsx', useSharedStrings: true});
const ws = wb.addWorksheet('Data');
ws.columns = [
  {header: 'Id', key: 'id'},
  {header: 'Name', key: 'name'},
];
for (const row of source) {
  ws.addRow(row).commit();
}
await wb.commit();

// Fork 0.2.x — declarative
import {streamWrite, streamRead} from '@sebbebroman/exceljs/node';

await streamWrite('out.xlsx', {
  useSharedStrings: true,
  sheets: [
    {
      name: 'Data',
      columns: [
        {header: 'Id', key: 'id'},
        {header: 'Name', key: 'name'},
      ],
      rows: source, // Iterable | AsyncIterable | array
    },
  ],
});

// Fork 0.2.x — callback (control flow)
await streamWrite('out.xlsx', async w => {
  const sheet = w.sheet('Data', {
    columns: [
      {header: 'Id', key: 'id'},
      {header: 'Name', key: 'name'},
    ],
  });
  await sheet.rows(source);
});

// Optional row stream read (no full plain Workbook)
for await (const {sheetName, rowNumber, values} of streamRead('out.xlsx')) {
  // values[1] = column A
}
```

`WorkbookWriter` / `WorkbookReader` are **not** public package exports. There is no `./stream/xlsx` export.

## 0.2.0 status

Phases 1–7 of the builder rewrite are complete for the **public** product:

| Phase | Scope                                                | Status |
| ----- | ---------------------------------------------------- | ------ |
| 1–2   | Builder + writeBuffer / writeFile                    | ✅     |
| 3     | Dense write / materialize optimizations              | ✅     |
| 4     | load / readFile + edit loop                          | ✅     |
| 5     | Advanced sheet features                              | ✅     |
| 6     | CSV + Node streamWrite / streamRead                  | ✅     |
| 7     | Cleanup: exports, docs, internal-only legacy entries | ✅     |

**Supported:** builder write path (sheets, rows, cells, styles, merges, columns), advanced sheet features (views, pageSetup, headerFooter, dataValidation, CF, notes, protect, tables, images, defined names), `writeBuffer`, `load`, Node `writeFile` / `readFile`, Node `streamWrite` / `streamRead`, edit loop via `workbook(plain)`, named `csv` parse/stringify, builder `.csv()`, Node `readCsvFile` / `writeCsvFile`.

**Load covers (core):** null/number/string/boolean/Date/formula/hyperlink/richText cell values, basic cell styles, merges, column widths, workbook meta.

**Load covers (Phase 5):** sheet views, pageSetup/headerFooter, data validations, conditional formatting, notes, sheet protection (hashed model for re-encode — password not recoverable), tables/images/defined names best-effort.

**Write-only / limited round-trip:** protect password (hash only on load), complex image anchors, pivot tables.

**Not yet:** pivot builder API, full ExcelJS feature surface.

**Removed internally:** mutable Workbook/Worksheet/Row/Cell/Column classes, `lib/doc`, namespace/CSV attachment entries, old compiler bridges, and legacy `index.d.ts`. Buffered and Node streaming paths use encoder models and sparse value arrays directly. Reusable feature helpers live in `lib/model`; internal stream coordinators are not exported. See [ARCHITECTURE.md](./ARCHITECTURE.md).

## Tree-shaking

- Prefer named imports: `import { workbook, writeBuffer } from '@sebbebroman/exceljs'`.
- Builder `.csv()` imports its helper module dynamically; bundlers can retain CSV chunks even in write-only clients.
- Do not import `@sebbebroman/exceljs/node` from browser bundles.
- `"sideEffects": false` on the package.
- There is no `./csv` package export — CSV is folded into the main entry.

## Compact ranges and native streams

Loaded `dataValidations` retain range keys such as `A1:B20`; they no longer allocate
one entry for every covered cell. Single-cell keys remain single-cell keys. When
inspecting a loaded rule, check its range coverage instead of assuming an `A1`
property exists for a rule stored under `A1:B20`. Defined names likewise normalize
ranges without building a cell matrix. Whole-grid validations now load and round-trip.

Node streaming uses native writable streams with 64 KiB XML batches. Declarative
row sources and `await sheet.rows(source)` honor destination backpressure. The
synchronous `sheet.row(values)` convenience cannot wait for drain; large loops
using it can still queue rows. Shared strings and style caches have separate memory
costs. Declarative sheets finish sequentially; callback sheets written out of order
can also require ZIP buffering until earlier sheets finish.

Internal XLSX file/stream wrappers, the old ZIP facade, `StreamBuf`, and the custom
event emitter have been removed. The public Node file and streaming functions
remain available. Tests use direct imports and native Vitest assertions, and
`pnpm typecheck:spec` checks the complete maintained test suite.

## Upgrading from 0.1.x to 0.2.0: opt-in protection hashing

Import `sheetProtection` from `@sebbebroman/exceljs/protection` to prepare a password hash, then pass that model to `.protect()`. The factory is synchronous. The previous `.protect(password, options)` and plain-model `sheet.protect` APIs have been removed; plain models store prepared data in `sheet.sheetProtection`. Passwordless protection accepts `.protect({sheet: true, selectLockedCells: false})` directly. Core writes and edits of loaded protection models require no crypto import.
