# Migration from ExcelJS 4.x to @sebbebroman/exceljs 0.1

`@sebbebroman/exceljs` 0.1 is a **breaking** redesign of the ExcelJS 4.x API. The public API is a fluent **builder** plus two package entry points. The mutable ExcelJS-style `Workbook` / `Worksheet` classes are no longer exported.

## Entry points

| Import                      | Role                                                                              |
| --------------------------- | --------------------------------------------------------------------------------- |
| `@sebbebroman/exceljs`      | Builder, `writeBuffer`, `load`, `csv`, types, enums (browser-safe)                |
| `@sebbebroman/exceljs/node` | Same + `writeFile` / `readFile` / `streamWrite` / `streamRead` / CSV file helpers |

Removed:

- `@sebbebroman/exceljs/csv` (side-effect import) — use named `csv` from the main entry instead
- `@sebbebroman/exceljs/stream/xlsx` — use `streamWrite` / `streamRead` on `./node` (no public `WorkbookWriter` / `WorkbookReader`)
- Default export `import ExcelJS from '…'`

## API map

| 4.x (mutable)                                                                               | 0.1.x (builder)                                                                        |
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
| `await ws.protect(pw, opts)`                                                                | `.protect(pw, opts)` (deferred hash; chain stays sync)                                 |
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

**ExcelJS 4.x → fork 0.1.x:** drop `import '@sebbebroman/exceljs/csv'`. There is no `workbook.csv` property on a class. Use `csv.parse` / `csv.stringify` or builder `.csv()`.

### Advanced sheet features (Phase 5)

```ts
import {workbook} from '@sebbebroman/exceljs';

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
  .protect('secret', {spinCount: 1000}) // hashed at write time; chain stays sync
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

// Fork 0.1.x — declarative
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

// Fork 0.1.x — callback (control flow)
await streamWrite('out.xlsx', async w => {
  const sheet = w.sheet('Data', {
    columns: [
      {header: 'Id', key: 'id'},
      {header: 'Name', key: 'name'},
    ],
  });
  for await (const row of source) sheet.row(row);
});

// Optional row stream read (no full plain Workbook)
for await (const {sheetName, rowNumber, values} of streamRead('out.xlsx')) {
  // values[1] = column A
}
```

`WorkbookWriter` / `WorkbookReader` are **not** public package exports. There is no `./stream/xlsx` export.

## 0.1.0 status

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

**Not yet:** pivot builder API, full ExcelJS feature surface, direct op-log encode without DocWorkbook.

**Internal (not public):** mutable Doc Workbook / Worksheet classes, `lib/exceljs.nodejs.ts`, `lib/csv-entry.ts`, stream `WorkbookWriter` / `WorkbookReader`. These back tests and the encode/decode bridge. See [ARCHITECTURE.md](./ARCHITECTURE.md).

## Tree-shaking

- Prefer named imports: `import { workbook, writeBuffer } from '@sebbebroman/exceljs'`.
- Import `csv` only when needed; write-only paths do not pull `fast-csv` unless you call `.csv()` or import `csv`.
- Do not import `@sebbebroman/exceljs/node` from browser bundles.
- `"sideEffects": false` on the package.
- There is no `./csv` package export — CSV is folded into the main entry.
