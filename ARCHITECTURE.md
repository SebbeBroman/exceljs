# Architecture (5.0 alpha)

Short map of how `@sebbebroman/excel-ts` turns builder calls into `.xlsx` bytes.

## Public surface

| Entry | Role |
|-------|------|
| `@sebbebroman/excel-ts` | Browser-safe: `workbook`, `writeBuffer`, `load`, `csv`, enums |
| `@sebbebroman/excel-ts/node` | Same + `writeFile` / `readFile` / `streamWrite` / `streamRead` / CSV file helpers |

Nothing else is exported (`sideEffects: false`). Legacy modules under `lib/`
(`exceljs.nodejs.ts`, `csv-entry.ts`, Doc classes, stream writers) are **internal**.

## Write path (builder → buffer)

```
workbook().sheet(…).row(…).cell(…)
        │
        ▼
   op-log (BuilderOp[])     lib/builder/*
        │
        ▼
 materialize DocWorkbook     lib/compile/ops-to-doc-workbook.ts
        │
        ▼
   XLSX.encode / zip          lib/xlsx/*  (+ fflate)
        │
        ▼
   Uint8Array (.xlsx)
```

- **Op-log:** fluent builder records ops; does not mutate a cell graph while chaining.
- **Materialize:** ops applied onto the legacy mutable `DocWorkbook` so the existing OOXML encoder can run. Dense rectangular op-logs use a bulk `addRows` path.
- **Encode:** `writeBuffer` → XLSX writer → zip (default deflate **level 1** for speed; override with `{ zip: { level: 6 } }` for smaller files). Optional features (drawings, tables, comments, pivots) load via dynamic `import()` where possible.
- **Cell-by-cell ops:** style-free `.cell()` / `.cells()` logs coalesce into bulk `.rows()` during optimize so they hit the dense materialize path.

Plain snapshots:

```
builder.build()  →  { meta, sheets }   (no I/O)
writeBuffer(plain | builder)
```

## Read path

### Full fidelity (`load` / `readFile`)

```
load(bytes) / readFile(path)
        │
        ▼
   XLSX.decode                 lib/xlsx/load.ts + xforms
        │
        ▼
   DocWorkbook (internal)
        │
        ▼
   plain { meta, sheets }     lib/compile/doc-to-plain.ts
```

Edit loop: `workbook(await load(buf)).sheet(…).cell(…).writeBuffer()`.

### Values-only / lazy view (`viewWorkbook` / `readRows`)

```
viewWorkbook(bytes) / readRows(bytes)
        │
        ▼
   openLightPackage            unzip + workbook.xml + SST only
        │                      (no sheet XML yet)
        ▼
   parseLightSheet(sheet, { start, end, cols })
        │                      saxen, early-exit when end is set
        ▼
   string[][] | CellValue[][]  (no DocWorkbook, styles, drawings)
```

- **Lazy per sheet:** only the requested sheet’s XML is parsed.
- **Early exit:** `rows({ end: 100 })` stops SAX after that row (does not walk the rest of `sheetData`).
- **Tradeoff:** values only — no styles, merges, formulas, hyperlinks, or theme, and dates arrive numeric. Use `load()` for full-fidelity re-encode. Never round-trip `workbook(view)` expecting fidelity.
- Engine: `lib/read/xlsx-light.ts`.
- **Zip limits:** `lib/utils/zip-reader.ts` caps entries (10k) and total uncompressed output (512 MiB); CRC is not re-verified (parse success ≠ integrity proof).

## CSV

Named helpers in `lib/csv/public.ts` (`csv.parse` / `csv.stringify`). Builder `.csv()` dynamically imports the stringify path so write-only bundles can drop `fast-csv`. Node `readCsvFile` / `writeCsvFile` wrap the same helpers with `fs`.

## Streaming (Node only)

- **`streamWrite`:** rows committed as written (bounded memory) via internal stream worksheet writer.
- **`streamRead`:** async row iteration without a full plain model.
- Legacy `WorkbookWriter` / `WorkbookReader` classes are not package exports.

## Why DocWorkbook remains

Phase 7 does **not** delete `lib/doc/*`. The op-log → DocWorkbook → XLSX bridge is an implementation detail; public callers never see the class. A later milestone may encode from ops/plain models directly and drop the mutable graph.

## Bundle size (indicative)

From `pnpm test:browser-bundle` (esbuild minify, **write-only** builder path;
optional CSV module stubbed so sizes match a tree-shaken write client):

| Build | Size (5.0.0-alpha.1) |
|-------|----------------------|
| Single-file minified | **287.2 KB** |
| Single-file gzip | **80.0 KB** |
| Code-split write-only entry | **190.9 KB** |
| Code-split write-only total | 550.9 KB (many small chunks) |

Re-measure after dependency or encoder changes:

```bash
pnpm test:browser-bundle   # also writes build/browser-smoke/sizes.txt
pnpm test:esm
```

## Key modules

| Path | Purpose |
|------|---------|
| `excel.ts` / `node.ts` | Package entries |
| `lib/builder/` | Op-log builder |
| `lib/model/types.ts` | Plain workbook types |
| `lib/compile/` | Ops ↔ DocWorkbook ↔ plain |
| `lib/xlsx/` | Encode/decode OOXML |
| `lib/csv/public.ts` | Public CSV API |
| `lib/stream/xlsx/stream-write.ts` | Node `streamWrite` |
| `lib/stream/xlsx/stream-read.ts` | Node `streamRead` |
| `lib/doc/` | Internal mutable document model |
| `excel.d.ts` | Public TypeScript types |
| `index.d.ts` | Legacy typings (internal reference) |
