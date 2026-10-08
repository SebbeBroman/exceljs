# Architecture (0.2)

Short map of how `@sebbebroman/exceljs` turns builder calls into `.xlsx` bytes.

## Public surface

| Entry                             | Role                                                                              |
| --------------------------------- | --------------------------------------------------------------------------------- |
| `@sebbebroman/exceljs`            | Browser-safe: `workbook`, `writeBuffer`, `load`, `csv`, enums                     |
| `@sebbebroman/exceljs/protection` | Optional synchronous password hashing; core accepts prepared protection models    |
| `@sebbebroman/exceljs/node`       | Same + `writeFile` / `readFile` / `streamWrite` / `streamRead` / CSV file helpers |

`./package.json` is also exported; the package declares `sideEffects: false`. Internal stream coordinators and OOXML transforms are not package exports. The mutable document API and namespace entries have been removed.

## Write path (builder → buffer)

```
workbook().sheet(…).row(…).cell(…)
        │
        ▼
   op-log (BuilderOp[])     lib/builder/*
        │
        ▼
 compile encoder model      lib/compile/ops-to-xlsx-model.ts
        │
        ▼
   XLSX.encode / zip          lib/xlsx/*  (+ fflate)
        │
        ▼
   Uint8Array (.xlsx)
```

- **Op-log:** fluent builder records ops; does not mutate a cell graph while chaining.
- **Compile:** optimized ops populate the OOXML encoder model directly, using sparse row maps and cell slots. No Workbook/Worksheet/Row/Cell document instances are constructed. Plain snapshots import row/cell models directly without a per-cell op-log. Tables place values directly into model cells; image anchors use dimension-only helpers.
- **Encode:** `writeBuffer` → XLSX writer → zip (default deflate **level 1** for speed; override with `{ zip: { level: 6 } }` for smaller files). Optional features (drawings, tables, comments, pivots) load via dynamic `import()` where possible.
- **Cell-by-cell ops:** style-free `.cell()` / `.cells()` logs coalesce into bulk `.rows()` during optimize so they use bulk row compilation.

Plain snapshots:

```
builder.build()  →  { meta, sheets }   (no I/O)
writeBuffer(plain | builder)
```

## Read path

### Workbook model (`load` / `readFile`)

```
load(bytes) / readFile(path)
        │
        ▼
   XLSX.decode                 lib/xlsx/load.ts + xforms
        │
        ▼
   reconciled encoder model
        │
        ▼
   plain { meta, sheets }     lib/compile/xlsx-model-to-plain.ts
```

Edit loop: `workbook(await load(buf)).sheet(…).cell(…).writeBuffer()`.

Model-loading `load` runs in two phases: package parts first (workbook,
SST, styles, rels, …), then sheets — each sheet flows parse → reconcile-ready
without waiting on other sheets. Sheets eligible for the fused fast path
(`lib/xlsx/xform/sheet/fast-sheet-data.ts`) parse + reconcile `sheetData`
in one saxen pass (style/date/shared-string/formula/hyperlink/comment
resolution inline, per-sheet style caches); sheet head/tail still parses
with WorksheetXform. Public `load` projects reconciled cell models directly
into the plain snapshot, avoiding document hydration. Workbook sheet order is restored
explicitly after concurrent parsing. Anything the fused parser
does not implement (run fonts, phonetics, extensions, `ignoreNodes`, …)
falls back to the classic path — correctness first. Disable for diagnostics
via `setFastSheetDataEnabled(false)` (test-only hook).

`writeBuffer` collects fully-materialized parts and runs one synchronous
`zipSync` (`lib/utils/buffer-zip.ts`) instead of the streaming ZipWriter's
event-loop hops (same parts, same level). Public Node `writeFile()` writes the bytes returned by `writeBuffer()`.
Node `streamWrite()` uses the streaming writer.

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
- **Tradeoff:** values only — no styles, merges, formulas, hyperlinks, or theme, and dates arrive numeric. Use `load()` to preserve supported workbook features; complex image anchors and pivot tables remain best-effort. Never round-trip `workbook(view)` expecting fidelity.
- Engine: `lib/read/xlsx-light.ts`.
- **Zip limits:** `lib/utils/zip-reader.ts` caps entries (10k) and total uncompressed output (512 MiB); CRC is not re-verified (parse success ≠ integrity proof).

## CSV

Named helpers in `lib/csv/public.ts` (`csv.parse` / `csv.stringify`). Builder `.csv()` dynamically imports the helper module. Bundlers can retain CSV chunks even in write-only clients because the builder exposes `.csv()`. Node `readCsvFile` / `writeCsvFile` wrap the same helpers with `fs`.

## Streaming (Node only)

- **`streamWrite`:** incoming arrays/keyed objects compile directly into one encoder row model and flush immediately; no mutable rows, cells, or column objects.
- **`streamRead`:** SAX parsing yields one-based sparse value arrays without a document graph. Styles can be cached to decode dates. Shared-string/style caches still have their own memory costs. Native writable streams batch 64 KiB of XML and propagate destination backpressure through iterable row sources.
- `WorkbookWriter` / `WorkbookReader` are internal package/stream coordinators.

## Legacy removal

There is no `lib/doc` directory, mutable Workbook/Worksheet/Row/Cell/Column API,
legacy CSV attachment, namespace entry, synchronous feature loader, or legacy
`index.d.ts`. Reusable range, name, note, image and enum helpers live in `lib/model`;
structural feature types live in `lib/model/schema.ts` and encoder types in
`lib/model/xlsx-model.ts`.

Historical tests tied to the unsupported mutable API are retired. Encoder tests,
public builder/CSV tests, range/name/anchor tests, and saved compatibility fixtures
remain. Golden XML covers values, styles, sparse edits, merges, notes, validation,
conditional formatting, tables, images, printing and sheet order. The fixture sweep
also checks 33 existing XLSX projections against the pre-rewrite output.

Validation and defined-name storage use compact rectangles. Buffered XLSX exposes
only its model loader and buffer writer internally; Node owns filesystem wrappers.
`StreamBuf`, the old ZIP facade, the custom event emitter and the cell matrix are
removed. Tests import modules directly and use native Vitest APIs.

See [the four independent cleanup comparisons](scripts/bench/internal-cleanup-results.md),
[the buffered comparison](scripts/bench/doc-bridge-results.md) and
[the legacy removal comparison](scripts/bench/legacy-removal-results.md).

## Bundle size (indicative)

From `pnpm test:browser-bundle` (esbuild minify, **write-only** builder path;
CSV enabled, no Node polyfills):

| Build                       | Size (0.2.0)         |
| --------------------------- | -------------------- |
| Single-file minified        | **268.0 KiB**        |
| Single-file gzip            | **76.8 KiB**         |
| Code-split write-only entry | **141.6 KiB**        |
| Code-split write-only total | 266.0 KiB (31 files) |

The core and unused `/protection` imports include no password hashing crypto. Code-split entry size excludes shared static dependencies; use the single-file size when comparing complete bundles. See [the protection comparison](scripts/bench/protection-bundle-results.md).

Re-measure after dependency or encoder changes:

```bash
pnpm test:browser-bundle   # also writes build/browser-smoke/sizes.txt
pnpm test:esm
```

## Key modules

| Path                              | Purpose                              |
| --------------------------------- | ------------------------------------ |
| `excel.ts` / `node.ts`            | Package entries                      |
| `lib/builder/`                    | Op-log builder                       |
| `lib/model/types.ts`              | Plain workbook types                 |
| `lib/compile/`                    | Ops/plain ↔ encoder models           |
| `lib/xlsx/`                       | Encode/decode OOXML                  |
| `lib/csv/public.ts`               | Public CSV API                       |
| `lib/stream/xlsx/stream-write.ts` | Node `streamWrite`                   |
| `lib/stream/xlsx/stream-read.ts`  | Node `streamRead`                    |
| `lib/model/`                      | Structural types and feature helpers |
| `excel.d.ts`                      | Public TypeScript types              |
