# Architecture (0.2)

Short map of how `@sebbebroman/exceljs` turns builder calls into `.xlsx` bytes.

## Public surface

| Entry                             | Role                                                                              |
| --------------------------------- | --------------------------------------------------------------------------------- |
| `@sebbebroman/exceljs`            | Browser-safe: `workbook`, `writeBuffer`, `load`, enums                            |
| `@sebbebroman/exceljs/csv`        | Optional CSV parse/stringify and views                                            |
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
- **Encode:** `writeBuffer` → `lib/xlsx/xlsx-writer.ts` → zip (default deflate **level 1** for speed; override with `{ zip: { level: 6 } }` for smaller files). Optional features load automatically through dynamic imports.
- **Styles:** basic builder exports use a standalone minimal stylesheet encoder. The full style manager loads dynamically when styles are enabled; differential formats load automatically when needed. There is no caller codec selection.
- **Utilities:** internal helpers are named functions; column conversion uses arithmetic with bounded memoization, and the address cache has a fixed limit.
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
   XLSX.decode                 lib/xlsx/xlsx-reader.ts + parsers
        │
        ▼
   reconciled encoder model
        │
        ▼
   plain { meta, sheets }     lib/compile/xlsx-model-to-plain.ts
```

Edit loop: `workbook(await load(buf)).sheet(…).cell(…).writeBuffer()`. Loaded tables reconstruct their top-left reference and body values from the worksheet, including totals metadata. Reader and writer orchestration and OOXML transforms have separate dependency trees. `lib/xlsx/parser/` contains parsing and reconciliation; `lib/xlsx/xform/` contains preparation and rendering. A small shared `xform-state.ts` holds model/reset utilities and the child contract. Parser type imports from writer modules are erased, so loading retains no XML rendering engine. Optional feature parsers and writers load automatically. `scripts/smoke-directional-bundles.mjs` guards the separation.

Model-loading `load` runs in two phases: package parts first (workbook,
SST, styles, rels, …), then sheets — each sheet flows parse → reconcile-ready
without waiting on other sheets. Sheets eligible for the fused fast path
(`lib/xlsx/parser/sheet/fast-sheet-data.ts`) parse + reconcile `sheetData`
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

Named helpers in `/csv` (`csv.parse` / `csv.stringify` / `viewCsv` / `readCsvRows`). Core builders and XLSX readers import no CSV parser or formatter. Node `readCsvFile` / `writeCsvFile` wrap the optional helpers with `fs`. Advanced OOXML transforms load automatically when a read/write encounters their feature data.

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
CSV excluded; advanced features automatic, no Node polyfills):

| Build                       | Size (0.2.0)         |
| --------------------------- | -------------------- |
| Single-file minified        | **237.8 KiB**        |
| Single-file gzip            | **65.8 KiB**         |
| Code-split write-only entry | **141.3 KiB**        |
| Code-split write-only total | 235.4 KiB (28 files) |

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

See [the CSV bundle comparison](scripts/bench/csv-split-results.md).

## Deterministic regression budgets

`pnpm test:bundle-size` builds fresh browser bundles for basic writing, full
`load`, `readRows` and `viewWorkbook`. Each entry checks minified and gzip sizes
for a single-file bundle, its complete initial static import closure, and all
reachable static/dynamic chunks. Unreachable chunks emitted for unused barrel
exports do not count. The budgets in `scripts/browser-bundle-budgets.json` leave
roughly 5–8% headroom; review intended growth before raising them. Dependency
checks also forbid parsers in writing, rendering in loading, and full transforms,
CSV or protection crypto in lightweight XLSX reads. The same checks run through
`test:browser-bundle` in CI and before publishing. Measurements are saved to
`build/bundle-regression/sizes.json`; they never depend on an earlier smoke run.

`pnpm test:work-budgets` runs clock-free work regression tests, also included in
`pnpm test` across the CI Node/OS matrix:

- Basic writing: one cell render per populated cell, one buffered ZIP pass, and
  no full style-manager calls when styles are disabled.
- Full loading: at most four style lookups for two repeated styles, independent
  of row count, and no classic per-cell parser calls on an eligible sheet.
- Dense row slicing: parsing a two-row slice visits exactly 16 cell tags in an
  eight-column fixture, even as the rest of the worksheet grows.
- Address caching: immediate lookups reuse objects; a wide scan evicts old
  objects rather than retaining the entire scanned address space.

The workload tests use 32- and 128-row fixtures and verify returned data. Spies
and a wrapper around the real SAX parser count work only in tests; production
has no instrumentation. These metrics catch repeated passes, lost caching and
missed early exits. They do not measure CPU time, GC pauses or compression
speed; timing benchmarks still serve that purpose. Bundle bytes can change with
toolchain versions, so dependency upgrades may require reviewing the budgets.
