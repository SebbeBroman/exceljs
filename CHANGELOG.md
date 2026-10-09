# Changelog

## 0.2.0

### Breaking: optional CSV

- CSV helpers and types move from core to `@sebbebroman/exceljs/csv`. Replace `builder.csv(opts)` with `csv.stringify(builder, opts)`; use `sheetName` to select the previous active sheet.
- CSV views and dense row reads move to `viewCsv` / `readCsvRows`. Core `viewWorkbook` / `readRows` support XLSX; Node CSV file helpers stay available.
- Unused CSV imports are eliminated from single-file browser bundles. Advanced XLSX features continue to load automatically.

See [MIGRATION.md](./MIGRATION.md#upgrading-from-01x-to-020-optional-csv) for details.

### Breaking: opt-in sheet protection hashing

The core builder no longer hashes passwords. Import the synchronous factory from `@sebbebroman/exceljs/protection` and pass its prepared model to `.protect()`.

Before (0.1.x):

```ts
import {workbook} from '@sebbebroman/exceljs';

const builder = workbook().sheet('Data').protect('secret', {spinCount: 100000});
```

After (0.2.0):

```ts
import {workbook} from '@sebbebroman/exceljs';
import {sheetProtection} from '@sebbebroman/exceljs/protection';

const builder = workbook()
  .sheet('Data')
  .protect(sheetProtection('secret', {spinCount: 100000}));
```

- `.protect(password, options)` is removed from both sheet and workbook builders. Hashing now happens when `sheetProtection()` is called, rather than when the workbook is written. The default spin count remains 100,000.
- For passwordless protection, replace `.protect()` or `.protect(undefined, options)` with `.protect({sheet: true, ...options})`.
- Plain models replace `sheet.protect = {password, options}` with `sheet.sheetProtection = sheetProtection(password, options)`. `ProtectConfig` is removed; use the exported `SheetProtection` type.
- Prepared hashes are retained by `build()` and reused across writes. Loaded protection models can still be edited and rewritten through the core without importing crypto.
- SHA-512 and the hashing helper are absent from basic browser write bundles. Unused `/protection` imports are tree-shaken out even in single-file builds. The measured write bundle saves 8.3 KiB minified / 3.8 KiB gzipped.

### Breaking: compact loaded validation ranges

Loaded `dataValidations` retain range keys such as `A1:B20` instead of expanding one entry per cell. Check a rule's range coverage rather than assuming it has an `A1` key. Single-cell keys remain supported. Whole-grid validation and defined-name ranges can now load and round-trip without per-cell expansion.

### Fixes

- Loaded tables retain their top-left reference, body values, and cached/custom totals, fixing load/edit/rewrite failures (`Table must have ref`). Offset and headerless tables, dates and formulas are covered by round-trip tests.
- Buffered and Node streaming reads/writes operate directly on XLSX models and sparse values. Removed the internal document bridge, mutable document classes, `lib/doc`, old compiler bridges and legacy declarations.
- Node streaming uses native writable streams and 64 KiB XML batches. Declarative row sources and `await sheet.rows(source)` honor destination backpressure. Synchronous `sheet.row(values)` loops can still queue rows.

### Smaller browser bundles and internal cleanup

- Buffered reader/writer orchestration and OOXML parsing/rendering have separate dependency trees. A full `load` bundle drops rendering; a write bundle drops parsing. Drawings, comments, tables and conditional formatting still load automatically.
- Basic exports use a standalone minimal stylesheet encoder when styles are disabled. The full stylesheet engine remains available when needed.
- Removed unused XLSX file/stream wrappers, ZIP facade, custom stream buffers, event emitter and cell matrix. Utilities use named exports, with unused helpers removed.
- Column labels use arithmetic conversion and a small cache; address caching is bounded. Far-right lookups no longer allocate all 16,384 column labels.
- Defined names and validations use compact rectangle geometry. This improves large-range work and memory use; individual-cell defined-name insertion can be slower.

The final parsing/rendering split independently reduces both full load and basic write bundles by about 28% minified, with roughly unchanged measured runtime. Current browser ESM fixtures (esbuild, single-file minification, advanced XLSX features automatic):

| Entry          |  Minified |     Gzip |
| -------------- | --------: | -------: |
| Basic write    | 157.1 KiB | 45.6 KiB |
| Full `load`    | 130.2 KiB | 36.9 KiB |
| `readRows`     |  33.8 KiB | 13.6 KiB |
| `viewWorkbook` |  33.3 KiB | 13.5 KiB |

These are fixture-specific sizes, not guarantees for every application. See the independent [legacy removal](https://github.com/SebbeBroman/exceljs/blob/master/scripts/bench/legacy-removal-results.md), [internal cleanup](https://github.com/SebbeBroman/exceljs/blob/master/scripts/bench/internal-cleanup-results.md), [CSV split](https://github.com/SebbeBroman/exceljs/blob/master/scripts/bench/csv-split-results.md), [five rewrites](https://github.com/SebbeBroman/exceljs/blob/master/scripts/bench/five-rewrites-results.md) and [transform split](https://github.com/SebbeBroman/exceljs/blob/master/scripts/bench/xform-split-results.md) reports for baselines, workloads and tradeoffs.

### Regression checks and release tooling

- Browser bundle tests enforce minified/gzip budgets for write, full load, `readRows` and `viewWorkbook`, covering single-file output, the initial static import closure and all reachable split chunks. Unreachable emitted chunks do not count. Dependency checks catch accidental CSV, crypto or opposite-direction transforms.
- Clock-free work tests count cell renders, ZIP passes, style lookups and parsed cell tags, and check address-cache eviction. These catch repeated work, lost caching and missed early exits; they do not measure execution speed.
- Tests use native Vitest imports/assertions. Source, specification and isolated public declaration checks cover all four package entries. The maintained suite has 983 tests across 103 files.
- `pnpm release:check` runs the complete publish validation. `pnpm release:pack` validates and creates a checked local tarball; `pnpm release` publishes with the same validation hook.

See [MIGRATION.md](./MIGRATION.md) for the public API and upgrade guide.
