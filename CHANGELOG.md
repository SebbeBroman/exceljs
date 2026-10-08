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

### Internal changes

Buffered XLSX reads and writes now use direct models rather than the legacy document graph. Removed legacy document classes, unused XLSX wrappers and ZIP facade, and custom stream buffers. Validation and defined-name ranges stay compact. Node streaming uses standard streams with backpressure. Tests use native Vitest imports.

See [MIGRATION.md](./MIGRATION.md) for the public API and upgrade guide.
