# CSV split comparison

Baseline: commit `96bc79bc` (0.2.0). Same esbuild minified browser ESM workloads and dependency versions. Advanced XLSX handling remains automatic; the codec experiment has been reverted.

| Workload                | Before minified / gzip (KiB) | After minified / gzip (KiB) | Saved minified / gzip (KiB) |
| ----------------------- | ---------------------------: | --------------------------: | --------------------------: |
| Basic XLSX writing      |                 267.9 / 76.7 |                237.7 / 65.8 |                 30.2 / 11.0 |
| XLSX row reading        |                  44.1 / 16.8 |                 34.1 / 13.7 |                  10.0 / 3.1 |
| Full-model XLSX loading |                 214.3 / 59.3 |                214.3 / 59.3 |                   0.0 / 0.0 |

Full model loading already excluded CSV in a single-file bundle. Basic writing saves 11.0 KiB gzip (about 14%). XLSX row reading also sheds its CSV parser.

## Retained API change

Import `csv`, `parseCsv`, or `stringifyCsv` from `@sebbebroman/exceljs/csv`. Replace `builder.csv(options)` with `csv.stringify(builder, options)`. The default is the first sheet; use `{sheetName}` to preserve an active-sheet selection.

CSV views and row reading use the separate named `viewCsv` / `readCsvRows` helpers from `/csv`. Keeping these separate avoids retaining the view machinery through the `csv` object. Node CSV file helpers still work. Unused CSV imports disappear in both single-file and split browser bundles.

`load(bytes)` and `writeBuffer(model)` still detect drawings, comments, tables, pivots and conditional formatting automatically, without requiring callers to know a workbook's contents or supply codecs. Existing feature and round-trip coverage is retained. Dynamic imports can delay these features in split bundles; a single-file bundle includes them.

The complete test suite, saved pre-rewrite compatibility fixtures, public declarations, native ESM and browser smoke checks validate the retained behavior.

Reproduce with `pnpm build`, then `node scripts/bench/csv-split.mjs /path/to/baseline/dist`. Raw measurements: [JSON](csv-split-results.json).
