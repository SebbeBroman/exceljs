# Local fast-csv in ExcelJS — browser comparison

Measured 2026-10-02 with Node v24.21.0, esbuild 0.28.1, headless Chrome 154 on this machine.

ExcelJS now links `fast-csv` to `../fast-csv/packages/fast-csv`. Its text CSV APIs and workbook view use `fast-csv/browser`. The local fork also exports a browser formatter that reuses `RowFormatter` without streams. Legacy Node stream CSV imports use the ESM namespace. Nothing was committed, pushed, or published.

Baseline: the ExcelJS source before these changes, registry fast-csv 5.0.7, and stream-browserify/util/buffer/process/events/string_decoder browser polyfills. Candidate: current ExcelJS and the local fast-csv browser entry. CSV is enabled in all bundles; there are no CSV stubs. Both variants use the existing ExcelJS fs/module browser shims.

Minified single-file bundles; KiB = 1024 bytes. Gzip uses Node zlib defaults; Brotli sizes and esbuild module metadata are in the JSON results.

| Imports / capability          | Registry raw KiB | Local raw KiB | Registry gzip KiB | Local gzip KiB | Gzip reduction |
| ----------------------------- | ---------------: | ------------: | ----------------: | -------------: | -------------: |
| parseCsv                      |            170.2 |          26.4 |              52.1 |            9.9 |          81.0% |
| stringifyCsv                  |            170.7 |          20.7 |              52.3 |            8.2 |          84.2% |
| csv (parse + stringify)       |            171.6 |          32.2 |              52.6 |           11.7 |          77.8% |
| viewWorkbook (CSV + XLSX)     |            192.0 |          48.1 |              60.6 |           18.2 |          69.9% |
| workbook + writeBuffer        |            473.8 |         334.0 |             137.0 |           96.0 |          29.9% |
| workbook + writeBuffer + load |            477.8 |         338.0 |             138.3 |           97.2 |          29.7% |

All local bundles include zero Node polyfill modules; the registry bundles include 22. The registry browser builds fail to resolve Node builtins without these polyfills. Named imports `parseCsv` and `stringifyCsv` let the local ESM graph drop the unused half of CSV. Importing the `csv` object retains both.

Browser workload: 2,000 rows × 8 columns, two warmups, median of seven measured runs, alternating baseline/candidate order. CSV parse uses an identity map to isolate parsing from ExcelJS date inference. XLSX buffers are prepared before read measurements.

| Operation                | Registry, microtask scheduler (ms) | Local (ms) |
| ------------------------ | ---------------------------------: | ---------: |
| csv parse (map identity) |                                7.9 |        0.8 |
| csv stringify            |                                6.9 |        2.2 |
| csv view                 |                                7.3 |        1.1 |
| xlsx write               |                               38.1 |       37.4 |
| xlsx load                |                               26.8 |       27.4 |

The standard process/browser nextTick polyfill schedules timers. In the unmodified baseline this makes CSV export take roughly 9.2 seconds for this fixture. The table uses an additional baseline run with nextTick implemented through queueMicrotask to avoid presenting that timer artifact as an intrinsic parser/formatter improvement. Both runs are retained in results.json. The alternate scheduler applies to the CSV helpers; the workbook-view and XLSX bundles remain the same. Small XLSX timing differences are within observed variation.

Verification passed: both project builds; 107 ExcelJS builder tests; 47 fast-csv parser/formatter browser tests; native ESM/file/stream smoke; browser bundle smoke; browser CSV fixture parity, malformed-input rejection, formatter-option parity, CSV view parity, and XLSX read/write. The browser starts without global process or Buffer.

The actual code-split write-only smoke has a 206.4 KiB initial entry and 331.6 KiB total across 29 files, including optional CSV. The smoke now clears its output directory before measuring to exclude stale chunks. The comparison JSON also records per-chunk raw and gzip sizes.

Rerun from the ExcelJS checkout (build the sibling fast-csv packages first):

```sh
pnpm build
npm install --prefix build/local-fast-csv/polyfills --no-audit --no-fund --ignore-scripts fast-csv@5.0.7 stream-browserify util buffer process events string_decoder
node scripts/bench/local-fast-csv.mjs
node scripts/bench/local-fast-csv.mjs --microtask-only
node scripts/bench/local-fast-csv.mjs --sizes-only
```

The runner preserves the baseline source snapshot in build/local-fast-csv/baseline; if absent it recreates it from ExcelJS HEAD. Generated bundles, metafiles, polyfill dependencies, and results live under the ignored build/local-fast-csv directory. This is a local experiment: the dependency link assumes the sibling checkout exists.
