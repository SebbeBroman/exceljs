# DocWorkbook bridge removal: before and after

Measured locally on 2026-10-08 with Node v24.21.0 and Chrome 154. The saved baseline is the clean pre-rewrite build in `build/doc-bridge-before/dist`; the comparison runs both builds in the same browser session.

This report records the buffered rewrite before the subsequent [complete legacy document removal](legacy-removal-results.md). Its suite counts and remaining-legacy description refer to that intermediate state.

## What changed

`writeBuffer(builder)` compiles optimized operations directly into the existing OOXML encoder model. `writeBuffer(snapshot)` imports row/cell placements directly, avoiding the former per-cell op-log. `load()` projects reconciled encoder models into the public snapshot without document hydration. Shared value/style conversion helpers keep the legacy classes and new model compiler consistent. Concurrent worksheet parsing now explicitly restores workbook order.

Tables, image anchors, notes, ranges and defined names still reuse feature helpers. Read image anchors retain their coordinate accessors through a dimension-only host, rather than retaining a document workbook. Node streaming and historical document tests still use legacy classes; this change removes the **buffered bridge**, not all of `lib/doc`.

The public API is unchanged. Direct snapshot writes additionally retain explicit row heights, hidden flags and row styles, which the old snapshot-to-ops adapter discarded. Snapshots do not implicitly insert column headers on write; their existing row placements are authoritative.

## Browser throughput

One string and seven numbers per row; shared strings enabled, styles disabled, existing level-1 compression. Three warmups and eleven measured runs per operation, alternating before/after order. GC is requested between runs, outside the measured operation. Reads use exactly the same baseline XLSX bytes. Builder writes include builder creation; snapshot writes use a preloaded snapshot; roundtrip is builder write plus full load.

| Rows × 8 | Operation      | Before (ms) | After (ms) | Time reduction |
| -------- | -------------- | ----------: | ---------: | -------------: |
| 2,000    | builder write  |        13.3 |       12.6 |           5.3% |
| 2,000    | snapshot write |        22.5 |       15.5 |          31.1% |
| 2,000    | full read      |        20.9 |       17.0 |          18.7% |
| 2,000    | roundtrip      |        32.8 |       29.2 |          11.0% |
| 20,000   | builder write  |       117.7 |      118.2 |          -0.4% |
| 20,000   | snapshot write |       200.1 |      137.8 |          31.1% |
| 20,000   | full read      |       157.2 |      136.4 |          13.2% |
| 20,000   | roundtrip      |       280.4 |      256.9 |           8.4% |

The 20,000-row builder write is effectively unchanged (0.4% slower in this run). Snapshot writes improve by about 31%, full reads by about 13%, and roundtrips by about 8%. Compression remains a large part of write time. These measurements describe the tested workload, not every supported feature.

All four before/after writer-reader combinations reproduced every cell at both sizes. The generated XLSX files have identical sizes: 91,677 bytes at 2,000 rows and 866,092 bytes at 20,000 rows.

## Browser bundles

Single-file, minified ESM builds with esbuild; CSV/optional feature code remains real, and no Node polyfills are added. These totals include the whole single-file bundle, not only an entry chunk.

| Bundle     | Before raw KiB | After raw KiB | Before gzip KiB | After gzip KiB |
| ---------- | -------------: | ------------: | --------------: | -------------: |
| write      |          334.5 |         297.2 |            96.4 |           86.6 |
| read-write |          338.5 |         301.1 |            97.6 |           88.0 |

Write-only gzip falls by about 9.7 KiB (10.1%); read/write by about 9.6 KiB (9.8%). The regular browser smoke fixture measures 297.3 KiB raw / 86.7 KiB gzip; its source entry differs slightly from this comparison harness.

## Peak process memory

100,000 × 8 cells, three fresh Node processes per build and operation, alternating order. Reads use the same saved XLSX. Writes include dense grid creation, builder creation and export. This is median **peak process RSS**, including the runtime, input and output; it is not an isolated allocation count or post-GC retained heap.

| Operation | Before MiB | After MiB | Reduction |
| --------- | ---------: | --------: | --------: |
| read      |      698.2 |     568.0 |     18.7% |
| write     |      689.8 |     603.3 |     12.5% |

An earlier memory run measured read peaks of 634.4 → 567.8 MiB (10.5% lower), and write peaks of 690.6 → 607.6 MiB (12.0% lower). Peak RSS varies with garbage collection; across both runs, reads used roughly 10–19% less peak memory and writes roughly 12–13% less.

## Compatibility and validation

- Full suite: **1,193 passed, 1 skipped**, across 156 files (baseline: 1,182 passed, 1 skipped).
- Eleven added tests cover exact uncompressed XLSX-part parity, legacy snapshot parity, shared/inline strings, dates, formulas, rich text, errors, keyed columns, style composition, merges, comments, validations, conditional formatting, protection, tables/totals, images, default metadata, sparse cells, addressless cells, randomized edits, existing merged-blank-row fixtures, snapshot metadata and input immutability.
- Public snapshot comparison: **33 existing XLSX fixtures match**, including one invalid-XML fixture rejected identically. Only the private image-anchor worksheet backpointer is excluded from normalized comparisons. `huge.xlsx` matched in the initial run and was omitted from the repeat fixture sweep to avoid retaining its large intermediate results.
- `test-issue-1842.xlsx` is excluded: its validation covers the whole Excel grid (`A1:XFD1048576`), and the unchanged baseline parser stalls expanding that range before projection. This rewrite does not change validation parsing.
- Library typecheck, public declaration checks, touched-file lint/format checks, native ESM/file/stream smoke, browser bundle smoke and asset-size budgets pass.
- The separate spec typecheck fails in both builds with the **same 3,152 diagnostics and no new diagnostics**. Baseline was checked from an archived clean source tree with the same dependencies.

## Reproduce

Save the baseline before changing sources:

```sh
pnpm build
mkdir -p build/doc-bridge-before
cp -R dist build/doc-bridge-before/dist
```

After applying the rewrite:

```sh
pnpm build
node scripts/bench/doc-bridge.mjs
node scripts/bench/doc-bridge-memory.mjs
node scripts/bench/doc-bridge-fixtures.mjs
pnpm test
pnpm typecheck
pnpm typecheck:public
pnpm test:esm
pnpm test:browser-bundle
```

Set `CHROME_PATH` for a different Chrome installation when running the browser comparison. Raw samples, bundle metadata and XLSX fixture outputs are in ignored `build/doc-bridge-comparison/`; the initial three-library comparison and pre-rewrite build are in `build/doc-bridge-before/`. The phase profiler also supports the direct model path: `node scripts/bench/browser-profile.mjs --label direct-model`.

This report records the local rewrite and its measurements.
