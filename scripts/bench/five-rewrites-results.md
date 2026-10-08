# Five independent internal rewrites

Baseline: `9dd8ce69` (CSV split committed). Each row compares with the preceding stage. Sizes are KiB; signed values are deltas. Measurements use the same esbuild browser ESM workload and options. Initial size includes the complete static import closure; total single-file size includes lazy code.

| Rewrite                      | Commit        | Write min/gzip Δ | Initial write min/gzip Δ | Load min/gzip Δ |
| ---------------------------- | ------------- | ---------------: | -----------------------: | --------------: |
| Loaded table conversion      | `a19ab84d`    |    +0.10 / +0.05 |            +0.00 / +0.01 |   +0.48 / +0.21 |
| Reader/writer split          | `1c1b1fcc`    |   -20.53 / -7.07 |           -15.09 / -5.22 |  -30.77 / -8.96 |
| Minimal styles encoder       | `59e2f82e`    |    +2.51 / +0.88 |           -19.60 / -3.78 |   -0.28 / -0.08 |
| Standalone utilities         | `f514abe3`    |    -1.23 / -0.41 |            -1.03 / -0.42 |   -1.32 / -0.46 |
| Bounded address/column cache | `this commit` |    -0.31 / -0.08 |            -0.34 / -0.08 |   -0.31 / -0.10 |

## Workload timings

Median milliseconds, 1,000 rows × 8 cells, three warmups and eleven measurements. These short Node runs are indicative; differences of a few percent are normal variation.

| Stage     | Write |  Load | readRows |
| --------- | ----: | ----: | -------: |
| baseline  |  6.29 |  9.68 |     4.37 |
| tables    |  6.26 |  9.54 |     4.64 |
| split     |  6.47 | 10.08 |     4.55 |
| styles    |  6.24 | 10.14 |     4.61 |
| utilities |  6.50 |  9.89 |     4.61 |
| addresses |  6.32 |  9.98 |     4.67 |

## Outcomes and tradeoffs

- Tables: write → load → rewrite previously threw `Table must have ref`; now works, including headerless/offset tables, dates, formulas and cached custom totals. The measured small-table load/edit/write loop takes about 1.2 ms. This is a correctness fix, with a small load-bundle increase.
- Split: reader and writer have separate dependency graphs; write imports no buffered unzip orchestration, and load imports no buffered ZIP writer. Shared OOXML transform classes still contain both parsing and rendering methods.
- Styles: basic exports no longer construct the full style manager. The initial write download falls, but a single-file bundle includes both encoders and grows slightly. Differential formats still load automatically, even with `useStyles: false`. A focused 1,000-construction/serialization benchmark fell from 10.80 ms to 0.11 ms; ordinary export timings are dominated by cells and ZIP.
- Utilities: named functions replace the retaining default object, and unused helpers are deleted. No public export changes.
- Cache: arithmetic replaces bulk filling all 16,384 columns; only the first 256 column labels are memoized. The address cache is capped at 2,048 entries. Row-only/column-only print-title references remain supported; invalid nonintegral column numbers now throw.

Fresh-process cache measurements (`--expose-gc`):

| Metric                                           | Before cache rewrite |              After |
| ------------------------------------------------ | -------------------: | -----------------: |
| Cold XFD lookup (ms)                             |                5.940 |              0.042 |
| 100,000 wide-column lookups (ms)                 |                1.145 |              3.944 |
| Retained column heap (KiB)                       |               1338.5 | approximately zero |
| Retained address heap after 100 × 100 scan (KiB) |               1809.0 |              251.8 |

GC deltas near zero can be negative due to measurement noise. The wide-column stress test is deliberately unfavorable to the bounded cache; it trades a few milliseconds per 100,000 conversions for lower startup work and bounded retention. A bounded `Map` prototype showed a modest load slowdown; the final implementation uses object lookups with FIFO eviction instead.

## Reproduce

Run `pnpm build`, then `node scripts/bench/five-rewrites.mjs <stage> [dist-directory]`. Compare archived dist directories to isolate each change. Focused scripts: `node scripts/bench/minimal-styles.mjs [dist-directory]` and `node --expose-gc scripts/bench/address-cache.mjs [dist-directory]`. Raw sizes, ranges and timings are in `five-rewrites-results.json`.

Validation: all historical OOXML fixtures remain unchanged. Table compatibility snapshots ignore only newly recovered placement/body/totals fields; dedicated round-trip tests verify those values. Source/spec/public type checks, native ESM, browser/optional-entry smokes and asset budgets passed.
