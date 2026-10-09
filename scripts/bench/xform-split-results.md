# Directional OOXML transform split

Baseline: `7d2f2803`, after the five independent rewrites. Same esbuild browser ESM entries and options; sizes are KiB (1,024 bytes). Single-file bundles include automatically loaded features. No public API changes.

| Browser entry | Before min / gzip | After min / gzip | Saved min / gzip |
| ------------- | ----------------: | ---------------: | ---------------: |
| load          |    182.11 / 49.89 |   130.24 / 36.92 |    51.88 / 12.97 |
| write         |    218.26 / 59.15 |   157.12 / 45.56 |    61.14 / 13.59 |
| readRows      |     33.84 / 13.60 |    33.84 / 13.60 |      0.00 / 0.00 |

Full load shrinks 28.5% minified / 26.0% gzip; basic write shrinks 28.0% / 23.0%.

## Split bundles

Initial download includes the static import closure. Reachable total includes static and dynamic imports; gzip sums each chunk compressed separately.

| Entry    | Initial before → after min / gzip | Reachable total before → after min / gzip |
| -------- | --------------------------------: | ----------------------------------------: |
| load     |    128.85 / 40.12 → 92.40 / 29.08 |           180.55 / 56.83 → 128.82 / 41.60 |
| write    |    128.96 / 39.98 → 99.55 / 33.34 |           216.68 / 67.25 → 157.68 / 52.53 |
| readRows |     33.63 / 14.31 → 33.63 / 14.31 |             33.63 / 14.31 → 33.63 / 14.31 |

Esbuild emits some chunks for unused barrel exports that the entry never reaches. Counting every emitted file overstates the dependency graph. `readRows` remains exactly 34,433 bytes minified / 14,658 gzip across its six reachable chunks, although all emitted files grow from 139,292 to 143,121 bytes. Deployment tooling may still copy these unused files; consumers do not import them.

## Runtime

Node, 1,000 rows × 8 cells, one shared XLSX fixture; five warmup pairs, 31 measured pairs with alternating order. Median milliseconds:

| Operation | Before | After |
| --------- | -----: | ----: |
| write     |  5.651 | 5.684 |
| load      |  8.860 | 8.506 |
| readRows  |  4.229 | 4.224 |

These short timings show no substantial speed change; they are not a performance guarantee. Separate-process measurements and all raw paired samples are retained in the JSON report.

## Implementation and verification

- Parsing/reconciliation moved to `lib/xlsx/parser/`; preparation/rendering remains in `lib/xlsx/xform/`. Shared state and types avoid duplicating model contracts. Writer-only XML templates, caches and style construction are absent from parsers.
- Automatic drawings, comments, tables and conditional formatting remain supported. Node streaming uses the corresponding direction.
- Existing XML fixtures test separate writer/parser factories. All 976 tests across 102 files pass; source, specification and public declarations type-check.
- Directional bundle guards, browser/ESM smoke checks, packed-package runtime checks and asset budgets pass.
- Tradeoff: separate classes require coordinated updates when OOXML behavior changes. Shared fixture tests cover both directions.

## Reproduction

Build the baseline into a separate `dist` directory, then build the candidate. From the repository root:

```sh
node scripts/bench/five-rewrites.mjs xform-before build/xform-split/before/dist
node scripts/bench/five-rewrites.mjs xform-after
node scripts/bench/xform-split.mjs build/xform-split/before/dist dist
node scripts/smoke-directional-bundles.mjs
```
