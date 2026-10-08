# Four independent internal cleanup measurements

Measured on 2026-10-08, Node v24.21.0, macOS arm64. The starting commit is
`051549f6` (the completed document-graph removal). Each comparison uses adjacent
saved stages, so later changes do not get credited for earlier improvements.

| Change                                          | Before                                              | After                                           | Result                                             |
| ----------------------------------------------- | --------------------------------------------------- | ----------------------------------------------- | -------------------------------------------------- |
| Compact validation ranges, 100k-cell rectangle  | 27.44 ms; 100,000 model keys; 90.72 MiB RSS         | 0.080 ms; 1 key; 54.97 MiB RSS                  | Avoids per-cell expansion                          |
| Compact defined names, 100k-cell rectangle      | 18.12 ms; 140.84 MiB RSS                            | 0.088 ms; 54.75 MiB RSS                         | Avoids building and collapsing a cell matrix       |
| Remove unused XLSX/ZIP wrappers, browser bundle | 295,254 bytes minified; 86,600 bytes gzip           | 283,226 bytes minified; 82,629 bytes gzip       | 12,028 minified bytes and 3,971 gzip bytes removed |
| Native streams, 100k × 8 rows to a file         | 527.79 ms; 129.23 MiB RSS                           | 501.78 ms; 110.61 MiB RSS                       | About 5% faster in this sample                     |
| Native streams, 100k × 8 rows to a slow sink    | 1,773.28 ms; 129.27 MiB RSS; 3,833,201 queued bytes | 1,560.38 ms; 99.38 MiB RSS; 24,457 queued bytes | Backpressure bounds destination queue              |
| Native test imports/assertions, full suite      | 19.90 s; 654 test type errors                       | 13.59 s; 0 test type errors                     | About 32% faster; complete spec typecheck passes   |

The wrapper comparison starts **after** the range rewrite: its exact before sizes
are **295,254 bytes minified / 86,600 bytes gzip**, across 180 bundled modules.
After wrapper removal there are 176 modules. The final production browser bundle
in this harness is 282,858 bytes minified / 82,518 bytes gzip; the small further
reduction belongs to the native-stream cleanup's removal of unused utility exports.
The standard browser smoke uses a slightly different entry and reports
276.3 KiB minified / 80.6 KiB gzip.

**1. Compact rectangles.** Defined names store rectangle references instead of
CellMatrix entries. Validation parsing retains compact range keys and subtracts
overlapping rectangles to preserve last-rule-wins coverage. Sorted coalescing
handles adjacent cells without pairwise merge scans. Shared rule objects reuse
normalization work. Whole-grid `A1:XFD1048576` references now complete in about
0.08–0.11 ms in the geometry benchmark. Under the same isolated process budgets,
the old validation parser exceeded its one-second timeout and the old name path
failed under its 128 MiB heap limit. These are geometry timings, not XLSX load times.
A separate regression loads the previously stalled `test-issue-1842.xlsx` and
checks whole-grid validation/name write-load-write behavior.

The dense-input case is deliberately measured too:

| 100,000 separate cell references | Before time / RSS      | After time / RSS       |
| -------------------------------- | ---------------------- | ---------------------- |
| Validation XML coalescing        | 123.18 ms / 252.77 MiB | 105.99 ms / 221.73 MiB |
| Defined-name add + normalization | 58.02 ms / 209.78 MiB  | 69.12 ms / 116.47 MiB  |

Adding individual name cells is about **19% slower** in this synthetic case,
while peak RSS falls about **44%**. The major gains apply to range inputs; this
rewrite is not a universal speedup. Loaded validation records now use range keys,
so consumers must inspect coverage rather than assume every cell has its own key.
That observable shape change is documented in `MIGRATION.md`.

**2. XLSX wrappers.** Removed the unused internal `readFile`, `read`, `writeFile`,
`write`, `createInputStream`, stream finalizer, and legacy host-model hook. Deleted
the old `zip-stream` facade and its unused filesystem existence helper. The public
Node filesystem APIs continue to use the maintained buffer API. This change's
measured benefit is bundle size and fewer internal paths; no throughput claim is
made for it.

**3. Native streams.** Removed StreamBuf, the bespoke event emitter and their
inheritance/no-op helpers. Streaming ZIP entries are native Writable streams with
64 KiB batching, drain handling and error propagation. Unused stream-source/file
ZIP protocols are removed. Declarative sheets finish sequentially. Destination
completion observes only the writable side, allowing PassThrough callers to read
a small completed workbook afterward.

At 50k × 8 rows, file writing measured 277.52 → 270.71 ms and slow-sink writing
905.07 → 798.65 ms. The slow sink's peak queue grows from 1,919,490 to 3,833,201
bytes when the old writer doubles its input; the new writer stays at 24,457 bytes
in both cases. Both versions emitted exactly the same byte counts for each fixed
metadata scenario. Compatibility tests separately compare every uncompressed
XML/media part, avoiding ZIP timestamp noise.

Backpressure applies to declarative row iterables and `await sheet.rows(source)`.
The synchronous `sheet.row(values)` method cannot wait for drain. Shared strings,
styles, and callback sheets written out of order retain their own buffering costs;
the queue measurement does not claim to bound all workbook memory.

**4. Test harness.** Replaced verquire's eager glob with direct imports, deleted
the Chai/dirty-Chai adapter and Mocha aliases, removed ambient compatibility
declarations, and disabled Vitest globals. Assertions and hooks now use native
Vitest imports. Transform helper tests use direct synchronous/async functions
instead of Promise-constructor wrappers, preserving model/XML comparisons.
A small pure XML normalizer retains the previous fixture comparison semantics.
Typed imports expose formerly hidden errors in mocks and helpers, which are fixed.
Test typechecking now runs in CI and the prepublish checks.

The identical 963-test / 98-file suite passes in every comparison run. Median
aggregate setup time falls from 8.62 s to 0.291 s; aggregate imports rise from
0.614 s to 2.56 s as imports move into the tests that need them. Reported setup
and import figures are sums across files, while the suite duration is wall time.
The suite had 706 type errors at the original starting point; earlier removals
reduced that to 654 before this fourth change, which removes the remainder.

Across all four changes, `lib/` falls from **182 files / 26,805 lines** to
**179 files / 25,366 lines**: a net reduction of **1,439 production lines**.
Four runtime utility files and three test compatibility files are deleted.

Validation: all 963 tests pass, library/spec/public typechecks pass, ESM and browser
smokes pass, asset budgets pass, and `vp check` passes. Regression coverage includes
random overlapping ranges, discontiguous references, saved XLSX projections and
stream parts, slow STORE sinks, interleaved callback sheets, destination errors,
premature closes, producer failures and Duplex completion.

**Method and reproduction.** Range workers run five timed trials with GC between
trials; the table takes the median of three fresh worker medians. Streaming and
test-suite comparisons each use three fresh processes, alternating version order.
Peak RSS is the process high-water mark and includes imports/runtime overhead.
The slow sink uses a 16 KiB high-water mark and a one-millisecond callback delay.
The browser comparison uses the same esbuild-minified entry and shims for each
stage. Small timing differences are local samples, not CI performance thresholds.

Saved builds/source controls are under ignored `build/internal-cleanup/`:
stage0 = starting point, stage1 = ranges, stage2 = wrapper removal, stage3 = native
streams with the old test harness, stage4 = native test harness. The final range
implementation is mirrored into both wrapper controls, and both test controls
use the same production implementation. Timed runs are performed sequentially.

```sh
node scripts/bench/internal-cleanup.mjs ranges stage0 stage1
node scripts/bench/internal-cleanup.mjs bundle stage0 stage1 stage2 stage3
node scripts/bench/internal-streams.mjs stage2 stage3
node scripts/bench/internal-tests.mjs
```

Raw samples are checked in as `internal-cleanup-results.json`. The commands use
local saved stages; a fresh clone needs equivalent staged build/source snapshots.
