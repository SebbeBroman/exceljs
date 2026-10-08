# Legacy document removal — 2026-10-08

The mutable document implementation is removed from both buffered and streaming
paths. There is no `lib/doc`, internal ExcelJS namespace entry, CSV attachment,
synchronous feature loader, old document compiler/projector, or legacy `index.d.ts`.
Public main and Node entry points are preserved.

## Code removed

Compared with the original Git HEAD, production TypeScript plus internal
`index.d.ts` shrank from **33,736 to 26,805 lines: 6,931 fewer lines (20.5%)**.
This includes the preceding buffered rewrite. It excludes retired tests, docs,
benchmarks and fixture data.

Reusable ranges, defined names, notes, image anchors and enums now live in
`lib/model`. Structural types are separated from the removed class API. Table
placement is a function writing model cells directly. Node stream writers compile
one row at a time; stream readers yield sparse value arrays. Internal WorkbookWriter
and WorkbookReader coordinate ZIP parts and streams, without a document graph.
Unused style-copy, HTML-escape, formula-translation and streaming-comment modules
were also removed. CSV parsing/stringifying remains supported through public helpers.

## Streaming before/after

Baseline: saved build immediately after the buffered bridge rewrite, while Node
streaming still used legacy rows/cells/columns. Current: direct row/value models.
Node v24.21.0, eight columns, unique inline strings, shared strings and styles off.
Async producer yields to the event loop every 1,000 rows. No input grid or collected
read rows are retained. Three isolated process samples per version/workload/size,
alternating order. Values and row counts are checked during reads.

| Workload       | Before median | After median |       Change | Before peak RSS | After peak RSS |
| -------------- | ------------: | -----------: | -----------: | --------------: | -------------: |
| Write 50k × 8  |      317.6 ms |     286.4 ms |  9.8% faster |       101.9 MiB |      101.4 MiB |
| Read 50k × 8   |      521.4 ms |     431.9 ms | 17.2% faster |       154.9 MiB |      114.6 MiB |
| Write 100k × 8 |      591.3 ms |     532.2 ms | 10.0% faster |       103.2 MiB |      102.4 MiB |
| Read 100k × 8  |    1,006.3 ms |     809.3 ms | 19.6% faster |       215.2 MiB |      150.7 MiB |

Read peak RSS is 26.0% lower at 50k and 29.9% lower at 100k. Writer memory is
roughly unchanged and flat over these sizes. Reader RSS still grows with input
size: removing row/cell objects does not eliminate ZIP/parser buffering or caches.
Shared strings/styles/hyperlinks have their own retained state when enabled.
These small local samples indicate direction, not universal performance guarantees.

Current browser write-only bundle: **290.0 KiB minified, 84.8 KiB gzip**;
code-split entry 162.2 KiB, total 287.9 KiB in 31 files. Immediately after the
buffered rewrite it was 297.3 KiB / 86.7 KiB gzip. See
[the earlier buffered comparison](doc-bridge-results.md) for that phase's timing
and memory measurements.

## Validation and coverage

- **963 tests pass across 98 files**.
- Library and public declaration checks, build, touched-source lint/format,
  Git whitespace checks, native ESM/Node file/CSV/stream smoke, browser smoke and
  asset-size gates pass.
- **33 existing XLSX fixtures match the saved buffered projections**, including
  identical rejection of invalid input. The private image-anchor worksheet host
  is excluded from normalized snapshot comparisons.
- **The same 33 fixtures match saved streaming row digests/counts/rejections**.
- Four shared-string/style combinations match every uncompressed streaming XML
  part and their default read rows. Twenty explicit-option combinations match
  cached styles, shared strings ignored/emitted, hyperlinks cached and worksheets
  ignored. Cached dates, sparse rows, rich text, errors, formulas, hyperlinks,
  column settings, sheet order and printing options are covered.
- Buffered golden tests retain exact XML/media parity for sparse values/edits,
  dates, formulas, styles, merges, comments, validation, conditional formatting,
  protection, tables/totals, images and metadata. Twelve deterministic randomized
  builder sequences are checked. Plain snapshot immutability and row metadata
  remain covered.
- Useful range/name/anchor tests moved to `spec/unit/model`. Sixty test files for
  the retired mutable API or now-unused utilities were removed, rather than
  continuing to test an implementation that is no longer shipped. Encoder and
  public API tests remain; golden fixtures preserve relevant feature coverage.
- The streaming fixture sweep caught the historical omission of false/zero/NaN
  formula results. The direct reader preserves that existing stream projection.
- No production or test imports of the removed document modules remain. A clean
  build removes stale output; `dist/lib/doc` is absent.

The separate optional spec typecheck still reports **706 diagnostics** in the
historical test infrastructure/specs (previous baseline: 3,152). It is not a passing
check. The new builder compatibility files have no diagnostics; production and
published declarations pass their strict checks.

`huge.xlsx` is omitted from automated fixture digests to keep test memory bounded;
it matched during the earlier buffered sweep. `test-issue-1842.xlsx` is excluded:
the baseline expands its whole-grid validation and stalls before projection. This
change does not rewrite validation parsing. Complex pivot/image round-tripping
retains the existing public API limitations.

## Reproduce

Before changing the streaming implementation, save its build:

```sh
pnpm build
mkdir -p build/legacy-removal-before
cp -R dist build/legacy-removal-before/dist
```

After the change:

```sh
pnpm build
node scripts/bench/legacy-removal.mjs
pnpm test
pnpm typecheck
pnpm typecheck:public
pnpm test:esm
pnpm test:browser-bundle
pnpm check:asset-size
```

Raw streaming samples are in ignored
`build/legacy-removal-before/stream-results.json`. Committed-source golden fixtures
under `spec/unit/builder/data` do not require the saved build. This report records the local rewrite and its measurements.
