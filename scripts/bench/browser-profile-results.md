# Updated browser measurements and optimization findings

## Follow-up: browser reader split and remaining dependencies

The five runtime dependencies are fflate (ZIP), local fast-csv (CSV), dayjs with date plugins (CSV dates), @noble/hashes (worksheet protection), and saxen (XML parsing). Their retained raw minified contributions in the write graph are approximately 15.2, 14.7, 13.0, 7.2, and 6.7 KiB respectively; gzip contributions cannot be added independently.

The fused worksheet reader now loads through a cached dynamic import in XLSX read methods. This preserves the read/write API and removes its direct SAX dependency from the initial export graph. The new browser-startup.mjs script builds static-reader and lazy-reader variants from the same current sources, counts all static dependencies, and records the chunks actually requested by the first plain export in Chrome. Startup gzip falls from 84.1 to 78.9 KiB; including first-export dynamic downloads, 85.7 to 80.5 KiB. Total per-file gzip remains 103.6 KiB. A single-file build is 96.1 KiB gzip, so this is a startup improvement for code-splitting bundlers rather than a reduction in total implementation size.

Build, browser smoke, 45 focused fused-reader/workbook tests, and all nine 2,000-row cross-library writer/reader checks passed. The follow-up browser rerun measured local write/read-same-file at 11.0/18.6 ms, versus SheetJS 16.2/22.6 and ExcelJS 33.8/63.5; these short-run fluctuations are not evidence of a throughput gain from the split. The earlier measurements below describe the preceding snapshot.

Full styles remain coupled: even useStyles:false constructs StylesXform.Mock, which inherits the complete style transform tree. A separate minimal style manager with date-format parity is a further opportunity. The document-model opportunity remains writing directly from the plain builder model and reading directly into the plain snapshot; the restricted dense prototype below demonstrates the potential but is not feature equivalent.

Measured locally on 2026-10-02: Chrome 154, Node v24.21.0, esbuild 0.28.1. Competitors are SheetJS CE 0.20.3 ESM and upstream ExcelJS 4.4.0 browser distribution. Both forks were rebuilt first. No commits, pushes, publishing, or project dependency additions were made.

The updated fork still leads both competitors for the tested plain-cell workload. The safe changes made during this investigation remove a duplicate worksheet serialization and cut the values-view bundle by 1.2 KiB gzip. The strongest larger opportunity is an explicit dense export entry that bypasses the full document model and XML transform classes. A limited prototype demonstrates its potential; it is not a replacement for the existing full API.

| Browser bundle (gzip KiB)     | Previous measurement | Updated forks before this work | After local improvements |
| ----------------------------- | -------------------: | -----------------------------: | -----------------------: |
| Fork write-only               |                95.97 |                          95.76 |                    95.79 |
| Fork read + write             |                97.24 |                          97.03 |                    97.06 |
| Fork values view (CSV + XLSX) |                18.23 |                          17.96 |                    16.78 |
| Fork CSV helpers              |                11.66 |                          11.44 |                    11.44 |

Current SheetJS write-only ESM is 107.1 KiB gzip, read/write ESM 157.6 KiB, and upstream ExcelJS is 251.6 KiB (232.2 KiB for bare). SheetJS mini is 84.8 KiB, with a narrower format/encoding feature set. The fork’s full write-only graph retains real optional CSV and feature implementations; nothing is stubbed. Named imports and browser-specific entries are used.

The table below is the final uninstrumented browser rerun. Each workload has two warmups and seven measured runs, rotating library order. Each row has eight cells: one string and seven numbers. Write timings include builder/workbook creation. All writers use compressed XLSX and shared strings. Shared-read timing uses exactly the same XLSX bytes for every library; SheetJS uses dense reads.

| Rows × 8; operation (ms)    | Local fork | SheetJS | ExcelJS |
| --------------------------- | ---------: | ------: | ------: |
| 2,000; write                |       11.9 |    17.6 |    35.8 |
| 2,000; read same file       |       17.3 |    21.7 |    63.9 |
| 2,000; write/read own file  |       27.7 |    36.2 |    98.6 |
| 2,000; extract row values   |        9.6 |    22.8 |    63.6 |
| 20,000; write               |      111.2 |   196.2 |   307.9 |
| 20,000; read same file      |      155.1 |   208.4 |   259.7 |
| 20,000; write/read own file |      275.7 |   386.1 |   576.6 |
| 20,000; extract row values  |      104.2 |   210.5 |   266.6 |

The safe serialization change cuts that phase, but changes in whole-workbook medians are small relative to run-to-run variation at 2,000 rows. Do not interpret the slightly different read results as a regression: neither change alters the full read path. On 20,000 rows the current fork writes about 1.76× as fast as SheetJS / 2.77× as fast as ExcelJS, and full reads are about 1.34× / 1.67× as fast. These ratios apply to this plain-data fixture, not every workbook feature.

CPU profiling used Chrome DevTools Protocol at 100 microseconds per sample, repeated workloads, and source-map attribution. Raw cpuprofile files are in build/browser-profile and can be imported into Chrome DevTools. Sampling and instrumented phase clocks affect absolute speed; the ordinary browser reruns above are the performance comparison. Self-time samples are exclusive, not cumulative.

| Hot path, 20,000 rows | CPU self-sample evidence before optimization                                                |
| --------------------- | ------------------------------------------------------------------------------------------- |
| Write                 | fflate: 59.7%; garbage collector: 7.3%                                                      |
| Full read             | saxen: 29.5%; fast-sheet-data: 16.3%; fflate: 14.7%; doc/row: 9.5%; garbage collector: 7.2% |
| Values view           | saxen: 43.9%; xlsx-light: 20.0%; fflate: 22.9%                                              |

Write time is mostly deflate/ZIP, followed by XML rendering, UTF8 encoding, and document construction. Full reads spend about 46% of sampled CPU in SAX plus fused cell parsing, then inflate and hydration. The values view already avoids document hydration but still parses XML through SAX. CSV-builder export separately spends about 23% of sampled CPU compiling the builder into a plain model, 20% in mapping/extracting row arrays, and 23% formatting fields; a direct dense-row route could avoid the intermediate model. The refreshed 2,000 × 8 CSV run measured 0.7 ms for identity-mapped parse and 2.1 ms for builder stringify, versus 6.9 / 7.1 ms for registry fast-csv with a microtask scheduler. CSV parsing is already fast and has less remaining absolute time to save.

Phase timing evidence (medians; some read phases are nested inside total load, so they must not be added to it):

| Phase                                | 2,000 rows before (ms) | 20,000 before (ms) | 20,000 after (ms) |
| ------------------------------------ | ---------------------: | -----------------: | ----------------: |
| write: materialize document          |                    1.1 |                7.2 |               8.5 |
| write: serialize document model      |                    1.1 |                9.3 |               4.7 |
| write: prepare styles/strings/models |                    0.4 |                2.6 |               2.5 |
| write: render XML + encode UTF8      |                    3.3 |               36.4 |              33.1 |
| write: deflate + ZIP                 |                    7.5 |               68.9 |              69.0 |
| read: inflate ZIP                    |                    2.5 |               24.0 |              24.5 |
| read: worksheet XML + fused cells    |                    8.0 |               78.9 |              76.4 |
| read: apply document model           |                    3.7 |               28.4 |              29.0 |
| read: document to plain snapshot     |                    1.1 |                8.4 |               7.7 |
| view: open ZIP + workbook/SST        |                    2.9 |               27.6 |              27.9 |
| view: parse/extract requested rows   |                    7.3 |               68.8 |              68.4 |

The write model originally serialized each worksheet once into worksheets and again into sheets. The new internal getXlsxModel method serializes each once; the public legacy model getter still returns independent snapshots for compatibility. Both buffered and streaming XLSX writes use the internal method, with fallback for alternate workbook hosts.

Bundle attribution uses esbuild bytesInOutput, counting only retained minified bytes. Module-level gzip is not additive, so the following proportions are raw minified contributions, not gzip estimates.

| Write-only bundle component | Retained KiB | Share |
| --------------------------- | -----------: | ----: |
| XML transforms              |        137.0 | 41.2% |
| document model              |         57.2 | 17.2% |
| dependencies                |         42.1 | 12.7% |
| utilities                   |         28.7 |  8.6% |
| XLSX orchestration          |         20.7 |  6.2% |
| fast-csv                    |         14.7 |  4.4% |
| model bridges               |         11.3 |  3.4% |
| builder                     |          9.3 |  2.8% |
| XML templates               |          8.0 |  2.4% |
| other                       |          3.3 |  1.0% |

The write-only bundle is only about 4 KiB smaller in raw bytes than read/write: shared DocWorkbook/XLSX classes retain their read methods, parsers, style transforms, and optional feature loaders. That architecture limits ordinary tree shaking. The values view is much smaller because it has a separate implementation. The utility change moves XML entity decoding into its own small module, avoiding retention of the default utility object and fs shim in the values-view bundle.

Code splitting helps startup, but does not solve the architecture: the real initial write graph (entry plus all static chunks) is 83.6 KiB gzip / 267.6 KiB raw. All chunks together are 102.9 KiB gzip across 29 files. Counting only the entry file would understate initial download size.

Three measured experiments inform next steps:

1. A separate dense-export prototype ships in 5.47 KiB gzip (11.5 KiB raw). It constructs OOXML directly from rows and reuses fflate. It is limited to one sheet with strings, finite numbers, booleans, and nulls: no styles, dates, formulas, metadata, images, tables or protection.

| Dense export experiment | Full writer (ms) | Prototype (ms) |
| ----------------------- | ---------------: | -------------: |
| 2,000 × 8               |             11.8 |            8.8 |
| 20,000 × 8              |            114.0 |           83.7 |

This is roughly 25–27% faster, with about 94% less gzipped JS for this restricted export capability. All three libraries read both prototype fixture sizes correctly. Special Unicode/XML-escaped strings, whitespace, booleans and nulls were also checked with SheetJS. The 20,000-row prototype output is about 3% larger than the full writer’s output, so the smaller/faster code does not imply a smaller XLSX file. Keep this as a dedicated entry/API if productized, rather than silently dropping workbook features.

2. Native CompressionStream(deflate-raw) was tested by replacing only the benchmark bundle’s ZIP implementation. Every uncompressed XML part was identical, and all three readers reproduced the cells. The native stream uses its platform default compression level; the existing fflate path uses level 1. Native output was 14–16% smaller, but slower:

| Rows × 8 | fflate write (ms) | Native write (ms) |
| -------- | ----------------: | ----------------: |
| 2,000    |              11.1 |              15.5 |
| 20,000   |             114.8 |             136.4 |

This does not support changing the default compressor for throughput. Native compression could still be an opt-in file-size/UI-responsiveness tradeoff; broader browser coverage would be needed.

3. fflate hash-memory settings 4/6/8 were tested at the same compression level. Output parts were identical after decompression; runtime differences were small and inconsistent. There is no demonstrated speed reason to change the default.

Prioritized next work:

- Separate dense export from the full document graph: the largest demonstrated bundle gain, with measurable throughput improvement. Expand only explicitly supported features and add cross-library regression coverage before exposing it.
- Separate full read and write orchestration/classes so write-only clients can drop SAX, read reconciliations, and unrelated transforms. Existing dynamic imports alone cannot remove retained class methods. This requires architectural work, not a manifest-only switch.
- Let a full read convert the fused worksheet model directly to the public plain snapshot when semantics allow, instead of hydrating compact DocWorkbook cells and converting them back. The measured hydration + plain-conversion phases take roughly 36–37 ms at 20,000 rows; that is an upper bound on eliminable bridge work, not a guaranteed speedup. Preserve styles, formula metadata, comments, hyperlinks, merges and feature fallback behavior.
- Explore a direct dense-builder CSV export route to avoid the builder → plain workbook → row arrays conversion. Lazy date-library loading may also shrink CSV-only bundles, but its asynchronous first use and date-format semantics need care.

Validation: both forks rebuilt; 217 relevant builder/document/utility/workbook tests passed, one skipped; native ESM/file/stream smoke and browser bundle smoke passed. All nine writer/reader combinations passed at 2,000 and 20,000 rows. An expanded integration run found an existing WorkbookReader hyperlinked-formula-source failure. Repeating that test with all four modified source modules restored to unchanged HEAD reproduced the same failure; baseline-test-failure.log records it. The reader regression was not changed as part of this profiling work.

Reproduce from the ExcelJS checkout (build the sibling fast-csv packages first):

```sh
pnpm build
node scripts/bench/browser-libraries.mjs
node scripts/bench/browser-libraries.mjs --rows 20000
node scripts/bench/browser-profile.mjs --label optimized
node scripts/bench/browser-native-zip.mjs
node scripts/bench/browser-dense-export.mjs
```

Generated browser bundles, source maps, CPU profiles, samples, and experiment outputs are kept in ignored build directories. The experimental native/dense routes do not alter the production API or compressor. The source changes, scripts, and this report remain local and uncommitted.
