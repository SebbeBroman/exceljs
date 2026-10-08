These are saved compatibility outputs from the implementation before legacy removal.

- `direct-model/*.xlsx`: buffered output produced by the old document compiler,
  with fixed workbook metadata. Filenames identify the original builder operations
  and string mode. Tests compare every uncompressed XML/media part.
- `direct-model/projections.json`: SHA-256 of normalized old document projections,
  keyed by a digest of the XLSX parts. Dates serialize as ISO strings; the private
  image-anchor `worksheet` backpointer is excluded.
- `legacy-projections.json`: buffered projection digests/rejections for the existing
  integration XLSX files.
- `legacy-stream-projections.json`: incremental row digests/counts/rejections from
  the old public stream reader for those same files.
- `stream-*.xlsx`: output of `stream-fixture.mjs` through the old public stream
  writer, covering shared/inline strings with/without styles.
- `stream-*.json`: old stream reader outputs with default and explicit options.

The tests read these references; they do not regenerate expectations from the
implementation under test. Intentionally changing behavior requires reviewing
and replacing the corresponding expectations. Useful feature coverage replaces
retired tests for the unsupported mutable Workbook/Worksheet/Row/Cell API.

`huge.xlsx` is omitted from automated digest checks to keep test memory bounded.
`test-issue-1842.xlsx` is omitted because the baseline parser expands a validation
covering every Excel cell and stalls before projection. All 33 other XLSX fixtures
are checked in both buffered and streaming modes.
