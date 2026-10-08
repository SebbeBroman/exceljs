# Opt-in protection bundle comparison

Compared commit `5f3c177c` (saved compiled `dist`) with the opt-in protection entry rewrite, using the same esbuild browser single-file ESM workload and minification settings.

| Browser workload | Before minified | After minified |             Saved | Before gzip | After gzip |             Saved |
| ---------------- | --------------: | -------------: | ----------------: | ----------: | ---------: | ----------------: |
| Basic XLSX write |       282,881 B |      274,374 B | 8,507 B (8.3 KiB) |    82,529 B |   78,587 B | 3,942 B (3.8 KiB) |
| XLSX row read    |        45,183 B |       45,183 B |                 0 |    17,226 B |   17,227 B |              -1 B |
| Full XLSX load   |       219,443 B |      219,443 B |                 0 |    60,699 B |   60,713 B |             -14 B |

Read-only paths already excluded hashing. Tiny gzip differences with identical minified sizes reflect symbol renaming/compression, not additional functionality.

Core writing previously retained `Encryptor` and noble SHA-512. It now takes prepared `sheetProtection` data and retains neither. Password hashing is available only through `@sebbebroman/exceljs/protection`; calling the factory hashes synchronously, then `.protect(model)` stores the prepared result. The old password overload and deferred plain-model password configuration are removed.

`node scripts/smoke-browser-bundle.mjs` asserts that core single-file and split bundles exclude crypto, an unused optional import disappears, and a used optional factory produces a hash that survives write/load. Native ESM and isolated public declarations also check the optional package export. Loaded protected worksheets can be edited and rewritten without the crypto entry.

Reproduce the comparison with `pnpm build`, then `node scripts/bench/protection-bundle.mjs /path/to/baseline/dist`. Raw measurements: [JSON](protection-bundle-results.json).
