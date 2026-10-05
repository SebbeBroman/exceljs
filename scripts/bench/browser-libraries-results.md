# Browser comparison — refreshed 2026-10-02

Current local fork with local fast-csv, SheetJS CE 0.20.3 ESM, upstream ExcelJS 4.4.0; Chrome 154; two warmups, median of seven alternating runs. Fixture: 2,000 × 8 plain cells (one string + seven numbers per row).

| Browser operation   | Local fork | SheetJS | ExcelJS |
| ------------------- | ---------: | ------: | ------: |
| Write (ms)          |       11.9 |    17.6 |    35.8 |
| Read same XLSX (ms) |       17.3 |    21.7 |    63.9 |
| Roundtrip (ms)      |       27.7 |    36.2 |    98.6 |

Read/write gzip bundles: fork 97.1 KiB; SheetJS ESM 157.6 KiB; ExcelJS 251.6 KiB. The fork’s values view is now 16.8 KiB gzip. SheetJS mini is 84.8 KiB with fewer supported formats/encodings.

All nine cross-library write/read combinations passed. Bigger-fixture results, detailed size attribution, CPU profiles, tested improvements, experiment limits and reproduction commands are in [browser-profile-results.md](browser-profile-results.md).
