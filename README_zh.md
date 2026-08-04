# @sebbebroman/excel-ts

读写 Excel 工作簿（`.xlsx`）和 CSV 的 TypeScript / ESM 库。

基于 [ExcelJS](https://github.com/exceljs/exceljs) 的现代化分支，面向 Node 与 Vite / SvelteKit 等打包工具，支持按需入口与 tree-shaking。

## 安装

```bash
npm install @sebbebroman/excel-ts
```

需要 **Node.js ≥ 22**。

## 快速开始

```ts
import { Workbook } from '@sebbebroman/excel-ts';

const workbook = new Workbook();
const sheet = workbook.addWorksheet('Data');
sheet.addRow(['name', 'value']);
await workbook.xlsx.writeFile('out.xlsx');
```

可选入口：

- `@sebbebroman/excel-ts/csv` — CSV
- `@sebbebroman/excel-ts/stream/xlsx` — 流式读写

更完整的英文说明（浏览器、tree-shaking、与上游差异）见 [README.md](./README.md)。

## 许可与致谢

[MIT](./LICENSE)。源自 Guyon Roche 与 [ExcelJS](https://github.com/exceljs/exceljs) 贡献者的工作；本包在其上做了 ESM / TypeScript 等修改。
