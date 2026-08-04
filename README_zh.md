# @sebbebroman/excel-ts

读写 Excel 工作簿（`.xlsx`）的 **builder-first** TypeScript / ESM 库。

基于 [ExcelJS](https://github.com/exceljs/exceljs) 的现代化分支，面向 Node 与 Vite / SvelteKit 等打包工具。

> **5.0.0-alpha.1** — 仅两个包入口（`.` 与 `./node`）。完整英文说明见 [README.md](./README.md)；迁移见 [MIGRATION-5.0.md](./MIGRATION-5.0.md)；架构见 [ARCHITECTURE.md](./ARCHITECTURE.md)。

## 安装

```bash
npm install @sebbebroman/excel-ts
# 或
pnpm add @sebbebroman/excel-ts
```

需要 **Node.js ≥ 22**。

## 快速开始

```ts
import { workbook } from '@sebbebroman/excel-ts';

const buffer = await workbook({ creator: 'Reports' })
  .sheet('Data')
  .row(['name', 'value'])
  .row(['alpha', 1])
  .style('A1:B1', { font: { bold: true } })
  .writeBuffer();
```

### Node 写文件

```ts
import { workbook, writeFile } from '@sebbebroman/excel-ts/node';

await writeFile(
  'out.xlsx',
  workbook().sheet('Data').rows([
    ['name', 'value'],
    ['alpha', 1],
  ]),
);
```

### 加载与编辑

```ts
import { workbook, load } from '@sebbebroman/excel-ts';

const data = await load(buffer); // 普通 { meta, sheets }
const out = await workbook(data)
  .sheet('Sheet1')
  .cell('A1', 'updated')
  .writeBuffer();
```

### CSV

```ts
import { workbook, csv } from '@sebbebroman/excel-ts';

// 命名 API，无需 side-effect import（已无 `./csv` 包入口）
const init = await csv.parse('name,value\nalpha,1');
const buffer = await workbook().sheet('Data', init).writeBuffer();
const text = await workbook().sheet('Data', init).csv();
```

### Node 流式读写

```ts
import { streamWrite, streamRead } from '@sebbebroman/excel-ts/node';

await streamWrite('big.xlsx', {
  sheets: [{ name: 'Data', rows: largeAsyncIterable }],
});

for await (const { sheetName, rowNumber, values } of streamRead('big.xlsx')) {
  // values[1] 为 A 列
}
```

## 入口

| 导入 | 用途 |
|------|------|
| `@sebbebroman/excel-ts` | builder、`writeBuffer`、`load`、`csv`、枚举（浏览器可用） |
| `@sebbebroman/excel-ts/node` | 另含 `writeFile` / `readFile` / `streamWrite` / `streamRead` / `readCsvFile` / `writeCsvFile` |

无其他包导出（无 `./csv`、无 `./stream/xlsx`、无默认 `ExcelJS` 类）。

## 许可与致谢

[MIT](./LICENSE)。源自 Guyon Roche 与 [ExcelJS](https://github.com/exceljs/exceljs) 贡献者的工作；本包在其上做了 ESM / TypeScript / builder API 等修改。
