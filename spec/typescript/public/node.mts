import {
  workbook,
  writeFile,
  readFile,
  readCsvFile,
  writeCsvFile,
  streamWrite,
  streamRead,
  type Workbook,
  type StreamWriteSpec,
  type StreamReadRow,
} from '@sebbebroman/exceljs/node';
import {Readable, Writable} from 'node:stream';

const builder = workbook().sheet('Data').row(['name', 42]);
await writeFile('out.xlsx', builder);
const model: Workbook = await readFile('out.xlsx');
await writeCsvFile('out.csv', model, {encoding: 'utf8'});
await readCsvFile('out.csv', {sheetName: 'Data'});

const spec: StreamWriteSpec = {sheets: [{name: 'Data', rows: [['name', 42]]}]};
await streamWrite('out.xlsx', spec);
await streamWrite(new Writable(), async book => {
  await book.sheet('Data').rows([['name', 42]]);
});
for await (const row of streamRead(new Readable())) {
  const typed: StreamReadRow = row;
  const name: string = typed.sheetName;
  void name;
}
