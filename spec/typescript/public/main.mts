import {workbook, load, readRows, csv, type Workbook} from '@sebbebroman/exceljs';
import * as main from '@sebbebroman/exceljs';

const builder = workbook().sheet('Data').row(['name', 42]);
const bytes: Uint8Array = await builder.writeBuffer();
const model: Workbook = await load(bytes);
const rows: string[][] = await readRows(bytes);
await workbook(model).writeBuffer();
await csv.stringify(builder);
await csv.parse('name,value\nalpha,1');
void rows;

// These exports belong exclusively to the Node entry point.
// @ts-expect-error Node file helpers must not be declared on the main entry.
void main.writeFile;
// @ts-expect-error Node file helpers must not be declared on the main entry.
void main.readFile;
// @ts-expect-error Node CSV file helpers must not be declared on the main entry.
void main.readCsvFile;
// @ts-expect-error Node CSV file helpers must not be declared on the main entry.
void main.writeCsvFile;
// @ts-expect-error Node streaming helpers must not be declared on the main entry.
void main.streamWrite;
// @ts-expect-error Node streaming helpers must not be declared on the main entry.
void main.streamRead;
