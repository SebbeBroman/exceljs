import {workbook, load} from '@sebbebroman/exceljs';
import * as main from '@sebbebroman/exceljs';
import {csv, viewCsv, readCsvRows, type CsvStringifyOptions} from '@sebbebroman/exceljs/csv';
const builder = workbook().sheet('S').cell('A1', 'hello').note('A1', 'note');
await load(await builder.writeBuffer());
const options: CsvStringifyOptions = {sheetName: 'S'};
await csv.stringify(builder, options);
await viewCsv('a,b\n1,2');
await readCsvRows(new TextEncoder().encode('a,b\n1,2'));
// @ts-expect-error CSV is available only through the optional entry.
void main.csv;
// @ts-expect-error CSV formatting no longer lives on the builder.
await builder.csv();
