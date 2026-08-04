import Excel from '../lib/exceljs.nodejs.ts';
import HrStopwatch from './utils/hr-stopwatch.ts';

const [, , filename] = process.argv;

const wb = new Excel.Workbook();
const ws = wb.addWorksheet('Foo');
ws.getCell('A1').value = 'Hello, World!';

const stopwatch = new HrStopwatch();
stopwatch.start();
wb.xlsx
  .writeFile(filename)
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.log(error.message);
  });
