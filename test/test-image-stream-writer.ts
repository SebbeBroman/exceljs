import {fileURLToPath} from 'node:url';
import path from 'node:path';
import Excel from '../lib/exceljs.nodejs.ts';
import HrStopwatch from './utils/hr-stopwatch.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const [, , filename] = process.argv;

const wb = new Excel.stream.xlsx.WorkbookWriter({filename});

const imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});

const ws = wb.addWorksheet('Foo');
ws.addBackgroundImage(imageId);

const stopwatch = new HrStopwatch();
stopwatch.start();

wb.commit()
  .then(() => {
    const micros = stopwatch.microseconds;
    console.log('Done.');
    console.log('Time taken:', micros);
  })
  .catch(error => {
    console.log(error.message);
  });
