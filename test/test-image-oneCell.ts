import {fileURLToPath} from 'node:url';
import {Workbook} from '../lib/exceljs.nodejs.ts';
import path from 'node:path';
import HrStopwatch from './utils/hr-stopwatch.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const filename = process.argv[2];

const wb = new Workbook();
const ws = wb.addWorksheet('blort');

const imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
ws.addImage(imageId, {
  tl: {col: 0.1125, row: 0.4},
  br: {col: 2.101046875, row: 3.4},
  editAs: 'oneCell',
});

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
    console.error(error.stack);
  });
