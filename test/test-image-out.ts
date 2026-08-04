import {fileURLToPath} from 'node:url';
import {Workbook} from '../excel.ts';
import fs from 'node:fs';
import path from 'node:path';
import HrStopwatch from './utils/hr-stopwatch.ts';

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const filename = process.argv[2];

const wb = new Workbook();
const ws = wb.addWorksheet('blort');

ws.getCell('B2').value = 'Hello, World!';

const imageId = wb.addImage({
  filename: path.join(__dirname, 'data/image2.png'),
  extension: 'png',
});
const backgroundId = wb.addImage({
  buffer: fs.readFileSync(path.join(__dirname, 'data/bubbles.jpg')),
  extension: 'jpeg',
});
ws.addImage(imageId, {
  tl: {col: 1, row: 1},
  br: {col: 3.5, row: 5.5},
});
ws.addImage(imageId, 'B7:E12');

ws.addBackgroundImage(backgroundId);

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
