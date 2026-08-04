import fs from 'node:fs';
import SharedStrings from '../lib/utils/shared-strings.ts';

const filename = process.argv[2];

const st = new SharedStrings(null);

const lst = ['Hello', 'Hello', 'World', 'Hello\nWorld!', 'Hello, "World!"'];
for (const item of lst) {
  st.add(item);
}

console.log(`Writing sharedstrings to ${filename}`);
const stream = fs.createWriteStream(filename);
st.write(stream).then(() => {
  stream.close();
  console.log('Done.');
});
