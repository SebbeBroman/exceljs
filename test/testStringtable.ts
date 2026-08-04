import fs from 'node:fs';
import _ from '../lib/utils/under-dash.ts';
import SharedStrings from '../lib/utils/shared-strings.ts';

const filename = process.argv[2];

const st = new SharedStrings(null);

const lst = ['Hello', 'Hello', 'World', 'Hello\nWorld!', 'Hello, "World!"'];
_.each(lst, item => {
  st.add(item);
});

console.log(`Writing sharedstrings to ${filename}`);
const stream = fs.createWriteStream(filename);
st.write(stream).then(() => {
  stream.close();
  console.log('Done.');
});
