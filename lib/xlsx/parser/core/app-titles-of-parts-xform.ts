import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';

class AppTitlesOfPartsXform extends BaseXform<{name: string}[]> {
  override parseOpen(node: XmlNode): boolean {
    // no parsing
    return node.name === 'TitlesOfParts';
  }

  override parseText(): void {}

  override parseClose(name?: string): boolean {
    return name !== 'TitlesOfParts';
  }
}

export default AppTitlesOfPartsXform;
export {AppTitlesOfPartsXform};
