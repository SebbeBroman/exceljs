import BaseXform from '../base-xform.js';
import type {XmlStreamLike, XmlNode} from '../base-xform.js';

class AppTitlesOfPartsXform extends BaseXform<{name: string}[]> {
  override render(xmlStream: XmlStreamLike, model?: {name: string}[] | null): void {
    xmlStream.openNode('TitlesOfParts');
    xmlStream.openNode('vt:vector', {size: model!.length, baseType: 'lpstr'});

    model!.forEach(sheet => {
      xmlStream.leafNode('vt:lpstr', undefined, sheet.name);
    });

    xmlStream.closeNode();
    xmlStream.closeNode();
  }

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
