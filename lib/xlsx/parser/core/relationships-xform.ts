import BaseXform from '../../base-parser.js';
import type {XmlNode} from '../../base-parser.js';
import RelationshipXform from './relationship-xform.js';
import type {RelationshipModel} from './relationship-xform.js';

class RelationshipsXform extends BaseXform<RelationshipModel[]> {
  declare map: {Relationship: RelationshipXform};
  _values?: RelationshipModel[];

  constructor() {
    super();

    this.map = {
      Relationship: new RelationshipXform(),
    };
  }

  override parseOpen(node: XmlNode): boolean {
    if (this.parser) {
      this.parser.parseOpen(node);
      return true;
    }
    switch (node.name) {
      case 'Relationships':
        this.model = [];
        return true;
      default:
        this.parser = this.map[node.name as keyof typeof this.map];
        if (this.parser) {
          this.parser.parseOpen(node);
          return true;
        }
        throw new Error(`Unexpected xml node in parseOpen: ${JSON.stringify(node)}`);
    }
  }

  override parseText(text: string): void {
    if (this.parser) {
      this.parser.parseText(text);
    }
  }

  override parseClose(name?: string): boolean {
    if (this.parser) {
      if (!this.parser.parseClose(name)) {
        this.model!.push(this.parser.model as RelationshipModel);
        this.parser = undefined;
      }
      return true;
    }
    switch (name) {
      case 'Relationships':
        return false;
      default:
        throw new Error(`Unexpected xml node in parseClose: ${name}`);
    }
  }
}

export default RelationshipsXform;
export {RelationshipsXform};
