import {EventEmitter} from 'node:events';
import {eachSaxChunk} from '../../utils/parse-sax.js';
import Enums from '../../model/enums.js';
import RelType from '../../xlsx/rel-type.js';

export interface HyperlinkReaderOptions {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbook: any;
  id: number | string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iterator: any;
  options: {
    hyperlinks?: 'cache' | 'emit' | 'ignore' | string;
    [key: string]: unknown;
  };
}

export interface HyperlinkRelationship {
  type: unknown;
  rId: string;
  target: string;
  targetMode: string;
}

class HyperlinkReader extends EventEmitter {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbook: any;
  id: number | string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iterator: any;
  options: HyperlinkReaderOptions['options'];
  hyperlinks?: Record<string, HyperlinkRelationship> | HyperlinkRelationship[];

  constructor({workbook, id, iterator, options}: HyperlinkReaderOptions) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options;
  }

  get count(): number {
    return (this.hyperlinks && (this.hyperlinks as HyperlinkRelationship[]).length) || 0;
  }

  each(
    fn: (value: HyperlinkRelationship, index: number, array: HyperlinkRelationship[]) => void,
  ): void {
    return (this.hyperlinks as HyperlinkRelationship[]).forEach(fn);
  }

  async read(): Promise<void> {
    const {iterator, options} = this;
    let emitHyperlinks = false;
    let hyperlinks: Record<string, HyperlinkRelationship> | null = null;
    switch (options.hyperlinks) {
      case 'emit':
        emitHyperlinks = true;
        break;
      case 'cache':
        this.hyperlinks = hyperlinks = {};
        break;
      default:
        break;
    }

    if (!emitHyperlinks && !hyperlinks) {
      this.emit('finished');
      return;
    }

    try {
      for await (const _chunk of eachSaxChunk(iterator, {
        onOpen: (name, attr) => {
          if (name === 'Relationship') {
            const rId = attr('Id') || '';
            switch (attr('Type')) {
              case RelType.Hyperlink:
                {
                  const relationship: HyperlinkRelationship = {
                    type: Enums.RelationshipType.Styles,
                    rId,
                    target: attr('Target') || '',
                    targetMode: attr('TargetMode') || '',
                  };
                  if (emitHyperlinks) {
                    this.emit('hyperlink', relationship);
                  } else {
                    hyperlinks![relationship.rId] = relationship;
                  }
                }
                break;

              default:
                break;
            }
          }
        },
        onText: () => {},
        onClose: () => {},
      })) {
        // Relationships are emitted as they are opened.
      }
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }
}

export default HyperlinkReader;
export {HyperlinkReader};
