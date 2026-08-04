import {EventEmitter} from '../../utils/event-emitter.js';
import parseSax from '../../utils/parse-sax.js';
import utils from '../../utils/utils.js';
import colCache from '../../utils/col-cache.js';
import Dimensions from '../../doc/range.js';
import Row from '../../doc/row.js';
import Column from '../../doc/column.js';

export interface WorksheetReaderOptions {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbook: any;
  id: number | string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iterator: any;
  options?: {
    worksheets?: string;
    hyperlinks?: string;
    [key: string]: unknown;
  };
}

// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyRow = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type AnyColumn = any;
// eslint-disable-next-line @typescript-eslint/no-explicit-any
type CellParseState = {
  ref: string;
  s: number;
  t: string;
  f?: {text: string};
  v?: {text: string};
};

class WorksheetReader extends EventEmitter {
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  workbook: any;
  id: number | string;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  iterator: any;
  options: NonNullable<WorksheetReaderOptions['options']>;
  name: string;
  state?: string;
  _columns: AnyColumn[] | null;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _keys: Record<string, any>;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  _dimensions: any;
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  hyperlinks?: Record<string, any>;

  constructor({workbook, id, iterator, options}: WorksheetReaderOptions) {
    super();

    this.workbook = workbook;
    this.id = id;
    this.iterator = iterator;
    this.options = options || {};

    // and a name
    this.name = `Sheet${this.id}`;

    // column definitions
    this._columns = null;
    this._keys = {};

    // keep a record of dimensions
    this._dimensions = new Dimensions();
  }

  // destroy - not a valid operation for a streaming writer
  // even though some streamers might be able to, it's a bad idea.
  destroy(): never {
    throw new Error('Invalid Operation: destroy');
  }

  // return the current dimensions of the writer
  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  get dimensions(): any {
    return this._dimensions;
  }

  // =========================================================================
  // Columns

  // get the current columns array.
  get columns(): AnyColumn[] | null {
    return this._columns;
  }

  // get a single column by col number. If it doesn't exist, it and any gaps before it
  // are created.
  getColumn(c: number | string): AnyColumn {
    if (typeof c === 'string') {
      // if it matches a key'd column, return that
      const col = this._keys[c];
      if (col) {
        return col;
      }

      // otherise, assume letter
      c = colCache.l2n(c);
    }
    if (!this._columns) {
      this._columns = [];
    }
    if (c > this._columns.length) {
      let n = this._columns.length + 1;
      while (n <= c) {
        // eslint-disable-next-line @typescript-eslint/no-explicit-any
        this._columns.push(new Column(this as any, n++));
      }
    }
    return this._columns[c - 1];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  getColumnKey(key: string): any {
    return this._keys[key];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  setColumnKey(key: string, value: any): void {
    this._keys[key] = value;
  }

  deleteColumnKey(key: string): void {
    delete this._keys[key];
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  eachColumnKey(f: (value: any, key: string) => void): void {
    for (const [key, value] of Object.entries(this._keys)) {
      f(value, key);
    }
  }

  async read(): Promise<void> {
    try {
      for await (const events of this.parse()) {
        for (const {eventType, value} of events) {
          this.emit(eventType, value);
        }
      }
      this.emit('finished');
    } catch (error) {
      this.emit('error', error);
    }
  }

  async *[Symbol.asyncIterator](): AsyncGenerator<AnyRow> {
    for await (const events of this.parse()) {
      for (const {eventType, value} of events) {
        if (eventType === 'row') {
          yield value;
        }
      }
    }
  }

  // eslint-disable-next-line @typescript-eslint/no-explicit-any
  async *parse(): AsyncGenerator<Array<{eventType: string; value: any}>> {
    const {iterator, options} = this;
    let emitSheet = false;
    let emitHyperlinks = false;
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let hyperlinks: Record<string, any> | null = null;
    switch (options.worksheets) {
      case 'emit':
        emitSheet = true;
        break;
      case 'prep':
        break;
      default:
        break;
    }
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
    if (!emitSheet && !emitHyperlinks && !hyperlinks) {
      return;
    }

    // references
    const {sharedStrings, styles, properties} = this.workbook;

    // xml position
    let inCols = false;
    let inRows = false;
    let inHyperlinks = false;

    // parse state
    // eslint-disable-next-line @typescript-eslint/no-explicit-any
    let cols: any[] | null = null;
    let row: AnyRow | null = null;
    let c: CellParseState | null = null;
    let current: {text: string} | null = null;
    for await (const events of parseSax(iterator)) {
      // eslint-disable-next-line @typescript-eslint/no-explicit-any
      const worksheetEvents: Array<{eventType: string; value: any}> = [];
      for (const {eventType, value} of events) {
        if (eventType === 'opentag') {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const node = value as any;
          if (emitSheet) {
            switch (node.name) {
              case 'cols':
                inCols = true;
                cols = [];
                break;
              case 'sheetData':
                inRows = true;
                break;

              case 'col':
                if (inCols) {
                  cols!.push({
                    min: parseInt(node.attributes.min, 10),
                    max: parseInt(node.attributes.max, 10),
                    width: parseFloat(node.attributes.width),
                    styleId: parseInt(node.attributes.style || '0', 10),
                  });
                }
                break;

              case 'row':
                if (inRows) {
                  const r = parseInt(node.attributes.r, 10);
                  // eslint-disable-next-line @typescript-eslint/no-explicit-any
                  row = new Row(this as any, r);
                  if (node.attributes.ht) {
                    row.height = parseFloat(node.attributes.ht);
                  }
                  if (node.attributes.s) {
                    const styleId = parseInt(node.attributes.s, 10);
                    const style = styles.getStyleModel(styleId);
                    if (style) {
                      row.style = style;
                    }
                  }
                }
                break;
              case 'c':
                if (row) {
                  c = {
                    ref: node.attributes.r,
                    s: parseInt(node.attributes.s, 10),
                    t: node.attributes.t,
                  };
                }
                break;
              case 'f':
                if (c) {
                  current = c.f = {text: ''};
                }
                break;
              case 'v':
                if (c) {
                  current = c.v = {text: ''};
                }
                break;
              case 'is':
              case 't':
                if (c) {
                  current = c.v = {text: ''};
                }
                break;
              case 'mergeCell':
                break;
              default:
                break;
            }
          }

          // =================================================================
          //
          if (emitHyperlinks || hyperlinks) {
            switch (node.name) {
              case 'hyperlinks':
                inHyperlinks = true;
                break;
              case 'hyperlink':
                if (inHyperlinks) {
                  const hyperlink = {
                    ref: node.attributes.ref,
                    rId: node.attributes['r:id'],
                  };
                  if (emitHyperlinks) {
                    worksheetEvents.push({eventType: 'hyperlink', value: hyperlink});
                  } else {
                    hyperlinks![hyperlink.ref] = hyperlink;
                  }
                }
                break;
              default:
                break;
            }
          }
        } else if (eventType === 'text') {
          // only text data is for sheet values
          if (emitSheet) {
            if (current) {
              current.text += value as string;
            }
          }
        } else if (eventType === 'closetag') {
          // eslint-disable-next-line @typescript-eslint/no-explicit-any
          const node = value as any;
          if (emitSheet) {
            switch (node.name) {
              case 'cols':
                inCols = false;
                // eslint-disable-next-line @typescript-eslint/no-explicit-any
                this._columns = Column.fromModel(cols as any);
                break;
              case 'sheetData':
                inRows = false;
                break;

              case 'row':
                this._dimensions.expandRow(row);
                worksheetEvents.push({eventType: 'row', value: row});
                row = null;
                break;

              case 'c':
                if (row && c) {
                  const address = colCache.decodeAddress(c.ref);
                  const cell = row.getCell(address.col);
                  if (c.s) {
                    const style = styles.getStyleModel(c.s);
                    if (style) {
                      cell.style = style;
                    }
                  }

                  if (c.f) {
                    // eslint-disable-next-line @typescript-eslint/no-explicit-any
                    const cellValue: any = {
                      formula: c.f.text,
                    };
                    if (c.v) {
                      if (c.t === 'str') {
                        cellValue.result = utils.xmlDecode(c.v.text);
                      } else {
                        cellValue.result = parseFloat(c.v.text);
                      }
                    }
                    cell.value = cellValue;
                  } else if (c.v) {
                    switch (c.t) {
                      case 's': {
                        const index = parseInt(c.v.text, 10);
                        if (sharedStrings) {
                          cell.value = sharedStrings[index];
                        } else {
                          cell.value = {
                            sharedString: index,
                          };
                        }
                        break;
                      }

                      case 'inlineStr':
                      case 'str':
                        cell.value = utils.xmlDecode(c.v.text);
                        break;

                      case 'e':
                        cell.value = {error: c.v.text};
                        break;

                      case 'b':
                        cell.value = parseInt(c.v.text, 10) !== 0;
                        break;

                      default:
                        if (utils.isDateFmt(cell.numFmt)) {
                          cell.value = utils.excelToDate(
                            parseFloat(c.v.text),
                            properties.model && properties.model.date1904,
                          );
                        } else {
                          cell.value = parseFloat(c.v.text);
                        }
                        break;
                    }
                  }
                  if (hyperlinks) {
                    const hyperlink = hyperlinks[c.ref];
                    if (hyperlink) {
                      cell.text = cell.value;
                      cell.value = undefined;
                      cell.hyperlink = hyperlink;
                    }
                  }
                  c = null;
                }
                break;
              default:
                break;
            }
          }
          if (emitHyperlinks || hyperlinks) {
            switch (node.name) {
              case 'hyperlinks':
                inHyperlinks = false;
                break;
              default:
                break;
            }
          }
        }
      }
      if (worksheetEvents.length > 0) {
        yield worksheetEvents;
      }
    }
  }
}

export default WorksheetReader;
export {WorksheetReader};
