// @vitest-migrated
import __req_0 from './test-data-validation-sheet.js';
import __req_1 from './test-conditional-formatting-sheet.js';
import __req_2 from './test-values-sheet.js';
import __req_3 from './test-spliced-sheet.js';
import tools from './tools.js';
import testWorkbookReader from './test-workbook-reader.js';
import __json_0 from './data/views.json';
import __json_1 from './data/sheet-values.json';
import __json_2 from './data/styles.json';
import __json_3 from './data/sheet-properties.json';
import __json_4 from './data/page-setup.json';
import __json_5 from './data/header-footer.json';
import __json_6 from './data/conditional-formatting.json';

const Row = verquire('doc/row');
const Column = verquire('doc/column');

const testSheets = {
  dataValidations: __req_0,
  conditionalFormatting: __req_1,
  values: __req_2,
  splice: __req_3,
};

/** Resolve a dotted path on an object (e.g. `splice.rows.removeOnly`). */
function getPath(obj, path) {
  const parts = typeof path === 'string' ? path.split('.') : path;
  let cur = obj;
  for (const part of parts) {
    if (cur == null) return undefined;
    cur = cur[part];
  }
  return cur;
}

function getOptions(docType, options) {
  let result;
  switch (docType) {
    case 'xlsx':
      result = {
        sheetName: 'values',
        checkFormulas: true,
        checkMerges: true,
        checkStyles: true,
        checkBadAlignments: true,
        checkSheetProperties: true,
        dateAccuracy: 3,
        checkViews: true,
      };
      break;
    case 'csv':
      result = {
        sheetName: 'sheet1',
        checkFormulas: false,
        checkMerges: false,
        checkStyles: false,
        checkBadAlignments: false,
        checkSheetProperties: false,
        dateAccuracy: 1000,
        checkViews: false,
      };
      break;
    default:
      throw new Error(`Bad doc-type: ${docType}`);
  }
  return Object.assign(result, options);
}

function createSheetMock() {
  return {
    _keys: {},
    _cells: {},
    rows: [],
    columns: [],
    properties: {
      outlineLevelCol: 0,
      outlineLevelRow: 0,
    },

    addColumn(colNumber, defn) {
      const newColumn = new Column(this, colNumber, defn);
      this.columns[colNumber - 1] = newColumn;
      return newColumn;
    },
    getColumn(colNumber) {
      let column = this.columns[colNumber - 1] || this._keys[colNumber];
      if (!column) {
        column = this.columns[colNumber - 1] = new Column(this, colNumber);
      }
      return column;
    },
    getRow(rowNumber) {
      let row = this.rows[rowNumber - 1];
      if (!row) {
        row = this.rows[rowNumber - 1] = new Row(this, rowNumber);
      }
      return row;
    },
    getCell(rowNumber, colNumber) {
      return this.getRow(rowNumber).getCell(colNumber);
    },
    getColumnKey(key) {
      return this._keys[key];
    },
    setColumnKey(key, value) {
      this._keys[key] = value;
    },
    deleteColumnKey(key) {
      delete this._keys[key];
    },
    eachColumnKey(f) {
      for (const [key, value] of Object.entries(this._keys)) {
        f(value, key);
      }
    },
    eachRow(opt, f) {
      if (!f) {
        f = opt;
        opt = {};
      }
      if (opt && opt.includeEmpty) {
        const n = this.rows.length;
        for (let i = 1; i <= n; i++) {
          f(this.getRow(i), i);
        }
      } else {
        this.rows.forEach((r, i) => {
          if (r) {
            f(r, i + 1);
          }
        });
      }
    },
  };
}

const testUtils = {
  views: tools.fix(__json_0),
  testValues: tools.fix(__json_1),
  styles: tools.fix(__json_2),
  properties: tools.fix(__json_3),
  pageSetup: tools.fix(__json_4),
  conditionalFormatting: tools.fix(
    __json_6
  ),
  headerFooter: tools.fix(__json_5),
  createSheetMock,

  createTestBook(workbook, docType, sheets) {
    const options = getOptions(docType);
    sheets = sheets || ['values'];

    workbook.views = [
      {x: 1, y: 2, width: 10000, height: 20000, firstSheet: 0, activeTab: 0},
    ];

    sheets.forEach(sheet => {
      const testSheet = getPath(testSheets, sheet);
      testSheet.addSheet(workbook, options);
    });

    return workbook;
  },

  checkTestBook(workbook, docType, sheets, options) {
    options = getOptions(docType, options);
    sheets = sheets || ['values'];

    expect(workbook).to.not.be.undefined();

    if (options.checkViews) {
      expect(workbook.views).to.deep.equal([
        {
          x: 1,
          y: 2,
          width: 10000,
          height: 20000,
          firstSheet: 0,
          activeTab: 0,
          visibility: 'visible',
        },
      ]);
    }

    sheets.forEach(sheet => {
      const testSheet = getPath(testSheets, sheet);
      testSheet.checkSheet(workbook, options);
    });
  },

  checkTestBookReader: testWorkbookReader.checkBook,
};

export {createSheetMock};
export default testUtils;
