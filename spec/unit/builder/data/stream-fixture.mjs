export function fixture(useSharedStrings, useStyles) {
  const sparse = []; sparse[3] = 'sparse'; sparse[7] = 77;
  return {
    creator: 'stream parity', created: new Date('2020-01-01'), modified: new Date('2020-01-01'),
    useSharedStrings, useStyles,
    sheets: [{
      name: 'Values', state: 'hidden', views: [{state: 'frozen', ySplit: 1}],
      pageSetup: {orientation: 'landscape', fitToWidth: 1, fitToHeight: 0},
      properties: {defaultRowHeight: 22, outlineLevelCol: 2},
      headerFooter: {oddHeader: 'Header', oddFooter: '&P'}, autoFilter: 'A1:D9',
      columns: [{header: ['Date', 'ignored'], key: 'date', width: 20, style: {numFmt: 'yyyy-mm-dd'}}, {header: 'Label', key: 'label', width: 28, style: {font: {italic: true}}}, {header: 'Number', key: 'n', hidden: true, outlineLevel: 1}, {width: 12, outlineLevel: 2}],
      rows: [
        {date: new Date('2021-02-03'), label: 'Ada', n: 12},
        [new Date('2022-04-05'), {richText: [{text: 'Bold', font: {bold: true}}, {text: ' plain'}]}, false, {error: '#N/A'}],
        [null, {text: 'Link', hyperlink: 'https://example.com?a=1&b=2', tooltip: 'tip'}, {formula: 'C2*2', result: 24}],
        [], [null], sparse, ['<& unicode é'],
      ],
    }, {name: 'Empty', rows: []}, {name: 'Last', rows: [[1, 2], [3, 4]]}],
  };
}
