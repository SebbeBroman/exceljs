function ColumnSum(columns) {
  this.columns = columns;
  this.sums = [];
  this.count = 0;
  for (const column of this.columns) {
    this.sums[column] = 0;
  }
}

ColumnSum.prototype = {
  add(row) {
    for (const column of this.columns) {
      this.sums[column] += row.getCell(column).value;
    }
    this.count++;
  },

  toString() {
    return this.sums.join(', ');
  },
  toAverages() {
    return this.sum
      .map(value => (value ? value / this.count : value))
      .join(', ');
  },
};

export default ColumnSum;
