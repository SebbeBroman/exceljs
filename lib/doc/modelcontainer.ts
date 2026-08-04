import XLSX from '../xlsx/xlsx.js';

class ModelContainer {
  model: unknown;
  private _xlsx?: InstanceType<typeof XLSX>;

  constructor(model: unknown) {
    this.model = model;
  }

  get xlsx(): InstanceType<typeof XLSX> {
    if (!this._xlsx) {
      this._xlsx = new XLSX(this);
    }
    return this._xlsx;
  }
}

export default ModelContainer;
export {ModelContainer};
