import type {DataValidation} from '../../index.js';

export type DataValidationsModel = Record<string, DataValidation | undefined>;

class DataValidations {
  model: DataValidationsModel;

  constructor(model?: DataValidationsModel | null) {
    this.model = model || {};
  }

  add(address: string, validation: DataValidation): DataValidation {
    return (this.model[address] = validation);
  }

  find(address: string): DataValidation | undefined {
    return this.model[address];
  }

  remove(address: string): void {
    this.model[address] = undefined;
  }
}

export default DataValidations;
export {DataValidations};
