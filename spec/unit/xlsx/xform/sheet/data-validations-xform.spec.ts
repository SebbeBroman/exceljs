import DataValidationsXformParser from '../../../../../lib/xlsx/parser/sheet/data-validations-xform.js';
import {describe} from 'vite-plus/test';
import testXformHelper from '../test-xform-helper.js';

import DataValidationsXform from '../../../../../lib/xlsx/xform/sheet/data-validations-xform.js';

const expectations = [
  {
    title: 'list type',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    preparedModel: {
      E1: {
        type: 'list',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['Ducks'],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="list" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="E1">
          <formula1>Ducks</formula1>
        </dataValidation>
      </dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'whole type',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    preparedModel: {
      A1: {
        type: 'whole',
        operator: 'between',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [5, 10],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'decimal type',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    preparedModel: {
      A1: {
        type: 'decimal',
        operator: 'notBetween',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: [5, 10],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="decimal" operator="notBetween" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'custom type',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    preparedModel: {
      A1: {
        type: 'custom',
        allowBlank: true,
        showInputMessage: true,
        showErrorMessage: true,
        formulae: ['OR(C21=5,C21=7)'],
      },
    },
    get parsedModel() {
      return this.preparedModel;
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="custom" allowBlank="1" showInputMessage="1" showErrorMessage="1" sqref="A1">
          <formula1>OR(C21=5,C21=7)</formula1>
        </dataValidation>
      </dataValidations>
    `,
    tests: ['render', 'renderIn', 'parse'],
  },
  {
    title: 'parse open office',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" allowBlank="true" showInputMessage="false" sqref="A1">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    parsedModel: {
      A1: {
        type: 'whole',
        operator: 'between',
        allowBlank: true,
        showInputMessage: false,
        formulae: [5, 10],
      },
    },
    tests: ['parse'],
  },
  {
    title: 'optimised',
    create: () => new DataValidationsXform(),
    createParser: () => new DataValidationsXformParser(),
    preparedModel: {
      A1: {type: 'whole', operator: 'between', formulae: [5, 10]},
      A2: {type: 'whole', operator: 'between', formulae: [5, 10]},
      B1: {type: 'whole', operator: 'between', formulae: [5, 10]},
      B2: {type: 'whole', operator: 'between', formulae: [5, 10]},
    },
    get parsedModel() {
      return {'A1:B2': this.preparedModel.A1};
    },
    xml: `
      <dataValidations count="1">
        <dataValidation type="whole" sqref="A1:B2">
          <formula1>5</formula1>
          <formula2>10</formula2>
        </dataValidation>
      </dataValidations>
    `,
    tests: ['render', 'parse'],
  },
];

describe('DataValidationsXform', () => {
  testXformHelper(expectations);
});
