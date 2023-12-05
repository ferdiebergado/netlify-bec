import { CellFormulaValue, CellValue } from 'exceljs';

function isCellFormulaValue(obj: any): obj is CellFormulaValue {
  return 'result' in obj;
}

function extractResult(value: CellValue): number {
  if (value) {
    if (typeof value === 'number') return value;

    if (isCellFormulaValue(value)) {
      const result = value.result;

      if (result && typeof result === 'number') {
        return result;
      }
    }
  }

  return 0;
}

function createTimestamp(): number {
  return new Date().getTime();
}

function getCellValueAsNumber(cellValue: string): number {
  const numericValue = +cellValue;
  return isNaN(numericValue) ? 0 : numericValue;
}

export {
  isCellFormulaValue,
  extractResult,
  createTimestamp,
  getCellValueAsNumber,
};
