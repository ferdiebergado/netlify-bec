import { CellFormulaValue, CellValue } from 'exceljs';

function isCellFormulaValue(obj: any): obj is CellFormulaValue {
  return 'result' in obj
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

export { isCellFormulaValue, extractResult, createTimestamp };
