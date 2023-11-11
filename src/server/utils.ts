import { CellFormulaValue, CellValue } from 'exceljs';

const isCellFormulaValue = (obj: any): obj is CellFormulaValue => {
  if (obj.result) return true;
  return false;
};

const extractResult = (value: CellValue): number => {
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
};

const timestamp = (): number => {
  return new Date().getTime();
};

export { isCellFormulaValue, extractResult, timestamp };
