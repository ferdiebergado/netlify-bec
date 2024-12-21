import { CellFormulaValue, CellValue } from 'exceljs';

/**
 * Checks if the given object is a CellFormulaValue.
 *
 * @param {any} obj The object to check.
 * @returns {obj is CellFormulaValue} Returns true if the object is a CellFormulaValue, otherwise false.
 */
export function isCellFormulaValue(
  value: CellValue,
): value is CellFormulaValue {
  return (
    typeof value === 'object' &&
    value !== null &&
    'formula' in value &&
    typeof (value as CellFormulaValue).formula === 'string'
  );
}

/**
 * Extracts a numeric result from a CellValue or CellFormulaValue object.
 *
 * @param {CellValue} value The CellValue to extract the numeric result from.
 * @returns {number} The extracted numeric result, or 0 if not found.
 */
export function extractResult(value: CellValue): number {
  if (value) {
    if (typeof value === 'number') return value;

    if (isCellFormulaValue(value)) {
      const { result } = value;

      if (typeof result === 'number') {
        return result;
      }
    }
  }

  return 0;
}

/**
 * Creates a timestamp representing the current time in milliseconds.
 *
 * @returns {number} The timestamp representing the current time.
 */
export function createTimestamp(): number {
  return new Date().getTime();
}

/**
 * Converts a cell value represented as a string to a number.
 *
 * @param {string} cellValue The string representation of the cell value.
 * @returns {number} The numeric representation of the cell value, or 0 if conversion fails.
 */
export function getCellValueAsNumber(cellValue: string): number {
  const numericValue = +cellValue;
  return Number.isNaN(numericValue) ? 0 : numericValue;
}

// Sequence generator function (commonly referred to as "range", cf. Python, Clojure, etc.)
export function range(start: number, stop: number, step: number): number[] {
  return Array.from(
    { length: Math.ceil((stop - start) / step) },
    (_, i) => start + i * step,
  );
}
