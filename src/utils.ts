import { Cell, CellFormulaValue, CellValue } from 'exceljs';

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
 * @param {Cell} cell The cell to process.
 * @returns {number} The numeric representation of the cell value, or 0 if conversion fails.
 */
export function getCellValueAsNumber(cell: Cell): number {
  const numericValue = +cell.text;
  return Number.isNaN(numericValue) ? 0 : numericValue;
}

// Sequence generator function (commonly referred to as "range", cf. Python, Clojure, etc.)
export function range(start: number, stop: number, step: number): number[] {
  return Array.from(
    { length: Math.ceil((stop - start) / step) },
    (_, i) => start + i * step,
  );
}

// Recursively freeze an object
export function deepFreeze<T extends Object>(obj: T): T {
  Object.freeze(obj);
  Object.keys(obj).forEach(key => {
    const value = (obj as any)[key];
    if (typeof value === 'object' && value !== null) {
      deepFreeze(value);
    }
  });
  return obj;
}

/**
 * Parses the input string and extracts the program title.
 * @param {string} inputString - The input string in the specified format.
 * @returns {string} - The extracted project name.
 */
export function extractProgramTitle(inputString: string): string | undefined {
  // Match the pattern for the project name
  const match = inputString.match(/- (.+?) \(/);

  // Return the project name if a match is found, otherwise an empty string
  return match ? match[1] : '';
}
