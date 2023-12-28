import { isCellFormulaValue, extractResult } from '../../src/server/utils';

describe('isCellFormulaValue', () => {
  test('returns true if object is a Cell Formula', () => {
    const formula = {
      formula: 'A2',
      result: undefined,
    };
    const actual = isCellFormulaValue(formula);
    const expected = true;

    expect(actual).toBe(expected);
  });
});

describe('extractResult', () => {
  test('returns the result of a formula from a Cell Formula', () => {
    const result = 10;
    const formula = {
      formula: 'A2',
      result,
    };
    const actual = extractResult(formula);
    const expected = result;

    expect(actual).toBe(expected);
  });
});
