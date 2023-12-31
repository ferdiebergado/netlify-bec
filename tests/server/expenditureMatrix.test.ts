import path from 'node:path';
import { readFile } from 'node:fs/promises';
import * as Excel from 'exceljs';
import { ExpenditureMatrix } from '../../src/server/expenditureMatrix';
import config from '../../src/server/config';
import type { ActivityInfo, ExpenseItem } from '../../src/types/globals';
import { extractResult } from '../../src/server/utils';

describe('Expenditure Matrix class', () => {
  let expenditureMatrix: ExpenditureMatrix;
  let sheet: Excel.Worksheet | undefined;

  function assertExpenseItem(row: number, expected: Record<string, any>) {
    const expenseGroup = sheet?.getCell(`J${row}`).text;
    expect(expenseGroup).toEqual(expected.expenseGroup);

    const gaaObject = sheet?.getCell(`L${row}`).text;
    expect(gaaObject).toEqual(expected.gaaObject);

    const expenseItem = sheet?.getCell(`N${row}`).text;

    expect(expenseItem).toEqual(expected.expenseItem);

    const quantity = sheet?.getCell(`O${row}`).text;

    expect(+quantity!).toBe(expected.quantity);

    const unitCost = sheet?.getCell(`P${row}`).text;

    expect(+unitCost!).toEqual(expected.unitCost);

    const freq = sheet?.getCell(`Q${row}`).text;

    expect(+freq!).toEqual(expected.freq);

    const ppmp = sheet?.getCell(`T${row}`).text;

    expect(ppmp).toEqual(expected.ppmp);

    const releaseManner = sheet?.getCell(`W${row}`).text;

    expect(releaseManner).toEqual(expected.releaseManner);
  }

  beforeEach(async () => {
    const emFile = config.paths.emTemplate as string;
    expenditureMatrix = await ExpenditureMatrix.createAsync(emFile);
    const filename: string = path.join(config.paths.data, 'be_test.xlsx');
    const buffer = await readFile(filename);
    const result = await expenditureMatrix.convert([{ filename, buffer }]);
    const wb = new Excel.Workbook();

    await wb.xlsx.load(result);

    sheet = wb.getWorksheet(1);
  });

  it('should correctly write the activity information', () => {
    const expected: Omit<ActivityInfo, 'month' | 'venue' | 'totalPax'> = {
      program: 'Cosmic Learning System',
      output: 'Oriented galaxy heads',
      outputIndicator: 'No. of galaxy heads oriented',
      activityTitle: 'Orientation of Galaxy Heads on Cosmic Education System',
      activityIndicator: 'No. of orientations conducted',
      outputPhysicalTarget: 150,
      activityPhysicalTarget: 1,
    };

    const program = sheet?.getCell('C13').text;

    expect(program).toEqual(expected.program);

    const output = sheet?.getCell('D14').text;

    expect(output).toEqual(expected.output);

    const outputIndicator = sheet?.getCell('H14').text;

    expect(outputIndicator).toEqual(expected.outputIndicator);

    const activityTitle = sheet?.getCell('G15').text;

    expect(activityTitle).toEqual(expected.activityTitle);

    const activityIndicator = sheet?.getCell('H15').text;

    expect(activityIndicator).toEqual(expected.activityIndicator);

    const outputPhysicalTarget = sheet?.getCell('AM14').value;

    expect(extractResult(outputPhysicalTarget)).toEqual(
      expected.outputPhysicalTarget,
    );

    const activityPhysicalTarget = sheet?.getCell('AM15').value;

    expect(extractResult(activityPhysicalTarget)).toEqual(
      expected.activityPhysicalTarget,
    );
  });

  it('should correctly write the board and lodging expenses', () => {
    const row = 16;
    const expected = {
      expenseGroup: 'Training and Scholarship Expenses',
      gaaObject: 'Training Expenses',
      expenseItem: 'Board and Lodging of Participants',
      quantity: 41,
      unitCost: 1500,
      freq: 5,
      costingTotal: `O${row}*P${row}*Q${row}`,
      ppmp: 'Y',
      releaseManner: 'Direct Payment',
    };

    assertExpenseItem(row, expected);
  });

  it('should correctly write the travel expenses', () => {
    const row = 21;

    const expected = {
      expenseGroup: 'Training and Scholarship Expenses',
      gaaObject: 'Training Expenses',
      expenseItem: 'Travel Expenses of Participants from Region I',
      quantity: 2,
      unitCost: 13900,
      freq: 1,
      costingTotal: `O${row}*P${row}*Q${row}`,
      ppmp: 'N',
      releaseManner: 'For Downloading (Program Support Funds)',
    };

    assertExpenseItem(row, expected);
  });

  it('should correctly write the honorarium', () => {
    const row = 43;
    const expected = {
      expenseGroup: 'Training and Scholarship Expenses',
      gaaObject: 'Training Expenses',
      expenseItem: 'Honorarium of Resource Persons',
      quantity: 8,
      unitCost: 30000,
      freq: 1,
      costingTotal: `O${row}*P${row}*Q${row}`,
      ppmp: 'N',
      releaseManner: 'Direct Payment',
    };

    assertExpenseItem(row, expected);
  });

  it('should correctly write the other expenses', () => {
    const row = 45;
    const expected = {
      expenseGroup: 'Supplies and Materials Expenses',
      gaaObject: 'Other Supplies and Materials Expenses',
      expenseItem: ' Supplies and Materials',
      quantity: 85,
      unitCost: 100,
      freq: 1,
      costingTotal: `O${row}*P${row}*Q${row}`,
      ppmp: 'N',
      releaseManner: 'Cash Advance',
    };

    assertExpenseItem(row, expected);
  });
});
