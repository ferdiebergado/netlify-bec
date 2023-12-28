import path from 'node:path';
import { readFile } from 'node:fs/promises';
import * as Excel from 'exceljs';
import { ExpenditureMatrix } from '../../src/server/expenditureMatrix';
import config from '../../src/server/config';
import type { ActivityInfo } from '../../src/types/globals';
import { extractResult } from '../../src/server/utils';

describe('Expenditure Matrix class', () => {
  let expenditureMatrix: ExpenditureMatrix;

  beforeEach(async () => {
    const emFile = config.paths.emTemplate;
    expenditureMatrix = await ExpenditureMatrix.createAsync(emFile);
  });

  it('should correctly write the activity information', async () => {
    const expected: Partial<ActivityInfo> = {
      program: 'Cosmic Learning System',
      output: 'Oriented galaxy heads',
      outputIndicator: 'No. of galaxy heads oriented',
      activityTitle: 'Orientation of Galaxy Heads on Cosmic Education System',
      activityIndicator: 'No. of orientations conducted',
      outputPhysicalTarget: 150,
      activityPhysicalTarget: 1,
    };

    const filename = path.join(config.paths.data, 'be_test.xlsx');
    const buffer = await readFile(filename);
    const result = await expenditureMatrix.convert([{ filename, buffer }]);
    const wb = new Excel.Workbook();

    await wb.xlsx.load(result);

    const sheet = wb.getWorksheet(1);

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

  it('should correctly write the board and lodging expenses');
});
