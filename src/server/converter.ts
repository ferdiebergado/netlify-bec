import { Workbook } from 'exceljs';
import { Activity } from './types';
import config from './config';
import parseActivity from './budget_estimate';
import {
  AUXILLIARY_SHEETS,
  BUDGET_ESTIMATE,
  EXPENDITURE_MATRIX,
} from './constants';
import {
  createActivityRow,
  createExpenseItemRow,
  createOutputRow,
  duplicateExpenseItem,
  duplicateProgram,
  orderByProgram,
} from './expenditure_matrix';

export default async function convert(buffers: Buffer[]): Promise<ArrayBuffer> {
  const em = new Workbook();
  await em.xlsx.readFile(config.paths.emTemplate);
  const emWs = em.getWorksheet(1);

  if (!emWs) throw new Error('Sheet not found');

  const programs: string[] = [];
  const activities: Activity[] = [];

  let targetRow = EXPENDITURE_MATRIX.TARGET_ROW_INDEX;
  let isFirst = true;
  let rank = 1;

  for (const buffer of buffers) {
    const be = new Workbook();
    await be.xlsx.load(buffer);

    // eslint-disable-next-line @typescript-eslint/no-loop-func
    be.eachSheet(sheet => {
      const { name } = sheet;

      // skip auxilliary sheets
      if (AUXILLIARY_SHEETS.includes(name)) {
        console.log('skipping', name);
        return;
      }

      // skip if not a budget estimate template
      if (
        sheet.getCell(BUDGET_ESTIMATE.PROGRAM_HEADING_CELL).text !== 'PROGRAM:'
      ) {
        console.log('skipping non budget estimate sheet:', name);
        return;
      }

      console.log('processing', name);

      const activity = parseActivity(sheet);

      activities.push(activity);
    });
  }

  activities.sort(orderByProgram).forEach(activity => {
    const { program, month, expenseItems } = activity;

    let programRowIndex: number = targetRow;

    // program
    if (isFirst) {
      programRowIndex = EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX;
      programs.push(program);
    } else {
      if (!programs.includes(program)) {
        duplicateProgram(emWs, targetRow);
        programs.push(program);
        targetRow++;
      }
    }

    const programRow = emWs.getRow(programRowIndex);
    programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;

    createOutputRow(emWs, targetRow, activity, rank, isFirst);

    rank++;

    if (!isFirst) targetRow++;

    createActivityRow(emWs, targetRow, activity, isFirst);

    if (!isFirst) targetRow++;

    // expense items
    duplicateExpenseItem(emWs, targetRow, expenseItems.length);

    for (const expense of expenseItems) {
      createExpenseItemRow(emWs, targetRow, expense, month, isFirst);

      targetRow++;
    }

    isFirst = false;
    targetRow--;
  });

  emWs.spliceRows(targetRow + 1, 2);

  console.log('writing to buffer...');

  const outBuff = await em.xlsx.writeBuffer();

  return outBuff;
}
