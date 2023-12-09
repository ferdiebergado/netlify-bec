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

  const loadAndParseWorkbook = async (buffer: Buffer): Promise<void> => {
    const be = new Workbook();
    await be.xlsx.load(buffer);

    return new Promise<void>(resolve => {
      be.eachSheet(sheet => {
        const { name } = sheet;

        if (AUXILLIARY_SHEETS.includes(name)) {
          // eslint-disable-next-line no-console
          // console.log('skipping', name);
          return;
        }

        if (
          sheet.getCell(BUDGET_ESTIMATE.PROGRAM_HEADING_CELL).text !==
          'PROGRAM:'
        ) {
          // eslint-disable-next-line no-console
          // console.log('skipping non budget estimate sheet:', name);
          return;
        }

        // eslint-disable-next-line no-console
        // console.log('processing', name);

        const activity = parseActivity(sheet);
        activities.push(activity);
      });

      resolve();
    });
  };

  // Use Promise.all to wait for all promises to resolve
  await Promise.all(buffers.map(buffer => loadAndParseWorkbook(buffer)));

  activities.sort(orderByProgram).forEach(activity => {
    const { program, month, expenseItems } = activity;

    let programRowIndex: number = targetRow;

    // program
    if (isFirst) {
      programRowIndex = EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX;
      programs.push(program);
    } else if (!programs.includes(program)) {
      duplicateProgram(emWs, targetRow);
      programs.push(program);
      targetRow += 1;
    }

    const programRow = emWs.getRow(programRowIndex);
    programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;

    createOutputRow(emWs, targetRow, activity, rank, isFirst);

    rank += 1;

    if (!isFirst) targetRow += 1;

    createActivityRow(emWs, targetRow, activity, isFirst);

    if (!isFirst) targetRow += 1;

    // expense items
    duplicateExpenseItem(emWs, targetRow, expenseItems.length);

    expenseItems.forEach(expense => {
      createExpenseItemRow(emWs, targetRow, expense, month, isFirst);
      targetRow += 1;
    });

    isFirst = false;
    targetRow -= 1;
  });

  emWs.spliceRows(targetRow + 1, 2);

  // console.log('writing to buffer...');

  const outBuff = await em.xlsx.writeBuffer();

  return outBuff;
}
