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

/**
 * Converts uploaded files into an ArrayBuffer representing the expenditure matrix.
 *
 * @param {Express.Multer.File[]} files - The array of uploaded files.
 * @returns {Promise<ArrayBuffer>} A promise that resolves to the ArrayBuffer of the expenditure matrix.
 */
export default async function convert(
  files: Express.Multer.File[],
): Promise<ArrayBuffer> {
  const em = new Workbook();
  await em.xlsx.readFile(config.paths.emTemplate);
  const emWs = em.getWorksheet(1);

  if (!emWs) throw new Error('Sheet not found');

  const programs: string[] = [];
  const activities: Activity[] = [];

  let targetRow = EXPENDITURE_MATRIX.TARGET_ROW_INDEX;
  let isFirst = true;
  let rank = 1;

  const loadAndParseWorkbook = async (
    file: Express.Multer.File,
  ): Promise<void> => {
    const { originalname, buffer } = file;

    // eslint-disable-next-line no-console
    console.log('processing file:', originalname);

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
        console.log('processing sheet:', name);

        const activity = parseActivity(sheet);
        activities.push(activity);
      });

      resolve();
    });
  };

  // Use Promise.all to wait for all promises to resolve
  await Promise.all(files.map(file => loadAndParseWorkbook(file)));

  const activityRows: number[] = [];

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

    const activityRowIndex = createActivityRow(
      emWs,
      targetRow,
      activity,
      isFirst,
    );

    activityRows.push(activityRowIndex);

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

  const { TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL } =
    EXPENDITURE_MATRIX;

  const lastRowIndex = targetRow + 2;

  const grandTotalRow = emWs.getRow(lastRowIndex);

  const setGrandTotalCell = (cell: string) => {
    const totalCells = activityRows.map(row => cell + row);
    grandTotalRow.getCell(cell).value = {
      formula: `SUM(${totalCells.toString()})`,
    };
  };

  setGrandTotalCell(TOTAL_COST_COL);
  setGrandTotalCell(TOTAL_OBLIGATION_COL);
  setGrandTotalCell(TOTAL_DISBURSEMENT_COL);

  for (let i = 0; i < 26; i += 1) {
    const col = i + 44;
    const cell = grandTotalRow.getCell(col);
    cell.value = {
      formula: `SUM(${activityRows.map(
        row => cell.address.replace(/\d+/, '') + row,
      )})`,
    };
  }

  const outBuff = await em.xlsx.writeBuffer();

  return outBuff;
}
