import { Worksheet } from 'exceljs';
import {
  EXPENDITURE_MATRIX,
  MANNER_VALIDATION,
  YES,
  YES_NO_VALIDATION,
} from './constants';
import { Activity, ExpenseItem } from './types';

/**
 * Duplicates a specified row/rows
 *
 * @param ws {Worksheet} - sheet were the rows will be duplicated
 * @param targetRowIndex {number} - index where the duplicate rows will be inserted
 * @param srcRowIndex {number} - index of the row that will be duplicated
 * @param numRows {number} - number of rows to be duplicated
 *
 * @returns void
 */
function duplicateRows(
  ws: Worksheet,
  targetRowIndex: number,
  srcRowIndex: number,
  numRows: number,
) {
  for (let j = 0; j < numRows; j += 1) {
    const newRow = ws.insertRow(targetRowIndex, []);
    const srcRow = ws.getRow(srcRowIndex);

    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = newRow.getCell(colNumber);

      targetCell.value = cell.value;
      targetCell.style = cell.style;
      targetCell.dataValidation = cell.dataValidation;
    });

    // eslint-disable-next-line no-param-reassign
    targetRowIndex += 1;
    // eslint-disable-next-line no-param-reassign
    srcRowIndex += 1;
  }
}

function duplicateProgram(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX, 1);
}

function duplicateOutput(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX, 1);
}

function duplicateActivity(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX, 1);
}

function duplicateExpenseItem(ws: Worksheet, targetRow: number, count: number) {
  duplicateRows(
    ws,
    targetRow,
    EXPENDITURE_MATRIX.EXPENSE_ITEM_ROW_INDEX,
    count,
  );
}

function createActivityRow(
  ws: Worksheet,
  targetRow: number,
  activity: Activity,
  isFirst: boolean = false,
) {
  const {
    ACTIVITY_ROW_INDEX,
    ACTIVITIES_COL,
    PERFORMANCE_INDICATOR_COL,
    PHYSICAL_TARGET_MONTH_COL_INDEX,
    PHYSICAL_TARGET_TOTAL_COL,
    TOTAL_COST_COL,
    PHYSICAL_TARGET_MONTH_START_COL_INDEX,
    PHYSICAL_TARGET_MONTH_END_COL_INDEX,
    OBLIGATION_MONTH_COL_INDEX,
    DISBURSEMENT_MONTH_COL_INDEX,
    TOTAL_OBLIGATION_COL,
    TOTAL_DISBURSEMENT_COL,
  } = EXPENDITURE_MATRIX;

  let activityRowIndex = targetRow;

  if (isFirst) {
    activityRowIndex = ACTIVITY_ROW_INDEX;
  } else {
    duplicateActivity(ws, targetRow);
  }

  const {
    activityTitle,
    activityIndicator,
    activityPhysicalTarget,
    month,
    expenseItems,
  } = activity;

  const sumFormula = (cell: string) => ({
    formula: `SUM(${cell}${activityRowIndex + 1}:${cell}${
      activityRowIndex + expenseItems.length
    })`,
  });

  const activityRow = ws.getRow(activityRowIndex);

  // activity
  activityRow.getCell(ACTIVITIES_COL).value = activityTitle;

  // activity indicator
  activityRow.getCell(PERFORMANCE_INDICATOR_COL).value = activityIndicator;

  // activity physical target
  activityRow.getCell(PHYSICAL_TARGET_MONTH_COL_INDEX + month).value =
    activityPhysicalTarget;

  // costing grand total
  activityRow.getCell(TOTAL_COST_COL).value = sumFormula(TOTAL_COST_COL);

  // physical target grand total
  const physicalTargetMonthStartCell = `${PHYSICAL_TARGET_MONTH_START_COL_INDEX}${activityRowIndex}`;
  const physicalTargetMonthEndCell = `${PHYSICAL_TARGET_MONTH_END_COL_INDEX}${activityRowIndex}`;

  activityRow.getCell(PHYSICAL_TARGET_TOTAL_COL).value = {
    formula: `SUM(${physicalTargetMonthStartCell}:${physicalTargetMonthEndCell})`,
  };

  // obligation and disbursement grand total per month
  [OBLIGATION_MONTH_COL_INDEX, DISBURSEMENT_MONTH_COL_INDEX].forEach(c => {
    for (let i = 0; i < 12; i += 1) {
      const cell = activityRow.getCell(c + i);
      const col = cell.$col$row;
      const re = /\$([A-Z]+)\$/;
      const matches = re.exec(col);

      cell.value = sumFormula(matches![1]);
    }
  });

  // obligation grand total
  activityRow.getCell(TOTAL_OBLIGATION_COL).value =
    sumFormula(TOTAL_OBLIGATION_COL);

  // disbursement grand total
  activityRow.getCell(TOTAL_DISBURSEMENT_COL).value = sumFormula(
    TOTAL_DISBURSEMENT_COL,
  );
}

function createOutputRow(
  ws: Worksheet,
  targetRow: number,
  activity: Activity,
  rank: number,
  isFirst = false,
) {
  const {
    OUTPUT_ROW_INDEX,
    OUTPUT_COL,
    RANK_COL,
    PERFORMANCE_INDICATOR_COL,
    PHYSICAL_TARGET_MONTH_COL_INDEX,
    PHYSICAL_TARGET_TOTAL_COL,
    PHYSICAL_TARGET_MONTH_START_COL_INDEX,
    PHYSICAL_TARGET_MONTH_END_COL_INDEX,
  } = EXPENDITURE_MATRIX;

  let outputRowIndex = targetRow;

  if (isFirst) {
    outputRowIndex = OUTPUT_ROW_INDEX;
  } else {
    duplicateOutput(ws, targetRow);
  }

  const { output, outputIndicator, outputPhysicalTarget, month } = activity;

  // output
  const outputRow = ws.getRow(outputRowIndex);
  outputRow.getCell(OUTPUT_COL).value = output;
  outputRow.getCell(RANK_COL).value = rank;

  // output indicator
  outputRow.getCell(PERFORMANCE_INDICATOR_COL).value = outputIndicator;

  // output physical target
  outputRow.getCell(PHYSICAL_TARGET_MONTH_COL_INDEX + month).value =
    outputPhysicalTarget;

  // physical target grand total
  const physicalTargetMonthStartCell = `${PHYSICAL_TARGET_MONTH_START_COL_INDEX}${outputRowIndex}`;
  const physicalTargetMonthEndCell = `${PHYSICAL_TARGET_MONTH_END_COL_INDEX}${outputRowIndex}`;

  outputRow.getCell(PHYSICAL_TARGET_TOTAL_COL).value = {
    formula: `SUM(${physicalTargetMonthStartCell}:${physicalTargetMonthEndCell})`,
  };
}

function orderByProgram(a: Activity, b: Activity): number {
  if (a.program < b.program) {
    return -1;
  }

  if (a.program > b.program) {
    return 1;
  }

  if (a.output < b.output) {
    return -1;
  }

  return 1;
}

function createExpenseItemRow(
  ws: Worksheet,
  targetRow: number,
  expense: ExpenseItem,
  month: number,
  isFirst = false,
) {
  const {
    EXPENSE_GROUP_COL,
    GAA_OBJECT_COL,
    EXPENSE_ITEM_COL,
    QUANTITY_COL,
    UNIT_COST_COL,
    FREQUENCY_COL,
    TOTAL_COST_COL,
    TEV_LOCATION_COL,
    PPMP_COL,
    APP_SUPPLIES_COL,
    APP_TICKET_COL,
    MANNER_OF_RELEASE_COL,
    TOTAL_OBLIGATION_COL,
    OBLIGATION_MONTH_START_COL,
    OBLIGATION_MONTH_END_COL,
    OBLIGATION_MONTH_COL_INDEX,
    TOTAL_DISBURSEMENT_COL,
    DISBURSEMENT_MONTH_START_COL,
    DISBURSEMENT_MONTH_END_COL,
    DISBURSEMENT_MONTH_COL_INDEX,
  } = EXPENDITURE_MATRIX;

  let rowIndex = targetRow;

  if (isFirst) rowIndex = targetRow - 1;

  const currentRow = ws.getRow(rowIndex);

  const {
    expenseGroup,
    gaaObject,
    expenseItem,
    quantity,
    unitCost,
    freq,
    tevLocation,
    ppmp,
    appSupplies,
    appTicket,
    mannerOfRelease,
  } = expense;

  // expense group
  const expenseGroupCell = currentRow.getCell(EXPENSE_GROUP_COL);
  expenseGroupCell.value = expenseGroup;

  // gaa object
  const gaaObjectCell = currentRow.getCell(GAA_OBJECT_COL);
  gaaObjectCell.value = gaaObject;

  // expense item
  currentRow.getCell(EXPENSE_ITEM_COL).value = expenseItem;

  // quantity
  currentRow.getCell(QUANTITY_COL).value = quantity;

  // unit cost
  currentRow.getCell(UNIT_COST_COL).value = unitCost;

  // frequency
  currentRow.getCell(FREQUENCY_COL).value = freq || 1;

  // total amount
  currentRow.getCell(TOTAL_COST_COL).value = {
    formula: `${QUANTITY_COL}${rowIndex}*${UNIT_COST_COL}${rowIndex}*${FREQUENCY_COL}${rowIndex}`,
  };

  // tev location
  currentRow.getCell(TEV_LOCATION_COL).value = tevLocation;

  // ppmp
  const ppmpCell = currentRow.getCell(PPMP_COL);
  ppmpCell.dataValidation = YES_NO_VALIDATION;
  if (ppmp) ppmpCell.value = YES;

  // app supplies
  const appSuppliesCell = currentRow.getCell(APP_SUPPLIES_COL);
  appSuppliesCell.dataValidation = YES_NO_VALIDATION;
  if (appSupplies) appSuppliesCell.value = YES;

  // app ticket
  const appTicketCell = currentRow.getCell(APP_TICKET_COL);
  appTicketCell.dataValidation = YES_NO_VALIDATION;
  if (appTicket) appTicketCell.value = YES;

  // manner of release
  const mannerOfReleaseCell = currentRow.getCell(MANNER_OF_RELEASE_COL);
  mannerOfReleaseCell.value = mannerOfRelease;
  mannerOfReleaseCell.dataValidation = MANNER_VALIDATION;

  // total obligation
  const obligationMonthStartCell = `${OBLIGATION_MONTH_START_COL}${rowIndex}`;
  const obligationMonthEndCell = `${OBLIGATION_MONTH_END_COL}${rowIndex}`;

  currentRow.getCell(TOTAL_OBLIGATION_COL).value = {
    formula: `SUM(${obligationMonthStartCell}:${obligationMonthEndCell})`,
  };

  const totalRef = {
    formula: `${TOTAL_COST_COL}${rowIndex}`,
  };

  // obligation month
  currentRow.getCell(OBLIGATION_MONTH_COL_INDEX + month).value = totalRef;

  // total disbursement
  const disbursementMonthStartCell = `${DISBURSEMENT_MONTH_START_COL}${rowIndex}`;
  const disbursementMonthEndCell = `${DISBURSEMENT_MONTH_END_COL}${rowIndex}`;

  currentRow.getCell(TOTAL_DISBURSEMENT_COL).value = {
    formula: `SUM(${disbursementMonthStartCell}:${disbursementMonthEndCell})`,
  };

  // disbursement month
  currentRow.getCell(DISBURSEMENT_MONTH_COL_INDEX + month).value = totalRef;
}

export {
  orderByProgram,
  duplicateProgram,
  createOutputRow,
  createActivityRow,
  duplicateExpenseItem,
  createExpenseItemRow,
};
