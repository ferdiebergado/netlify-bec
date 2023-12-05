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
  for (let j = 0; j < numRows; j++) {
    const newRow = ws.insertRow(targetRowIndex, []);
    const srcRow = ws.getRow(srcRowIndex);

    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = newRow.getCell(colNumber);

      targetCell.value = cell.value;
      targetCell.style = cell.style;
      targetCell.dataValidation = cell.dataValidation;
    });

    targetRowIndex++;
    srcRowIndex++;
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
  let activityRowIndex = targetRow;

  if (isFirst) {
    activityRowIndex = EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX;
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

  const sumFormula = (cell: string) => {
    return {
      formula: `SUM(${cell}${activityRowIndex + 1}:${cell}${
        activityRowIndex + expenseItems.length
      })`,
    };
  };

  const activityRow = ws.getRow(activityRowIndex);

  // activity
  activityRow.getCell(EXPENDITURE_MATRIX.ACTIVITIES_COL).value = activityTitle;

  // activity indicator
  activityRow.getCell(EXPENDITURE_MATRIX.PERFORMANCE_INDICATOR_COL).value =
    activityIndicator;

  // activity physical target
  activityRow.getCell(
    EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_COL_INDEX + month,
  ).value = activityPhysicalTarget;

  // costing grand total
  activityRow.getCell(EXPENDITURE_MATRIX.TOTAL_COST_COL).value = sumFormula(
    EXPENDITURE_MATRIX.TOTAL_COST_COL,
  );

  // physical target grand total
  activityRow.getCell(EXPENDITURE_MATRIX.PHYSICAL_TARGET_TOTAL_COL).value = {
    formula: `SUM(${EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_START_COL_INDEX}${activityRowIndex}:${EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_END_COL_INDEX}${activityRowIndex})`,
  };

  // obligation and disbursement grand total per month
  [
    EXPENDITURE_MATRIX.OBLIGATION_MONTH_COL_INDEX,
    EXPENDITURE_MATRIX.DISBURSEMENT_MONTH_COL_INDEX,
  ].forEach(c => {
    for (let i = 0; i < 12; i++) {
      const cell = activityRow.getCell(c + i);
      const col = cell.$col$row;
      const re = /\$([A-Z]+)\$/;
      const matches = re.exec(col);

      cell.value = sumFormula(matches![1]);
    }
  });

  // obligation grand total
  activityRow.getCell(EXPENDITURE_MATRIX.TOTAL_OBLIGATION_COL).value =
    sumFormula(EXPENDITURE_MATRIX.TOTAL_OBLIGATION_COL);

  // disbursement grand total
  activityRow.getCell(EXPENDITURE_MATRIX.TOTAL_DISBURSEMENT_COL).value =
    sumFormula(EXPENDITURE_MATRIX.TOTAL_DISBURSEMENT_COL);
}

function createOutputRow(
  ws: Worksheet,
  targetRow: number,
  activity: Activity,
  rank: number,
  isFirst = false,
) {
  let outputRowIndex = targetRow;

  if (isFirst) {
    outputRowIndex = EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX;
  } else {
    duplicateOutput(ws, targetRow);
  }

  const { output, outputIndicator, outputPhysicalTarget, month } = activity;

  // output
  const outputRow = ws.getRow(outputRowIndex);
  outputRow.getCell(EXPENDITURE_MATRIX.OUTPUT_COL).value = output;
  outputRow.getCell(EXPENDITURE_MATRIX.RANK_COL).value = rank;

  // output indicator
  outputRow.getCell(EXPENDITURE_MATRIX.PERFORMANCE_INDICATOR_COL).value =
    outputIndicator;

  // output physical target
  outputRow.getCell(
    EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_COL_INDEX + month,
  ).value = outputPhysicalTarget;

  // physical target grand total
  outputRow.getCell(EXPENDITURE_MATRIX.PHYSICAL_TARGET_TOTAL_COL).value = {
    formula: `SUM(${EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_START_COL_INDEX}${outputRowIndex}:${EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_END_COL_INDEX}${outputRowIndex})`,
  };
}

function orderByProgram(a: Activity, b: Activity): number {
  if (a.program < b.program) {
    return -1;
  } else if (a.program > b.program) {
    return 1;
  } else {
    if (a.output < b.output) {
      return -1;
    } else {
      return 1;
    }
  }
}

function createExpenseItemRow(
  ws: Worksheet,
  targetRow: number,
  expense: ExpenseItem,
  month: number,
  isFirst = false,
) {
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
  const expenseGroupCell = currentRow.getCell(
    EXPENDITURE_MATRIX.EXPENSE_GROUP_COL,
  );
  expenseGroupCell.value = expenseGroup;

  // gaa object
  const gaaObjectCell = currentRow.getCell(EXPENDITURE_MATRIX.GAA_OBJECT_COL);
  gaaObjectCell.value = gaaObject;

  // expense item
  currentRow.getCell(EXPENDITURE_MATRIX.EXPENSE_ITEM_COL).value = expenseItem;

  // quantity
  currentRow.getCell(EXPENDITURE_MATRIX.QUANTITY_COL).value = quantity;

  // unit cost
  currentRow.getCell(EXPENDITURE_MATRIX.UNIT_COST_COL).value = unitCost;

  // frequency
  currentRow.getCell(EXPENDITURE_MATRIX.FREQUENCY_COL).value = freq || 1;

  // total amount
  currentRow.getCell(EXPENDITURE_MATRIX.TOTAL_COST_COL).value = {
    formula: `${EXPENDITURE_MATRIX.QUANTITY_COL}${rowIndex}*${EXPENDITURE_MATRIX.UNIT_COST_COL}${rowIndex}*${EXPENDITURE_MATRIX.FREQUENCY_COL}${rowIndex}`,
  };

  // tev location
  currentRow.getCell(EXPENDITURE_MATRIX.TEV_LOCATION_COL).value = tevLocation;

  // ppmp
  const ppmpCell = currentRow.getCell(EXPENDITURE_MATRIX.PPMP_COL);
  ppmpCell.dataValidation = YES_NO_VALIDATION;
  if (ppmp) ppmpCell.value = YES;

  // app supplies
  const appSuppliesCell = currentRow.getCell(
    EXPENDITURE_MATRIX.APP_SUPPLIES_COL,
  );
  appSuppliesCell.dataValidation = YES_NO_VALIDATION;
  if (appSupplies) appSuppliesCell.value = YES;

  // app ticket
  const appTicketCell = currentRow.getCell(EXPENDITURE_MATRIX.APP_TICKET_COL);
  appTicketCell.dataValidation = YES_NO_VALIDATION;
  if (appTicket) appTicketCell.value = YES;

  // manner of release
  const mannerOfReleaseCell = currentRow.getCell(
    EXPENDITURE_MATRIX.MANNER_OF_RELEASE_COL,
  );
  mannerOfReleaseCell.value = mannerOfRelease;
  mannerOfReleaseCell.dataValidation = MANNER_VALIDATION;

  // total obligation
  currentRow.getCell(EXPENDITURE_MATRIX.TOTAL_OBLIGATION_COL).value = {
    formula: `SUM(${EXPENDITURE_MATRIX.OBLIGATION_MONTH_START_COL}${rowIndex}:${EXPENDITURE_MATRIX.OBLIGATION_MONTH_END_COL}${rowIndex})`,
  };

  const totalRef = {
    formula: `${EXPENDITURE_MATRIX.TOTAL_COST_COL}${rowIndex}`,
  };

  // obligation month
  currentRow.getCell(
    EXPENDITURE_MATRIX.OBLIGATION_MONTH_COL_INDEX + month,
  ).value = totalRef;

  // total disbursement
  currentRow.getCell(EXPENDITURE_MATRIX.TOTAL_DISBURSEMENT_COL).value = {
    formula: `SUM(${EXPENDITURE_MATRIX.DISBURSEMENT_MONTH_START_COL}${rowIndex}:${EXPENDITURE_MATRIX.DISBURSEMENT_MONTH_END_COL}${rowIndex})`,
  };

  // disbursement month
  currentRow.getCell(
    EXPENDITURE_MATRIX.DISBURSEMENT_MONTH_COL_INDEX + month,
  ).value = totalRef;
}

export {
  orderByProgram,
  duplicateProgram,
  createOutputRow,
  createActivityRow,
  duplicateExpenseItem,
  createExpenseItemRow,
};
