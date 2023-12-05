import { Worksheet } from 'exceljs';
import {
  BUDGET_ESTIMATE,
  EXPENSE_GROUP,
  GAA_OBJECT,
  MANNER_OF_RELEASE,
} from './constants';
import {
  Activity,
  ExpenseGroup,
  ExpenseItem,
  GAAObject,
  MannerOfRelease,
} from './types';
import { extractResult, getCellValueAsNumber } from './utils';

function createExpenseItem(
  expenseGroup: ExpenseGroup,
  gaaObject: GAAObject,
  expenseItem: string,
  quantity: number,
  freq: number,
  unitCost: number,
  mannerOfRelease: MannerOfRelease,
  tevLocation: string,
  ppmp: boolean,
  appSupplies: boolean,
  appTicket: boolean,
): ExpenseItem {
  return {
    expenseGroup,
    gaaObject,
    expenseItem,
    quantity,
    freq,
    unitCost,
    mannerOfRelease,
    tevLocation,
    ppmp,
    appSupplies,
    appTicket,
  };
}

function getExpenseItems(
  sheet: Worksheet,
  startRowIndex: number,
  startColIndex: number,
  numRows: number,
  prefix: string,
  mannerOfRelease: MannerOfRelease = MANNER_OF_RELEASE.DIRECT_PAYMENT,
) {
  const items: ExpenseItem[] = [];

  for (let i = 0; i < numRows; i++) {
    const row = sheet.getRow(startRowIndex);

    const quantity = getCellValueAsNumber(
      row.getCell(BUDGET_ESTIMATE.QUANTITY_CELL_INDEX).text,
    );
    if (quantity === 0) continue;

    const item = row.getCell(startColIndex).text;

    let expenseGroup: ExpenseGroup =
      EXPENSE_GROUP.TRAINING_SCHOLARSHIPS_EXPENSES;
    let gaaObject: GAAObject = GAA_OBJECT.TRAINING_EXPENSES;
    let tevLocation = '';
    let ppmp = false;
    let appSupplies = false;
    let appTicket = false;

    if (item.toLowerCase().includes('supplies')) {
      expenseGroup = EXPENSE_GROUP.SUPPLIES_EXPENSES;
      gaaObject = GAA_OBJECT.OTHER_SUPPLIES;
      appSupplies = true;
    }

    const expenseItem = `${prefix} ${item}`;

    if (
      expenseItem.toLowerCase().includes('travel') &&
      expenseItem.toLowerCase().includes('participants')
    )
      tevLocation = item;

    const freq = getCellValueAsNumber(
      row.getCell(BUDGET_ESTIMATE.FREQ_CELL_INDEX).text || '1',
    );

    const unitCost = parseFloat(
      row.getCell(BUDGET_ESTIMATE.UNIT_COST_CELL_INDEX).text,
    );

    items.push(
      createExpenseItem(
        expenseGroup,
        gaaObject,
        expenseItem,
        quantity,
        freq,
        unitCost,
        mannerOfRelease,
        tevLocation,
        ppmp,
        appSupplies,
        appTicket,
      ),
    );

    startRowIndex++;
  }

  return items;
}

function boardLodging(sheet: Worksheet) {
  const blPrefix = 'Board and Lodging of';
  const blMannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_BOARD;
  const bl = getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.BOARD_LODGING_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
    4,
    blPrefix,
    blMannerOfRelease,
  );
  const blOthers = getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.BOARD_LODGING_OTHER_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
    1,
    blPrefix,
    blMannerOfRelease,
  );
  return [...bl, ...blOthers];
}

function travelExpenses(sheet: Worksheet) {
  const tevPrefix = 'Travel Expenses of Participants from';
  const tevMannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_PSF;
  const tevPax = getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.TRAVEL_REGION_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_SECOND_COL_INDEX,
    18,
    tevPrefix,
    tevMannerOfRelease,
  );

  const tevPrefixNonPax = 'Travel Expenses of';
  const tevNonPax = getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.TRAVEL_CO_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_SECOND_COL_INDEX,
    3,
    tevPrefixNonPax,
  );
  const tevNonPaxOther = getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.TRAVEL_OTHER_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_SECOND_COL_INDEX,
    1,
    tevPrefixNonPax,
  );

  return [...tevPax, ...tevNonPax, ...tevNonPaxOther];
}

function honorarium(sheet: Worksheet) {
  const honorariumPrefix = 'Honorarium of';
  return getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.HONORARIUM_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
    2,
    honorariumPrefix,
  );
}

function otherExpenses(sheet: Worksheet) {
  return getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.SUPPLIES_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_COL_INDEX,
    3,
    '',
    MANNER_OF_RELEASE.CASH_ADVANCE,
  );
}

export default function parseActivity(ws: Worksheet) {
  const stDate = ws.getCell(BUDGET_ESTIMATE.START_DATE_CELL).text;
  const month = new Date(stDate).getMonth();
  const bl = boardLodging(ws);
  const tev = travelExpenses(ws);
  const hon = honorarium(ws);
  const others = otherExpenses(ws);

  const info: Activity = {
    // program
    program: ws.getCell(BUDGET_ESTIMATE.PROGRAM_CELL).text,

    // output
    output: ws.getCell(BUDGET_ESTIMATE.OUTPUT_CELL).text,

    // output indicator
    outputIndicator: ws.getCell(BUDGET_ESTIMATE.OUTPUT_INDICATOR_CELL).text,

    // activity
    activityTitle: ws.getCell(BUDGET_ESTIMATE.ACTIVITY_CELL).text,

    // activity indicator
    activityIndicator: ws.getCell(BUDGET_ESTIMATE.ACTIVITY_INDICATOR_CELL).text,

    // month
    month,

    // venue
    venue: ws.getCell(BUDGET_ESTIMATE.VENUE_CELL).text,

    // total pax
    totalPax: extractResult(ws.getCell(BUDGET_ESTIMATE.TOTAL_PAX_CELL).value),

    // output physical target
    outputPhysicalTarget: +ws.getCell(
      BUDGET_ESTIMATE.ACTIVITY_PHYSICAL_TARGET_CELL,
    ).text,

    // activity physical target
    activityPhysicalTarget: +ws.getCell(
      BUDGET_ESTIMATE.ACTIVITY_PHYSICAL_TARGET_CELL,
    ).text,

    // expense items
    expenseItems: [...bl, ...tev, ...hon, ...others],
  };

  return info;
}
