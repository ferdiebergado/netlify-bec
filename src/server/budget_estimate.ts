import { Worksheet } from 'exceljs';
import {
  BUDGET_ESTIMATE,
  EXPENSE_GROUP,
  GAA_OBJECT,
  MANNER_OF_RELEASE,
  VENUES_BY_AIR,
} from './constants';
import {
  Activity,
  ExpenseGroup,
  ExpenseItem,
  GAAObject,
  MannerOfRelease,
} from './types';
import { extractResult, getCellValueAsNumber } from './utils';

/**
 * Creates an ExpenseItem object.
 *
 * @param {ExpenseGroup} expenseGroup - The expense group.
 * @param {GAAObject} gaaObject - The GAA object.
 * @param {string} expenseItem - The expense item description.
 * @param {number} quantity - The quantity of the item.
 * @param {number} freq - The frequency of the item.
 * @param {number} unitCost - The unit cost of the item.
 * @param {MannerOfRelease} mannerOfRelease - The manner of release for the expense.
 * @param {string} tevLocation - The location for TEV (Travel Expense Voucher) if applicable.
 * @param {boolean} ppmp - Indicates if PPMP (Program Project Management Plan) is applicable.
 * @param {boolean} appSupplies - Indicates if the item is applicable for supplies expenses.
 * @param {boolean} appTicket - Indicates if the item is applicable for airfare ticket expenses.
 * @returns {ExpenseItem} The created ExpenseItem object.
 */
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

/**
 * Gets an array of ExpenseItem objects from a worksheet based on provided parameters.
 *
 * @param {Worksheet} sheet - The worksheet containing expense data.
 * @param {number} startRowIndex - The starting row index for reading expense data.
 * @param {number} startColIndex - The starting column index for reading expense data.
 * @param {number} numRows - The number of rows to read.
 * @param {string} prefix - The prefix for creating expense item descriptions.
 * @param {string} venue - The venue information.
 * @param {MannerOfRelease} mannerOfRelease - The manner of release for the expenses.
 * @returns {ExpenseItem[]} An array of ExpenseItem objects.
 */
function getExpenseItems(
  sheet: Worksheet,
  startRowIndex: number,
  startColIndex: number,
  numRows: number,
  prefix: string,
  venue: string = '',
  mannerOfRelease: MannerOfRelease = MANNER_OF_RELEASE.DIRECT_PAYMENT,
): ExpenseItem[] {
  const { QUANTITY_CELL_INDEX, FREQ_CELL_INDEX, UNIT_COST_CELL_INDEX } =
    BUDGET_ESTIMATE;

  const items: ExpenseItem[] = [];

  for (let i = 0; i < numRows; i += 1) {
    const row = sheet.getRow(startRowIndex);

    const quantity = getCellValueAsNumber(
      row.getCell(QUANTITY_CELL_INDEX).text,
    );
    // eslint-disable-next-line no-continue
    if (quantity === 0) continue;

    const item = row.getCell(startColIndex).text;

    let expenseGroup: ExpenseGroup =
      EXPENSE_GROUP.TRAINING_SCHOLARSHIPS_EXPENSES;
    let gaaObject: GAAObject = GAA_OBJECT.TRAINING_EXPENSES;
    let tevLocation = '';
    const ppmp = false;
    let appSupplies = false;
    let appTicket = false;

    if (item.toLowerCase().includes('supplies')) {
      expenseGroup = EXPENSE_GROUP.SUPPLIES_EXPENSES;
      gaaObject = GAA_OBJECT.OTHER_SUPPLIES;
      appSupplies = true;
    }

    const expenseItem = `${prefix} ${item}`;
    const expenseItemLowered = expenseItem.toLowerCase();

    if (
      expenseItemLowered.includes('travel') &&
      expenseItemLowered.includes('participants')
    )
      tevLocation = item;

    const freq = getCellValueAsNumber(row.getCell(FREQ_CELL_INDEX).text || '1');
    const unitCost = parseFloat(row.getCell(UNIT_COST_CELL_INDEX).text);

    if (VENUES_BY_AIR.includes(venue)) appTicket = true;

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

    // eslint-disable-next-line no-param-reassign
    startRowIndex += 1;
  }

  return items;
}

/**
 * Reads and parses board and lodging expenses from a worksheet.
 *
 * @param {Worksheet} sheet - The worksheet containing board and lodging expense data.
 * @returns {ExpenseItem[]} An array of ExpenseItem objects representing board and lodging expenses.
 */
function boardLodging(sheet: Worksheet): ExpenseItem[] {
  const blPrefix = 'Board and Lodging of';
  const { FOR_DOWNLOAD_BOARD, DIRECT_PAYMENT } = MANNER_OF_RELEASE;
  const {
    BOARD_LODGING_DIRECT_PAYMENT_CELL,
    BOARD_LODGING_START_ROW_INDEX,
    EXPENSE_ITEM_FIRST_COL_INDEX,
    BOARD_LODGING_OTHER_ROW_INDEX,
  } = BUDGET_ESTIMATE;

  let blMannerOfRelease: MannerOfRelease = FOR_DOWNLOAD_BOARD;

  if (sheet.getCell(BOARD_LODGING_DIRECT_PAYMENT_CELL).value)
    blMannerOfRelease = DIRECT_PAYMENT;

  const bl = getExpenseItems(
    sheet,
    BOARD_LODGING_START_ROW_INDEX,
    EXPENSE_ITEM_FIRST_COL_INDEX,
    4,
    blPrefix,
    undefined,
    blMannerOfRelease,
  );

  const blOthers = getExpenseItems(
    sheet,
    BOARD_LODGING_OTHER_ROW_INDEX,
    EXPENSE_ITEM_FIRST_COL_INDEX,
    1,
    blPrefix,
    undefined,
    blMannerOfRelease,
  );

  return [...bl, ...blOthers];
}

/**
 * Reads and parses travel expenses from a worksheet based on the provided venue.
 *
 * @param {Worksheet} sheet - The worksheet containing travel expense data.
 * @param {string} venue - The venue information.
 * @returns {ExpenseItem[]} An array of ExpenseItem objects representing travel expenses.
 */
function travelExpenses(sheet: Worksheet, venue: string): ExpenseItem[] {
  const tevPrefix = 'Travel Expenses of Participants from';
  const tevMannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_PSF;
  const {
    TRAVEL_REGION_ROW_INDEX,
    EXPENSE_ITEM_SECOND_COL_INDEX,
    TRAVEL_CO_ROW_INDEX,
    TRAVEL_OTHER_ROW_INDEX,
  } = BUDGET_ESTIMATE;
  const tevPax = getExpenseItems(
    sheet,
    TRAVEL_REGION_ROW_INDEX,
    EXPENSE_ITEM_SECOND_COL_INDEX,
    18,
    tevPrefix,
    undefined,
    tevMannerOfRelease,
  );

  const tevPrefixNonPax = 'Travel Expenses of';
  const tevNonPax = getExpenseItems(
    sheet,
    TRAVEL_CO_ROW_INDEX,
    EXPENSE_ITEM_SECOND_COL_INDEX,
    3,
    tevPrefixNonPax,
    venue,
  );

  const tevNonPaxOther = getExpenseItems(
    sheet,
    TRAVEL_OTHER_ROW_INDEX,
    EXPENSE_ITEM_SECOND_COL_INDEX,
    1,
    tevPrefixNonPax,
    venue,
  );

  return [...tevPax, ...tevNonPax, ...tevNonPaxOther];
}

/**
 * Reads and parses honorarium expenses from a worksheet.
 *
 * @param {Worksheet} sheet - The worksheet containing honorarium expense data.
 * @returns {ExpenseItem[]} An array of ExpenseItem objects representing honorarium expenses.
 */
function honorarium(sheet: Worksheet): ExpenseItem[] {
  const honorariumPrefix = 'Honorarium of';
  return getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.HONORARIUM_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
    2,
    honorariumPrefix,
  );
}

/**
 * Reads and parses other expenses from a worksheet.
 *
 * @param {Worksheet} sheet - The worksheet containing other expense data.
 * @returns {ExpenseItem[]} An array of ExpenseItem objects representing other expenses.
 */
function otherExpenses(sheet: Worksheet): ExpenseItem[] {
  return getExpenseItems(
    sheet,
    BUDGET_ESTIMATE.SUPPLIES_ROW_INDEX,
    BUDGET_ESTIMATE.EXPENSE_ITEM_COL_INDEX,
    3,
    '',
    undefined,
    MANNER_OF_RELEASE.CASH_ADVANCE,
  );
}

/**
 * Parses activity information from a worksheet, including related expense items.
 *
 * @param {Worksheet} ws - The worksheet containing activity information.
 * @returns {Activity} An object representing parsed activity information.
 */
export default function parseActivity(ws: Worksheet): Activity {
  const {
    VENUE_CELL,
    START_DATE_CELL,
    PROGRAM_CELL,
    OUTPUT_CELL,
    OUTPUT_INDICATOR_CELL,
    ACTIVITY_CELL,
    ACTIVITY_INDICATOR_CELL,
    TOTAL_PAX_CELL,
    OUTPUT_PHYSICAL_TARGET_CELL,
    ACTIVITY_PHYSICAL_TARGET_CELL,
  } = BUDGET_ESTIMATE;
  const venue = ws.getCell(VENUE_CELL).text;
  const stDate = ws.getCell(START_DATE_CELL).text;
  const month = new Date(stDate).getMonth();
  const bl = boardLodging(ws);
  const tev = travelExpenses(ws, venue);
  const hon = honorarium(ws);
  const others = otherExpenses(ws);

  const info: Activity = {
    // program
    program: ws.getCell(PROGRAM_CELL).text,

    // output
    output: ws.getCell(OUTPUT_CELL).text,

    // output indicator
    outputIndicator: ws.getCell(OUTPUT_INDICATOR_CELL).text,

    // activity
    activityTitle: ws.getCell(ACTIVITY_CELL).text,

    // activity indicator
    activityIndicator: ws.getCell(ACTIVITY_INDICATOR_CELL).text,

    // month
    month,

    // venue
    venue,

    // total pax
    totalPax: extractResult(ws.getCell(TOTAL_PAX_CELL).value),

    // output physical target
    outputPhysicalTarget: +ws.getCell(OUTPUT_PHYSICAL_TARGET_CELL).text,

    // activity physical target
    activityPhysicalTarget: +ws.getCell(ACTIVITY_PHYSICAL_TARGET_CELL).text,

    // expense items
    expenseItems: [...bl, ...tev, ...hon, ...others],
  };

  return info;
}
