import { Workbook } from './workbook';
import type {
  ExpenseItem,
  ExpenseOptions,
  ActivityInfo,
  Activity,
} from './types/globals';
import {
  AUXILLIARY_SHEETS,
  BOARD_LODGING_EXPENSE_PREFIX,
  BUDGET_ESTIMATE,
  ExpenseGroup,
  GAAObject,
  HONORARIUM_EXPENSE_PREFIX,
  ReleaseManner,
  TRAVEL_EXPENSE_PREFIX,
  VENUES_BY_AIR,
} from './constants';
import { extractResult, getCellValueAsNumber } from './utils';
import type { Worksheet } from 'exceljs';

/**
 * Represents a specialized workbook for managing budget estimates.
 *
 * @class BudgetEstimate
 * @extends {Workbook<BudgetEstimate>}
 */
export class BudgetEstimate extends Workbook<BudgetEstimate> {
  /**
   * Creates an instance of the BudgetEstimate class.
   *
   * @public
   * @constructor
   */
  constructor() {
    super();
  }

  /**
   * Overrides the abstract method in the base class to create an instance of BudgetEstimate.
   *
   * @protected
   * @returns {BudgetEstimate} The created instance of ExpenditureMatrix.
   */

  protected createInstance(): BudgetEstimate {
    return this;
  }

  /**
   * Gets an array of ExpenseItem objects from a budget estimate based on provided parameters.
   *
   * @param {number} startRowIndex The starting row index for reading expense data.
   * @param {number} startColIndex The starting column index for reading expense data.
   * @param {number} numRows The number of rows to read.
   * @param {ExpenseOptions} options The data to be stored to the expense item.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects.
   */
  static getExpenseItems(
    sheet: Worksheet,
    startRowIndex: number,
    startColIndex: number,
    numRows: number,
    options: ExpenseOptions,
  ): ExpenseItem[] {
    const expenseItems: ExpenseItem[] = [];
    const { QUANTITY_CELL_INDEX, FREQ_CELL_INDEX, UNIT_COST_CELL_INDEX } =
      BUDGET_ESTIMATE;
    const { prefix, releaseManner, venue, hasPPMP } = options;
    let currentRowIndex = startRowIndex;

    for (let i = 0; i < numRows; i += 1) {
      const row = sheet.getRow(currentRowIndex);

      currentRowIndex += 1;

      const unitCost = Number.parseFloat(
        row.getCell(UNIT_COST_CELL_INDEX).text,
      );
      const quantity = getCellValueAsNumber(
        row.getCell(QUANTITY_CELL_INDEX).text,
      );
      // eslint-disable-next-line no-continue
      if (quantity === 0 || unitCost === 0) continue;

      const item = row.getCell(startColIndex).text.trim();

      let expenseGroup: ExpenseGroup =
        ExpenseGroup.TRAINING_SCHOLARSHIPS_EXPENSES;
      let gaaObject: GAAObject = GAAObject.TRAINING_EXPENSES;
      let tevLocation = '';
      let hasAPPSupplies = false;
      let hasAPPTicket = false;

      if (item.toLowerCase().includes('supplies')) {
        expenseGroup = ExpenseGroup.SUPPLIES_EXPENSES;
        gaaObject = GAAObject.OTHER_SUPPLIES;
        hasAPPSupplies = true;
      }

      const expenseItem = `${prefix} ${item}`;
      const expenseItemLowered = expenseItem.toLowerCase();

      if (
        expenseItemLowered.includes('travel') &&
        expenseItemLowered.includes('participants')
      )
        tevLocation = item;

      const freq = getCellValueAsNumber(
        row.getCell(FREQ_CELL_INDEX).text || '1',
      );

      if (venue && VENUES_BY_AIR.includes(venue)) hasAPPTicket = true;

      const newExpenseItem: ExpenseItem = {
        expenseGroup,
        gaaObject,
        expenseItem,
        quantity,
        freq,
        unitCost,
        releaseManner,
        tevLocation,
        hasPPMP,
        hasAPPSupplies,
        hasAPPTicket,
      };

      expenseItems.push(newExpenseItem);
    }

    return expenseItems;
  }

  /**
   * Parses activity information from the budget estimate.
   *
   * @returns {ActivityInfo} An object representing parsed activity information.
   */
  getActivityInfo(): ActivityInfo | undefined {
    const sheet = this.getActiveSheet();

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

    const program = sheet.getCell(PROGRAM_CELL).text;
    const output = sheet.getCell(OUTPUT_CELL).text;
    const outputIndicator = sheet.getCell(OUTPUT_INDICATOR_CELL).text;
    const activityTitle = sheet.getCell(ACTIVITY_CELL).text;
    const activityIndicator = sheet.getCell(ACTIVITY_INDICATOR_CELL).text;
    const outputPhysicalTarget = Number.parseInt(
      sheet.getCell(OUTPUT_PHYSICAL_TARGET_CELL).text,
    );
    const activityPhysicalTarget = Number.parseInt(
      sheet.getCell(ACTIVITY_PHYSICAL_TARGET_CELL).text,
    );
    const venue = sheet.getCell(VENUE_CELL).text;
    const startDate = sheet.getCell(START_DATE_CELL).text;
    const month = new Date(startDate).getMonth();
    const totalPax = extractResult(sheet.getCell(TOTAL_PAX_CELL).value);

    const info: ActivityInfo = {
      program,
      output,
      outputIndicator,
      activityTitle,
      activityIndicator,
      month,
      venue,
      totalPax,
      outputPhysicalTarget,
      activityPhysicalTarget,
    };

    return info;
  }

  private _parseActivity() {
    const info = this.getActivityInfo();

    if (info) {
      const lodging = this.getBoardAndLodging();
      const tev = this.getTravelExpenses(info.venue);
      const honorarium = this.getHonorarium();
      const otherExpenses = this.getOtherExpenses();
      const expenseItems: ExpenseItem[] = [
        ...lodging,
        ...tev,
        ...honorarium,
        ...otherExpenses,
      ];
      const activity: Activity = {
        info,
        expenseItems,
      };

      return activity;
    }
  }

  /**
   * Extracts the activities from the sheets of the Budget Estimate
   *
   * @returns Activity[]
   */
  getActivities(): Activity[] {
    const activities: Activity[] = [];

    this.wb.eachSheet(sheet => {
      const { name } = sheet;

      if (AUXILLIARY_SHEETS.includes(name)) {
        // eslint-disable-next-line no-console
        // console.log('skipping', name);
        return;
      }

      // TODO: add more columns to check
      if (
        sheet.getCell(BUDGET_ESTIMATE.PROGRAM_HEADING_CELL).text !== 'PROGRAM:'
      ) {
        // eslint-disable-next-line no-console
        // console.log('skipping non budget estimate sheet:', name);
        return;
      }

      // eslint-disable-next-line no-console
      // console.log('processing sheet:', name);

      this.ws = sheet;

      const activity = this._parseActivity();

      if (activity) activities.push(activity);
    });

    return activities;
  }

  /**
   * Reads and parses board and lodging expenses from the budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing board and lodging expenses.
   */
  getBoardAndLodging(): ExpenseItem[] {
    const sheet = this.getActiveSheet();

    const prefix = BOARD_LODGING_EXPENSE_PREFIX;
    const { FOR_DOWNLOAD_BOARD, DIRECT_PAYMENT } = ReleaseManner;
    const {
      BOARD_LODGING_DIRECT_PAYMENT_CELL,
      BOARD_LODGING_START_ROW_INDEX,
      EXPENSE_ITEM_FIRST_COL_INDEX,
      BOARD_LODGING_OTHER_ROW_INDEX,
    } = BUDGET_ESTIMATE;

    let releaseManner: ReleaseManner = FOR_DOWNLOAD_BOARD;
    let hasPPMP = false;

    if (sheet.getCell(BOARD_LODGING_DIRECT_PAYMENT_CELL).value) {
      releaseManner = DIRECT_PAYMENT;
      hasPPMP = true;
    }

    const expenseData: ExpenseOptions = {
      prefix,
      releaseManner,
      hasPPMP,
    };

    const lodging = BudgetEstimate.getExpenseItems(
      sheet,
      BOARD_LODGING_START_ROW_INDEX,
      EXPENSE_ITEM_FIRST_COL_INDEX,
      4,
      expenseData,
    );

    const lodgingOthers = BudgetEstimate.getExpenseItems(
      sheet,
      BOARD_LODGING_OTHER_ROW_INDEX,
      EXPENSE_ITEM_FIRST_COL_INDEX,
      1,
      expenseData,
    );

    return [...lodging, ...lodgingOthers];
  }

  /**
   * Reads and parses travel expenses from a budget estimate based on the provided venue.
   *
   * @param {string} venue The venue of the activity.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing travel expenses.
   */
  getTravelExpenses(venue: string): ExpenseItem[] {
    const sheet = this.getActiveSheet();

    const basePrefix = TRAVEL_EXPENSE_PREFIX;
    const prefixPax = `${basePrefix} Participants from`;
    const releaseManner = ReleaseManner.FOR_DOWNLOAD_PSF;
    const {
      TRAVEL_REGION_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      TRAVEL_CO_ROW_INDEX,
      TRAVEL_OTHER_ROW_INDEX,
    } = BUDGET_ESTIMATE;

    let expenseData: ExpenseOptions = {
      prefix: prefixPax,
      releaseManner,
      hasPPMP: false,
    };

    const tevPax = BudgetEstimate.getExpenseItems(
      sheet,
      TRAVEL_REGION_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      18,
      expenseData,
    );

    expenseData = {
      prefix: basePrefix,
      releaseManner: ReleaseManner.DIRECT_PAYMENT,
      venue,
    };

    const tevNonPax = BudgetEstimate.getExpenseItems(
      sheet,
      TRAVEL_CO_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      3,
      expenseData,
    );

    const tevNonPaxOther = BudgetEstimate.getExpenseItems(
      sheet,
      TRAVEL_OTHER_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      1,
      expenseData,
    );

    return [...tevPax, ...tevNonPax, ...tevNonPaxOther];
  }

  /**
   * Reads and parses honorarium expenses from a budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing honorarium expenses.
   */
  getHonorarium(): ExpenseItem[] {
    const sheet = this.getActiveSheet();

    const prefix = HONORARIUM_EXPENSE_PREFIX;
    const releaseManner = ReleaseManner.DIRECT_PAYMENT;
    const expenseData: ExpenseOptions = {
      prefix,
      releaseManner,
      hasPPMP: false,
    };
    const honorarium = BudgetEstimate.getExpenseItems(
      sheet,
      BUDGET_ESTIMATE.HONORARIUM_ROW_INDEX,
      BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
      2,
      expenseData,
    );

    return honorarium;
  }

  /**
   * Reads and parses other expenses from a budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing other expenses.
   */
  getOtherExpenses(): ExpenseItem[] {
    const sheet = this.getActiveSheet();
    const expenseData: ExpenseOptions = {
      prefix: '',
      releaseManner: ReleaseManner.CASH_ADVANCE,
      hasPPMP: false,
    };

    const otherExpenses = BudgetEstimate.getExpenseItems(
      sheet,
      BUDGET_ESTIMATE.MEAL_EXPENSES_ROW_INDEX,
      BUDGET_ESTIMATE.EXPENSE_ITEM_COL_INDEX,
      3,
      expenseData,
    );

    return otherExpenses;
  }
}
