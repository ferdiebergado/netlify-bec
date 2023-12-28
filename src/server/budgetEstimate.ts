import { Workbook } from './workbook';
import type {
  ExpenseItem,
  ExpenseOptions,
  ActivityInfo,
  Activity,
} from '../types/globals';
import {
  AUXILLIARY_SHEETS,
  BUDGET_ESTIMATE,
  ExpenseGroup,
  GAAObject,
  ReleaseManner,
  VENUES_BY_AIR,
} from './constants';
import { extractResult, getCellValueAsNumber } from './utils';

export class BudgetEstimate extends Workbook<BudgetEstimate> {
  constructor() {
    super();
  }

  protected createInstance(): BudgetEstimate {
    return this;
  }

  /**
   * Gets an array of ExpenseItem objects from a budget estimate based on provided parameters.
   *
   * @param {number} startRowIndex - The starting row index for reading expense data.
   * @param {number} startColIndex - The starting column index for reading expense data.
   * @param {number} numRows - The number of rows to read.
   * @param {ExpenseOptions} data - The data to be stored to the expense item.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects.
   */
  getExpenseItems(
    startRowIndex: number,
    startColIndex: number,
    numRows: number,
    data: ExpenseOptions,
  ): ExpenseItem[] {
    const expenseItems: ExpenseItem[] = [];

    if (this.ws) {
      const { QUANTITY_CELL_INDEX, FREQ_CELL_INDEX, UNIT_COST_CELL_INDEX } =
        BUDGET_ESTIMATE;
      const { prefix, releaseManner, venue, hasPPMP } = data;

      for (let i = 0; i < numRows; i += 1) {
        const row = this.ws.getRow(startRowIndex);

        const quantity = getCellValueAsNumber(
          row.getCell(QUANTITY_CELL_INDEX).text,
        );
        // eslint-disable-next-line no-continue
        if (quantity === 0) continue;

        const item = row.getCell(startColIndex).text;

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
        const unitCost = parseFloat(row.getCell(UNIT_COST_CELL_INDEX).text);

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

        // eslint-disable-next-line no-param-reassign
        startRowIndex += 1;
      }
    }

    return expenseItems;
  }

  /**
   * Parses activity information from the budget estimate
   *
   * @returns {ActivityInfo} An object representing parsed activity information.
   */
  getActivityInfo(): ActivityInfo | undefined {
    if (this.ws) {
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

      const program = this.ws.getCell(PROGRAM_CELL).text;
      const output = this.ws.getCell(OUTPUT_CELL).text;
      const outputIndicator = this.ws.getCell(OUTPUT_INDICATOR_CELL).text;
      const activityTitle = this.ws.getCell(ACTIVITY_CELL).text;
      const activityIndicator = this.ws.getCell(ACTIVITY_INDICATOR_CELL).text;
      const outputPhysicalTarget = Number.parseInt(
        this.ws.getCell(OUTPUT_PHYSICAL_TARGET_CELL).text,
      );
      const activityPhysicalTarget = Number.parseInt(
        this.ws.getCell(ACTIVITY_PHYSICAL_TARGET_CELL).text,
      );
      const venue = this.ws.getCell(VENUE_CELL).text;
      const startDate = this.ws.getCell(START_DATE_CELL).text;
      const month = new Date(startDate).getMonth();
      const totalPax = extractResult(this.ws.getCell(TOTAL_PAX_CELL).value);

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

    this.wb?.eachSheet(sheet => {
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
    if (this.ws) {
      const prefix = 'Board and Lodging of';
      const { FOR_DOWNLOAD_BOARD, DIRECT_PAYMENT } = ReleaseManner;
      const {
        BOARD_LODGING_DIRECT_PAYMENT_CELL,
        BOARD_LODGING_START_ROW_INDEX,
        EXPENSE_ITEM_FIRST_COL_INDEX,
        BOARD_LODGING_OTHER_ROW_INDEX,
      } = BUDGET_ESTIMATE;

      let releaseManner: ReleaseManner = FOR_DOWNLOAD_BOARD;
      let hasPPMP = false;

      if (this.ws.getCell(BOARD_LODGING_DIRECT_PAYMENT_CELL).value) {
        releaseManner = DIRECT_PAYMENT;
        hasPPMP = true;
      }

      const expenseData: ExpenseOptions = {
        prefix,
        releaseManner,
        hasPPMP,
      };

      const lodging = this.getExpenseItems(
        BOARD_LODGING_START_ROW_INDEX,
        EXPENSE_ITEM_FIRST_COL_INDEX,
        4,
        expenseData,
      );

      const lodgingOthers = this.getExpenseItems(
        BOARD_LODGING_OTHER_ROW_INDEX,
        EXPENSE_ITEM_FIRST_COL_INDEX,
        1,
        expenseData,
      );

      return [...lodging, ...lodgingOthers];
    }

    return [];
  }

  /**
   * Reads and parses travel expenses from a worksheet based on the provided venue.
   *
   * @param {string} venue - The venue of the activity.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing travel expenses.
   */
  getTravelExpenses(venue: string): ExpenseItem[] {
    const basePrefix = 'Travel Expenses of';
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
    };

    const tevPax = this.getExpenseItems(
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

    const tevNonPax = this.getExpenseItems(
      TRAVEL_CO_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      3,
      expenseData,
    );

    const tevNonPaxOther = this.getExpenseItems(
      TRAVEL_OTHER_ROW_INDEX,
      EXPENSE_ITEM_SECOND_COL_INDEX,
      1,
      expenseData,
    );

    return [...tevPax, ...tevNonPax, ...tevNonPaxOther];
  }

  /**
   * Reads and parses honorarium expenses from a worksheet.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing honorarium expenses.
   */
  getHonorarium(): ExpenseItem[] {
    const prefix = 'Honorarium of';
    const releaseManner = ReleaseManner.DIRECT_PAYMENT;
    const expenseData: ExpenseOptions = {
      prefix,
      releaseManner,
    };
    const honorarium = this.getExpenseItems(
      BUDGET_ESTIMATE.HONORARIUM_ROW_INDEX,
      BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
      2,
      expenseData,
    );

    return honorarium;
  }

  /**
   * Reads and parses other expenses from a worksheet.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing other expenses.
   */
  getOtherExpenses(): ExpenseItem[] {
    const expenseData: ExpenseOptions = {
      prefix: '',
      releaseManner: ReleaseManner.CASH_ADVANCE,
    };

    const otherExpenses = this.getExpenseItems(
      BUDGET_ESTIMATE.SUPPLIES_ROW_INDEX,
      BUDGET_ESTIMATE.EXPENSE_ITEM_COL_INDEX,
      3,
      expenseData,
    );

    return otherExpenses;
  }
}
