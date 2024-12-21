import { Workbook } from './workbook.js';
import type {
  ExpenseItem,
  ExpenseOptions,
  ActivityInfo,
  Activity,
  SheetConfig,
} from './types/globals.js';
import {
  AUXILLIARY_SHEETS,
  BUDGET_ESTIMATE,
  ExpenseGroup,
  GAAObject,
  HONORARIUM_EXPENSE_PREFIX,
  ReleaseManner,
  TRAVEL_EXPENSE_PREFIX,
  VENUES_BY_AIR,
} from './constants.js';
import { getCellValueAsNumber } from './utils.js';
import { BudgetEstimateParseError } from './parseError.js';

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
   * @private
   * @param {SheetConfig} sheetConfig The config for the sheet for reading expense data.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects.
   */
  private _getExpenseItems({
    startRowIndex,
    startColIndex,
    numRows,
    options,
  }: SheetConfig): ExpenseItem[] {
    const { QUANTITY_CELL_INDEX, FREQ_CELL_INDEX, UNIT_COST_CELL_INDEX } =
      BUDGET_ESTIMATE;
    const { prefix, releaseManner, venue, hasPPMP } = options;
    const sheet = this.getActiveSheet();

    return Array.from({ length: numRows }, (_, i) => {
      const rowIndex = startRowIndex + i;
      const row = sheet.getRow(rowIndex);

      const unitCost = Number.parseFloat(
        row.getCell(UNIT_COST_CELL_INDEX).text,
      );
      const quantity = getCellValueAsNumber(row.getCell(QUANTITY_CELL_INDEX));
      const freq = getCellValueAsNumber(row.getCell(FREQ_CELL_INDEX)) || 1;

      if (quantity > 0 && freq > 0) {
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

        if (venue && VENUES_BY_AIR.includes(venue)) hasAPPTicket = true;

        return {
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
        } as ExpenseItem;
      }

      return null;
    }).filter(Boolean) as ExpenseItem[];
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
      // TOTAL_PAX_CELL,
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

    console.log('startDate:', startDate);

    if (!startDate) {
      throw new BudgetEstimateParseError(
        'Please provide the "From" and "To" Date of the activity.',
        {
          file: this.activeFile as string,
          sheet: sheet.name,
          activity: activityTitle,
        },
      );
    }

    const month = new Date(startDate).getMonth();
    // const totalPax = extractResult(sheet.getCell(TOTAL_PAX_CELL).value);

    const info: ActivityInfo = {
      program,
      output,
      outputIndicator,
      activityTitle,
      activityIndicator,
      month,
      venue,
      // totalPax,
      outputPhysicalTarget,
      activityPhysicalTarget,
    };

    return info;
  }

  /**
   * Parses activity information and organizes it into an Activity object.
   *
   * @private
   *
   * @returns {Activity | undefined} An object representing the parsed activity information,
   * or undefined if activity information is not available.
   */
  private _parseActivity(): Activity | undefined {
    /**
     * Gets detailed information about the activity.
     *
     * @returns {ActivityInfo | undefined} An object containing parsed information about the activity,
     * or undefined if information is not available.
     */
    const info = this.getActivityInfo();

    if (info) {
      /**
       * Gets information about board and lodging expenses.
       *
       * @returns {ExpenseItem[]} An array of expense items related to board and lodging.
       */
      const lodging = this.getBoardAndLodging();

      /**
       * Gets information about travel expenses based on the activity's venue.
       *
       * @param {string} venue - The venue of the activity.
       * @returns {ExpenseItem[]} An array of expense items related to travel.
       */
      const tev = this.getTravelExpenses(info.venue);

      /**
       * Gets information about travel expenses Program Support Fund (tevPSF).
       *
       * @returns {ExpenseItem[]} An array of expense items related to tevPSF.
       */
      const tevPSF = this.getTevPSF();

      /**
       * Gets information about honorarium expenses.
       *
       * @returns {ExpenseItem[]} An array of expense items related to honorarium.
       */
      const honorarium = this.getHonorarium();

      /**
       * Gets information about other miscellaneous expenses.
       *
       * @returns {ExpenseItem[]} An array of other expense items.
       */
      const otherExpenses = this.getOtherExpenses();

      /**
       * Combines different types of expenses into a single array.
       *
       * @type {ExpenseItem[]}
       */
      const expenseItems: ExpenseItem[] = [
        ...lodging,
        ...tev,
        ...honorarium,
        ...otherExpenses,
      ];

      /**
       * Represents the finalized activity object with parsed information.
       *
       * @type {Activity}
       */
      const activity: Activity = {
        info,
        tevPSF,
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
    return this.wb.worksheets
      .filter(sheet => {
        const { name } = sheet;

        // Skip auxiliary sheets
        if (AUXILLIARY_SHEETS.includes(name)) return false;

        // Skip sheets without the required program heading
        if (
          sheet.getCell(BUDGET_ESTIMATE.PROGRAM_HEADING_CELL).text !==
          'PROGRAM:'
        ) {
          return false;
        }

        return true;
      })
      .map(sheet => {
        console.log('parsing sheet', sheet.name);
        this.ws = sheet;

        try {
          return this._parseActivity();
        } catch (error) {
          if (error instanceof BudgetEstimateParseError) {
            throw error;
          } else {
            console.error(error);
            throw new BudgetEstimateParseError(
              'Please check the layout and details of the activity in the following sheet:',
              {
                file: this.activeFile as string,
                sheet: sheet.name,
              },
            );
          }
        }
      })
      .filter(activity => activity !== undefined);
  }

  /**
   * Reads and parses board and lodging expenses from the budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing board and lodging expenses.
   */
  getBoardAndLodging(): ExpenseItem[] {
    const { DIRECT_PAYMENT } = ReleaseManner;
    const {
      NUM_DAYS_CELL,
      BOARD_LODGING_TOTAL_PAX_CELL,
      BOARD_LODGING_UNIT_COST_CELL,
    } = BUDGET_ESTIMATE;
    const sheet = this.getActiveSheet();
    const quantity = getCellValueAsNumber(
      sheet.getCell(BOARD_LODGING_TOTAL_PAX_CELL),
    );
    const unitCost = getCellValueAsNumber(
      sheet.getCell(BOARD_LODGING_UNIT_COST_CELL),
    );
    const freq = getCellValueAsNumber(sheet.getCell(NUM_DAYS_CELL));

    const expenseItem: ExpenseItem = {
      expenseGroup: ExpenseGroup.TRAINING_SCHOLARSHIPS_EXPENSES,
      gaaObject: GAAObject.TRAINING_EXPENSES,
      expenseItem: 'Board and Lodging of Participants',
      quantity,
      unitCost,
      freq,
      hasPPMP: true,
      releaseManner: DIRECT_PAYMENT,
    };

    return [expenseItem];
  }

  /**
   * Reads and parses travel expenses of participants.
   *
   * @returns {ExpenseItem[]}
   */
  getTevPSF(): ExpenseItem[] {
    const basePrefix = TRAVEL_EXPENSE_PREFIX;
    const prefixPax = `${basePrefix} Participants from`;
    const releaseManner = ReleaseManner.FOR_DOWNLOAD_PSF;
    const { TRAVEL_REGION_ROW_INDEX, EXPENSE_ITEM_SECOND_COL_INDEX } =
      BUDGET_ESTIMATE;

    const options: ExpenseOptions = {
      prefix: prefixPax,
      releaseManner,
      hasPPMP: false,
    };

    const tevConfig: SheetConfig = {
      startRowIndex: TRAVEL_REGION_ROW_INDEX,
      startColIndex: EXPENSE_ITEM_SECOND_COL_INDEX,
      numRows: 18,
      options,
    };

    const tevPax = this._getExpenseItems(tevConfig);

    return tevPax;
  }

  /**
   * Reads and parses travel expenses from a budget estimate based on the provided venue.
   *
   * @param {string} venue The venue of the activity.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing travel expenses.
   */
  getTravelExpenses(venue: string): ExpenseItem[] {
    const basePrefix = TRAVEL_EXPENSE_PREFIX;
    const {
      EXPENSE_ITEM_SECOND_COL_INDEX,
      TRAVEL_CO_ROW_INDEX,
      TRAVEL_OTHER_ROW_INDEX,
    } = BUDGET_ESTIMATE;

    const options = {
      prefix: basePrefix,
      releaseManner: ReleaseManner.DIRECT_PAYMENT,
      venue,
    };

    const tevNonPaxConfig: SheetConfig = {
      startRowIndex: TRAVEL_CO_ROW_INDEX,
      startColIndex: EXPENSE_ITEM_SECOND_COL_INDEX,
      numRows: 3,
      options,
    };
    const tevNonPax = this._getExpenseItems(tevNonPaxConfig);

    const tevNonPaxOtherConfig: SheetConfig = {
      startRowIndex: TRAVEL_OTHER_ROW_INDEX,
      startColIndex: EXPENSE_ITEM_SECOND_COL_INDEX,
      numRows: 1,
      options,
    };
    const tevNonPaxOther = this._getExpenseItems(tevNonPaxOtherConfig);

    return [...tevNonPax, ...tevNonPaxOther];
  }

  /**
   * Reads and parses honorarium expenses from a budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing honorarium expenses.
   */
  getHonorarium(): ExpenseItem[] {
    const prefix = HONORARIUM_EXPENSE_PREFIX;
    const releaseManner = ReleaseManner.DIRECT_PAYMENT;
    const options: ExpenseOptions = {
      prefix,
      releaseManner,
      hasPPMP: false,
    };

    const honorariumConfig: SheetConfig = {
      startRowIndex: BUDGET_ESTIMATE.HONORARIUM_ROW_INDEX,
      startColIndex: BUDGET_ESTIMATE.EXPENSE_ITEM_FIRST_COL_INDEX,
      numRows: 2,
      options,
    };

    const honorarium = this._getExpenseItems(honorariumConfig);

    return honorarium;
  }

  /**
   * Reads and parses other expenses from a budget estimate.
   *
   * @returns {ExpenseItem[]} An array of ExpenseItem objects representing other expenses.
   */
  getOtherExpenses(): ExpenseItem[] {
    console.log('getting other expenses...');

    const options: ExpenseOptions = {
      prefix: '',
      releaseManner: ReleaseManner.CASH_ADVANCE,
      hasPPMP: false,
    };

    const otherExpensesConfig: SheetConfig = {
      startRowIndex: BUDGET_ESTIMATE.MEAL_EXPENSES_ROW_INDEX,
      startColIndex: BUDGET_ESTIMATE.EXPENSE_ITEM_COL_INDEX,
      numRows: 3,
      options,
    };

    const otherExpenses = this._getExpenseItems(otherExpensesConfig);

    console.log('other expenses:', otherExpenses);

    return otherExpenses;
  }
}
