import Worksheet from './worksheet';
import type {
  ExpenseItem,
  YesNo,
  ActivityInfo,
  ExpenseGroup,
  GAAObject,
  MannerOfRelease,
} from './types';
import { extractResult } from './utils';
import {
  BUDGET_ESTIMATE,
  EXPENSE_PREFIX,
  EXPENSE_GROUP,
  GAA_OBJECT,
  MANNER_OF_RELEASE,
} from './constants';

class BudgetEstimate extends Worksheet {
  // eslint-disable-next-line @typescript-eslint/no-useless-constructor
  constructor(xls: ArrayBuffer, sheet: string = BUDGET_ESTIMATE.SHEET_NAME) {
    super(xls, sheet);
  }

  get activityInfo() {
    if (!this.ws) throw new Error(BudgetEstimate.LOAD_ERROR_MSG);

    const stDate = this.ws.getCell(BUDGET_ESTIMATE.CELL_START_DATE).text;
    const month = new Date(stDate).getMonth();

    const info: ActivityInfo = {
      // program
      program: this.ws.getCell(BUDGET_ESTIMATE.CELL_PROGRAM).text,

      // output
      output: this.ws.getCell(BUDGET_ESTIMATE.CELL_OUTPUT).text,

      // output indicator
      outputIndicator: this.ws.getCell(BUDGET_ESTIMATE.CELL_OUTPUT_INDICATOR)
        .text,

      // activity
      activity: this.ws.getCell(BUDGET_ESTIMATE.CELL_ACTIVITY).text,

      // activity indicator
      activityIndicator: this.ws.getCell(
        BUDGET_ESTIMATE.CELL_ACTIVITY_INDICATOR,
      ).text,

      // month
      month,

      // venue
      venue: this.ws.getCell(BUDGET_ESTIMATE.CELL_VENUE).text,

      // total pax
      totalPax: extractResult(
        this.ws.getCell(BUDGET_ESTIMATE.CELL_TOTAL_PAX).value,
      ),
    };

    return info;
  }

  get expenseItems() {
    return [
      ...this.boardLodging(),
      ...this.travelExpenses(),
      ...this.honorarium(),
      ...this.suppliesContingency(),
    ];
  }

  boardLodging() {
    const prefix = EXPENSE_PREFIX.BOARD_LODGING;
    const col = BUDGET_ESTIMATE.COL_BOARD_LODGING;
    const boardLodgingPax = this._collectExpenseItems(
      col,
      BUDGET_ESTIMATE.ROW_BOARD_LODGING_START,
      BUDGET_ESTIMATE.ROW_BOARD_LODGING_END,
      prefix,
    );

    const boardLodgingOther = this._collectExpenseItems(
      col,
      BUDGET_ESTIMATE.ROW_BOARD_LODGING_OTHER,
      BUDGET_ESTIMATE.ROW_BOARD_LODGING_OTHER,
      prefix,
    );

    return [...boardLodgingPax, ...boardLodgingOther];
  }

  travelExpenses() {
    const prefix = EXPENSE_PREFIX.TRAVEL;

    const travelRegion = this._collectExpenseItems(
      BUDGET_ESTIMATE.COL_TRAVEL_REGION,
      BUDGET_ESTIMATE.ROW_TRAVEL_REGION_START,
      BUDGET_ESTIMATE.ROW_TRAVEL_REGION_END,
      prefix,
    );

    const travelCO = this._collectExpenseItems(
      BUDGET_ESTIMATE.COL_TRAVEL_CO,
      BUDGET_ESTIMATE.ROW_TRAVEL_CO_START,
      BUDGET_ESTIMATE.ROW_TRAVEL_CO_END,
      prefix,
    );

    const travelOther = this._collectExpenseItems(
      BUDGET_ESTIMATE.COL_TRAVEL_OTHER,
      BUDGET_ESTIMATE.ROW_TRAVEL_OTHER,
      BUDGET_ESTIMATE.ROW_TRAVEL_OTHER,
      prefix,
    );

    return [...travelRegion, ...travelCO, ...travelOther];
  }

  honorarium() {
    return this._collectExpenseItems(
      BUDGET_ESTIMATE.COL_HONORARIUM,
      BUDGET_ESTIMATE.ROW_HONORARIUM_START,
      BUDGET_ESTIMATE.ROW_HONORARIUM_END,
      EXPENSE_PREFIX.HONORARIUM,
    );
  }

  suppliesContingency() {
    return this._collectExpenseItems(
      BUDGET_ESTIMATE.COL_SUPPLIES_CONTINGENCY,
      BUDGET_ESTIMATE.ROW_SUPPLIES_CONTINGENCY_START,
      BUDGET_ESTIMATE.ROW_SUPPLIES_CONTINGENCY_END,
      '',
    );
  }

  _collectExpenseItems(
    col: string,
    startRow: number,
    endRow: number,
    itemPrefix: string,
  ) {
    const expenseItems = [];

    for (let row = startRow; row <= endRow; row++) {
      const expense = this._parseBudgetEstimateRow(col, row, itemPrefix);
      if (expense) expenseItems.push(expense);
    }

    return expenseItems;
  }

  _parseBudgetEstimateRow(
    itemCol: string,
    row: string | number,
    itemPrefix: string,
  ) {
    if (!this.ws) throw new Error(BudgetEstimate.LOAD_ERROR_MSG);

    const unitCost = extractResult(
      this.ws.getCell(BUDGET_ESTIMATE.COL_AMOUNT + row).value,
    );

    if (!unitCost || unitCost === 0) return;

    const expenseItem = itemPrefix + this.ws.getCell(itemCol + row).text.trim();
    const quantity = extractResult(
      this.ws.getCell(BUDGET_ESTIMATE.COL_NUM_PAX + row).value,
    );
    const freq =
      extractResult(this.ws.getCell(BUDGET_ESTIMATE.COL_DAYS + row).value) || 1;

    let expenseGroup: ExpenseGroup =
      EXPENSE_GROUP.TRAINING_SCHOLARSHIPS_EXPENSES;
    let gaaObject: GAAObject = GAA_OBJECT.TRAINING_EXPENSES;
    let ppmp: YesNo = 'N';
    let appSupplies: YesNo = 'N';
    let appTicket: YesNo = 'N';
    let mannerOfRelease: MannerOfRelease = MANNER_OF_RELEASE.DIRECT_PAYMENT;

    if (expenseItem.includes(EXPENSE_PREFIX.BOARD_LODGING)) {
      mannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_BOARD;
    }

    if (expenseItem.includes(EXPENSE_PREFIX.TRAVEL)) {
      const keywords = BUDGET_ESTIMATE.NON_PAX_KEYWORDS;

      if (!keywords.some(s => expenseItem.includes(s))) {
        mannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_PSF;
      }
    }

    if (expenseItem.toLowerCase().includes('supplies')) {
      expenseGroup = EXPENSE_GROUP.SUPPLIES_EXPENSES;
      gaaObject = GAA_OBJECT.OTHER_SUPPLIES;
      appSupplies = 'Y';
      mannerOfRelease = MANNER_OF_RELEASE.CASH_ADVANCE;
    }

    if (expenseItem === 'Contingency') {
      mannerOfRelease = MANNER_OF_RELEASE.CASH_ADVANCE;
    }

    const expense: ExpenseItem = {
      expenseGroup,
      gaaObject,
      expenseItem,
      quantity,
      freq,
      unitCost,
      ppmp,
      appSupplies,
      appTicket,
      mannerOfRelease,
    };

    return expense;
  }
}

export default BudgetEstimate;
