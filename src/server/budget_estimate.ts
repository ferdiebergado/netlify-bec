import Worksheet from "./worksheet";
import {
  BudgetEstimateCell,
  BudgetEstimateRow,
  BudgetEstimateCol,
  ExpensePrefix,
  ExpenseGroup,
  GAAObject,
  MannerOfRelease,
  ExpenseItem,
  YesNo,
  ActivityInfo,
} from "./types";
import { extractResult } from "./utils";

class BudgetEstimate extends Worksheet {
  // eslint-disable-next-line @typescript-eslint/no-useless-constructor
  constructor(xls: ArrayBuffer, sheet: string) {
    super(xls, sheet);
  }

  get program() {
    return this.ws?.getCell(BudgetEstimateCell.PROGRAM).text;
  }

  get output() {
    return this.ws?.getCell(BudgetEstimateCell.OUTPUT).text;
  }

  get outputIndicator() {
    return this.ws?.getCell(BudgetEstimateCell.OUTPUT_INDICATOR).text;
  }

  get activity() {
    return this.ws?.getCell(BudgetEstimateCell.ACTIVITY).text;
  }

  get activityIndicator() {
    return this.ws?.getCell(BudgetEstimateCell.ACTIVITY_INDICATOR).text;
  }

  get month() {
    const stDate = this.ws?.getCell(BudgetEstimateCell.START_DATE).text;

    if (stDate) return new Date(stDate).getMonth();
  }

  get venue() {
    return this.ws?.getCell(BudgetEstimateCell.VENUE).text;
  }

  get totalPax() {
    return extractResult(this.ws?.getCell(BudgetEstimateCell.TOTAL_PAX).value);
  }

  get activityInfo() {
    const info: ActivityInfo = {
      program: this.program,
      output: this.output,
      outputIndicator: this.outputIndicator,
      activity: this.activity,
      activityIndicator: this.activityIndicator,
      month: this.month,
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
    const prefix = ExpensePrefix.BOARD_LODGING;
    const col = BudgetEstimateCol.BOARD_LODGING;
    const boardLodgingPax = this._parseExpenses(
      col,
      BudgetEstimateRow.BOARD_LODGING_START,
      BudgetEstimateRow.BOARD_LODGING_END,
      prefix
    );

    const boardLodgingOther = this._parseExpenses(
      col,
      BudgetEstimateRow.BOARD_LODGING_OTHER,
      BudgetEstimateRow.BOARD_LODGING_OTHER,
      prefix
    );

    return [...boardLodgingPax, ...boardLodgingOther];
  }

  travelExpenses() {
    const prefix = ExpensePrefix.TRAVEL;

    const travelRegion = this._parseExpenses(
      BudgetEstimateCol.TRAVEL_REGION,
      BudgetEstimateRow.TRAVEL_REGION_START,
      BudgetEstimateRow.TRAVEL_REGION_END,
      prefix
    );

    const travelCO = this._parseExpenses(
      BudgetEstimateCol.TRAVEL_CO,
      BudgetEstimateRow.TRAVEL_CO_START,
      BudgetEstimateRow.TRAVEL_CO_END,
      prefix
    );

    const travelOther = this._parseExpenses(
      BudgetEstimateCol.TRAVEL_OTHER,
      BudgetEstimateRow.TRAVEL_OTHER,
      BudgetEstimateRow.TRAVEL_OTHER,
      prefix
    );

    return [...travelRegion, ...travelCO, ...travelOther];
  }

  honorarium() {
    return this._parseExpenses(
      BudgetEstimateCol.HONORARIUM,
      BudgetEstimateRow.HONORARIUM_START,
      BudgetEstimateRow.HONORARIUM_END,
      ExpensePrefix.HONORARIUM
    );
  }

  suppliesContingency() {
    return this._parseExpenses(
      BudgetEstimateCol.SUPPLIES_CONTINGENCY,
      BudgetEstimateRow.SUPPLIES_CONTINGENCY_START,
      BudgetEstimateRow.SUPPLIES_CONTINGENCY_END,
      ""
    );
  }

  _parseExpenses(
    col: string,
    startRow: number,
    endRow: number,
    itemPrefix: string
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
    itemPrefix: string
  ) {
    const unitCost = extractResult(this.ws?.getCell("K" + row).value);

    if (!unitCost || unitCost == 0) return;

    const expenseItem =
      itemPrefix + this.ws?.getCell(itemCol + row).text.trim();
    const quantity = extractResult(this.ws?.getCell("H" + row).value);
    const freq = extractResult(this.ws?.getCell("J" + row).value) || 1;

    let expenseGroup = ExpenseGroup.TRAINING_SCHOLARSHIPS_EXPENSES;
    let gaaObject = GAAObject.TRAINING_EXPENSES;
    let ppmp: YesNo = "N";
    let appSupplies: YesNo = "N";
    let appTicket: YesNo = "N";
    let mannerOfRelease = MannerOfRelease.DIRECT_PAYMENT;

    if (expenseItem.includes(ExpensePrefix.BOARD_LODGING)) {
      mannerOfRelease = MannerOfRelease.FOR_DOWNLOAD_BOARD;
    }

    if (expenseItem.includes(ExpensePrefix.TRAVEL)) {
      const keywords = ["Resource", "Technical", "Bureau", "Other"];

      if (!keywords.some((s) => expenseItem.includes(s))) {
        mannerOfRelease = MannerOfRelease.FOR_DOWNLOAD_PSF;
      }
    }

    if (expenseItem.toLowerCase().includes("supplies")) {
      expenseGroup = ExpenseGroup.SUPPLIES_EXPENSES;
      gaaObject = GAAObject.OTHER_SUPPLIES;
      appSupplies = "Y";
      mannerOfRelease = MannerOfRelease.CASH_ADVANCE;
    }

    if (expenseItem === "Contingency") {
      mannerOfRelease = MannerOfRelease.CASH_ADVANCE;
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
