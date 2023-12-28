import { Workbook } from './workbook';
import {
  EXPENDITURE_MATRIX,
  MANNER_VALIDATION,
  YES,
  YES_NO_VALIDATION,
} from './constants';
import type { Activity, ExcelFile, ExpenseItem } from '../types/globals';
import { BudgetEstimate } from './budgetEstimate';

export class ExpenditureMatrix extends Workbook<ExpenditureMatrix> {
  programs: string[] = [];
  activities: Activity[] = [];

  constructor() {
    super();
  }

  protected createInstance(): ExpenditureMatrix {
    this.setActiveSheet(1);

    return this;
  }

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
  private _duplicateRows(
    targetRowIndex: number,
    srcRowIndex: number,
    numRows: number,
  ) {
    const sheet = this.getActiveSheet();
    for (let j = 0; j < numRows; j += 1) {
      const newRow = sheet.insertRow(targetRowIndex, []);
      const srcRow = sheet.getRow(srcRowIndex);

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

  /**
   * Duplicates the program row in a worksheet.
   *
   * @param ws - The worksheet where the program row will be duplicated.
   * @param targetRow - The index where the duplicate program row will be inserted.
   */
  private _duplicateProgram(targetRow: number) {
    this._duplicateRows(targetRow, EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX, 1);
  }

  /**
   * Duplicates the output row in the expenditure matrix.
   *
   * @param ws - The worksheet where the output row will be duplicated.
   * @param targetRow - The index where the duplicate output row will be inserted.
   */
  private _duplicateOutput(targetRow: number) {
    this._duplicateRows(targetRow, EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX, 1);
  }

  /**
   * Duplicates the activity row in a worksheet.
   *
   * @param ws - The worksheet where the activity row will be duplicated.
   * @param targetRow - The index where the duplicate activity row will be inserted.
   */
  private _duplicateActivity(targetRow: number) {
    this._duplicateRows(targetRow, EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX, 1);
  }

  /**
   * Duplicates the expense item row in a worksheet.
   *
   * @param ws - The worksheet where the expense item row will be duplicated.
   * @param targetRow - The index where the duplicate expense item row will be inserted.
   * @param count - The number of expense item rows to be duplicated.
   */
  private _duplicateExpenseItem(targetRow: number, count: number) {
    this._duplicateRows(
      targetRow,
      EXPENDITURE_MATRIX.EXPENSE_ITEM_ROW_INDEX,
      count,
    );
  }
  /**
   * Creates or duplicates an activity row in a worksheet.
   *
   * @param targetRow - The index where the activity row will be inserted.
   * @param activity - The activity information.
   * @param isFirst - A flag indicating if it is the first row. Default is `false`.
   *
   * @returns number - The index of the activity row
   */
  private _createActivityRow(
    targetRow: number,
    activity: Activity,
    isFirst: boolean = false,
  ): number {
    const sheet = this.getActiveSheet();
    let activityRowIndex = targetRow;

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

    if (isFirst) {
      activityRowIndex = ACTIVITY_ROW_INDEX;
    } else {
      this._duplicateActivity(targetRow);
    }

    const {
      info: { activityTitle, activityIndicator, activityPhysicalTarget, month },
      expenseItems,
    } = activity;

    const sumFormula = (cell: string) => ({
      formula: `SUM(${cell}${activityRowIndex + 1}:${cell}${
        activityRowIndex + expenseItems.length
      })`,
    });

    const activityRow = sheet.getRow(activityRowIndex);

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
        const col = cell.address.replace(/\d+/, '');

        cell.value = sumFormula(col);
      }
    });

    // obligation grand total
    activityRow.getCell(TOTAL_OBLIGATION_COL).value =
      sumFormula(TOTAL_OBLIGATION_COL);

    // disbursement grand total
    activityRow.getCell(TOTAL_DISBURSEMENT_COL).value = sumFormula(
      TOTAL_DISBURSEMENT_COL,
    );

    return activityRowIndex;
  }

  /**
   * Creates or duplicates an output row in a worksheet.
   *
   * @param ws - The worksheet where the output row will be created or duplicated.
   * @param targetRow - The index where the output row will be inserted.
   * @param activity - The activity information.
   * @param rank - The rank of the output.
   * @param isFirst - A flag indicating if it is the first row. Default is `false`.
   *
   * @returns void
   */
  private _createOutputRow(
    targetRow: number,
    activity: Activity,
    rank: number,
    isFirst = false,
  ): void {
    const sheet = this.getActiveSheet();

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
      this._duplicateOutput(targetRow);
    }

    const { output, outputIndicator, outputPhysicalTarget, month } =
      activity.info;

    // output
    const outputRow = sheet.getRow(outputRowIndex);
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

  /**
   * Creates or duplicates an expense item row in a worksheet.
   *
   * @param ws - The worksheet where the expense item row will be created or duplicated.
   * @param targetRow - The index where the expense item row will be inserted.
   * @param expense - The expense item information.
   * @param month - The month index.
   * @param isFirst - A flag indicating if it is the first row. Default is `false`.
   */
  private _createExpenseItemRow(
    targetRow: number,
    expense: ExpenseItem,
    month: number,
    isFirst = false,
  ) {
    const sheet = this.getActiveSheet();

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

    const currentRow = sheet.getRow(rowIndex);

    const {
      expenseGroup,
      gaaObject,
      expenseItem,
      quantity,
      unitCost,
      freq,
      tevLocation,
      hasPPMP,
      hasAPPSupplies,
      hasAPPTicket,
      releaseManner,
    } = expense;

    // expense group
    currentRow.getCell(EXPENSE_GROUP_COL).value = expenseGroup;

    // gaa object
    currentRow.getCell(GAA_OBJECT_COL).value = gaaObject;

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
    if (hasPPMP) ppmpCell.value = YES;

    // app supplies
    const appSuppliesCell = currentRow.getCell(APP_SUPPLIES_COL);
    appSuppliesCell.dataValidation = YES_NO_VALIDATION;
    if (hasAPPSupplies) appSuppliesCell.value = YES;

    // app ticket
    const appTicketCell = currentRow.getCell(APP_TICKET_COL);
    appTicketCell.dataValidation = YES_NO_VALIDATION;
    if (hasAPPTicket) appTicketCell.value = YES;

    // manner of release
    const mannerOfReleaseCell = currentRow.getCell(MANNER_OF_RELEASE_COL);
    mannerOfReleaseCell.value = releaseManner;
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

  /**
   * Orders activities based on program and output.
   *
   * @param a - The first activity.
   * @param b - The second activity.
   * @returns A number indicating the order.
   */
  private _orderByProgram(this: void, a: Activity, b: Activity): number {
    if (a.info.program < b.info.program) {
      return -1;
    }

    if (a.info.program > b.info.program) {
      return 1;
    }

    if (a.info.output < b.info.output) {
      return -1;
    }

    return 1;
  }

  async convert(files: ExcelFile[]): Promise<ArrayBuffer> {
    const sheet = this.getActiveSheet();

    await Promise.allSettled(files.map(file => this._addToActivities(file)));

    let targetRow = EXPENDITURE_MATRIX.TARGET_ROW_INDEX;
    let isFirst = true;
    let rank = 1;

    const activityRows: number[] = [];

    this.activities.sort(this._orderByProgram).forEach(activity => {
      const {
        info: { program, month },
        expenseItems,
      } = activity;

      let programRowIndex: number = targetRow;

      // program
      if (isFirst) {
        programRowIndex = EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX;
        this.programs.push(program);
      } else if (!this.programs.includes(program)) {
        this._duplicateProgram(targetRow);
        this.programs.push(program);
        targetRow += 1;
      }

      const programRow = sheet.getRow(programRowIndex);
      programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;

      // output
      this._createOutputRow(targetRow, activity, rank, isFirst);

      rank += 1;

      if (!isFirst) targetRow += 1;

      const activityRowIndex = this._createActivityRow(
        targetRow,
        activity,
        isFirst,
      );

      activityRows.push(activityRowIndex);

      if (!isFirst) targetRow += 1;

      // expense items
      this._duplicateExpenseItem(targetRow, expenseItems.length);

      expenseItems.forEach(expense => {
        this._createExpenseItemRow(targetRow, expense, month, isFirst);
        targetRow += 1;
      });

      isFirst = false;
      targetRow -= 1;
    });

    sheet.spliceRows(targetRow, 2);

    const { TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL } =
      EXPENDITURE_MATRIX;
    const lastRowIndex = targetRow + 1;
    const grandTotalRow = sheet.getRow(lastRowIndex);

    const setGrandTotalCell = (cell: string) => {
      const cellsWithTotals = activityRows.map(row => cell + row);

      grandTotalRow.getCell(cell).value = {
        formula: `SUM(${cellsWithTotals.toString()})`,
      };
    };

    [TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL].forEach(
      total => setGrandTotalCell(total),
    );

    for (let i = 0; i < 26; i += 1) {
      const col = i + 44;
      const cell = grandTotalRow.getCell(col);
      const cellsWithTotals = activityRows.map(
        row => cell.address.replace(/\d+/, '') + row,
      );

      cell.value = {
        formula: `SUM(${cellsWithTotals.toString()})`,
      };
    }

    const outBuff = await this.wb.xlsx.writeBuffer();

    return outBuff;
  }

  /**
   * Appends the activities of the specified Excel File to the activities array
   *
   * @param file {ExcelFile} - The Excel file containing the budget estimate
   */
  private async _addToActivities({ buffer }: ExcelFile): Promise<void> {
    const be = await BudgetEstimate.createAsync<BudgetEstimate>(buffer);
    const activities = be.getActivities();

    this.activities.push(...activities);
  }
}
