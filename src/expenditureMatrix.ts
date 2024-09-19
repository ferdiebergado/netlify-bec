import { Workbook } from './workbook';
import {
  EXPENDITURE_MATRIX,
  MANNER_VALIDATION,
  YES,
  YES_NO_VALIDATION,
} from './constants';
import type {
  Activity,
  ActivityInfo,
  ExcelFile,
  ExpenseItem,
} from './types/globals';
import { BudgetEstimate } from './budgetEstimate';
import type { Row, Worksheet } from 'exceljs';

/**
 * Represents a specialized workbook for managing expenditure matrices.
 *
 * @class ExpenditureMatrix
 * @extends {Workbook<ExpenditureMatrix>}
 */
export class ExpenditureMatrix extends Workbook<ExpenditureMatrix> {
  /**
   * An array to store program names.
   *
   * @public
   * @type {string[]}
   */
  programs: string[] = [];

  /**
   * An array to store activity information.
   *
   * @public
   * @type {Activity[]}
   */
  activities: Activity[] = [];

  /**
   * Creates an instance of the ExpenditureMatrix class.
   *
   * @public
   * @constructor
   */
  constructor() {
    // Calls the constructor of the base class (Workbook)
    super();
  }

  /**
   * Overrides the abstract method in the base class to create an instance of ExpenditureMatrix and set the active sheet.
   *
   * @protected
   * @returns {ExpenditureMatrix} The created instance of ExpenditureMatrix.
   */
  protected createInstance(): ExpenditureMatrix {
    // Sets the active sheet to the first sheet (index 1)
    this.setActiveSheet(1);

    // Returns the created instance of ExpenditureMatrix
    return this;
  }

  /**
   * Duplicates a specified row.
   *
   * @param ws {Worksheet} Sheet were the rows will be duplicated
   * @param targetRowIndex {number} Index where the duplicate rows will be inserted
   * @param srcRowIndex {number} Index of the row that will be duplicated
   * @param numRows {number} Number of rows to be duplicated
   *
   * @returns void
   */
  static duplicateRow(
    sheet: Worksheet,
    targetRowIndex: number,
    srcRowIndex: number,
    numRows: number = 1,
  ): void {
    const srcRow = sheet.getRow(srcRowIndex);
    let currentRowIndex = targetRowIndex;

    for (let j = 0; j < numRows; j += 1) {
      const newRow = sheet.insertRow(currentRowIndex, []);

      srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
        const targetCell = newRow.getCell(colNumber);

        targetCell.value = cell.value;
        targetCell.style = cell.style;
        targetCell.dataValidation = cell.dataValidation;
      });

      currentRowIndex += 1;
    }
  }

  static clearPreviousPhysicalTargets(row: Row) {
    for (let i = 0; i < 12; i += 1) {
      row.getCell(
        EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_COL_INDEX + i,
      ).value = '';
    }
  }

  static clearPreviousFinancialPrograms(row: Row) {
    const { OBLIGATION_MONTH_COL_INDEX, DISBURSEMENT_MONTH_COL_INDEX } =
      EXPENDITURE_MATRIX;

    for (let i = 0; i < 12; i += 1) {
      row.getCell(OBLIGATION_MONTH_COL_INDEX + i).value = '';
      row.getCell(DISBURSEMENT_MONTH_COL_INDEX + i).value = '';
    }
  }

  /**
   * Duplicates the program row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the duplicate program row will be inserted.
   *
   * @returns void
   */
  private _duplicateProgram(targetRowIndex: number): void {
    ExpenditureMatrix.duplicateRow(
      this.getActiveSheet(),
      targetRowIndex,
      EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX,
    );
  }

  /**
   * Duplicates the output row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the duplicate output row will be inserted.
   *
   * @returns void
   */
  private _duplicateOutput(targetRowIndex: number): void {
    ExpenditureMatrix.duplicateRow(
      this.getActiveSheet(),
      targetRowIndex,
      EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX,
    );
  }

  /**
   * Duplicates the activity row in a worksheet.
   *
   * @param targetRowIndex {number} The index where the duplicate activity row will be inserted.
   *
   * @returns void
   */
  private _duplicateActivity(targetRowIndex: number): void {
    ExpenditureMatrix.duplicateRow(
      this.getActiveSheet(),
      targetRowIndex,
      EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX,
    );
  }

  /**
   * Duplicates the expense item row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the duplicate expense item row will be inserted.
   * @param copies {number} The number of expense item rows to be copied.
   *
   * @returns void
   */
  private _duplicateExpenseItem(targetRowIndex: number, copies: number): void {
    ExpenditureMatrix.duplicateRow(
      this.getActiveSheet(),
      targetRowIndex,
      EXPENDITURE_MATRIX.EXPENSE_ITEM_ROW_INDEX,
      copies,
    );
  }
  /**
   * Creates or duplicates an activity row in the expenditure matrix.
   *
   * @param targetRow {number} The index where the activity row will be inserted.
   * @param activity {Activity} The activity information.
   * @param isFirstActivity {boolean} A flag indicating if the activity being created is the very first activity. Default is `false`.
   *
   * @returns number - The index of the activity row
   */
  private _createActivityRow(
    targetRow: number,
    activity: Activity,
    isFirstActivity: boolean = false,
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

    if (isFirstActivity) {
      activityRowIndex = ACTIVITY_ROW_INDEX;
    } else {
      this._duplicateActivity(activityRowIndex);
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

    if (!isFirstActivity) {
      ExpenditureMatrix.clearPreviousPhysicalTargets(activityRow);
    }

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
   * Creates or duplicates an output row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the output row will be inserted.
   * @param activity {Activity} The activity information.
   * @param rank {number} The rank of the output.
   * @param isFirstActivity {boolean} A flag indicating if the activity is the first activity to be created. Default is `false`.
   *
   * @returns void
   */
  private _createOutputRow(
    targetRowIndex: number,
    activity: Activity,
    rank: number,
    isFirstActivity: boolean = false,
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

    let outputRowIndex = targetRowIndex;

    if (isFirstActivity) {
      outputRowIndex = OUTPUT_ROW_INDEX;
    } else {
      this._duplicateOutput(outputRowIndex);
    }

    const { output, outputIndicator, outputPhysicalTarget, month } =
      activity.info;

    // output
    const outputRow = sheet.getRow(outputRowIndex);
    outputRow.getCell(OUTPUT_COL).value = output;
    outputRow.getCell(RANK_COL).value = rank;

    // output indicator
    outputRow.getCell(PERFORMANCE_INDICATOR_COL).value = outputIndicator;

    if (!isFirstActivity) {
      ExpenditureMatrix.clearPreviousPhysicalTargets(outputRow);
    }

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
   * Creates an expense item row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the expense item row will be inserted.
   * @param expense {ExpenseItem} The expense item information.
   * @param month {number} The month index.
   * @param isFirstActivity {boolean} A flag indicating if the activity being created is the very first activity. Default is `false`.
   *
   * @returns void
   */
  private _createExpenseItemRow(
    targetRowIndex: number,
    expense: ExpenseItem,
    month: number,
    isFirstActivity: boolean = false,
  ): void {
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

    let currentRowIndex = targetRowIndex;

    if (isFirstActivity) currentRowIndex -= 1;

    const currentRow = sheet.getRow(currentRowIndex);

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
      formula: `${QUANTITY_COL}${currentRowIndex}*${UNIT_COST_COL}${currentRowIndex}*${FREQUENCY_COL}${currentRowIndex}`,
    };

    // tev location
    currentRow.getCell(TEV_LOCATION_COL).value = tevLocation;

    // ppmp
    const ppmpCell = currentRow.getCell(PPMP_COL);
    ppmpCell.dataValidation = YES_NO_VALIDATION;
    ppmpCell.value = 'N';
    if (hasPPMP) ppmpCell.value = 'Y';

    // app supplies
    const appSuppliesCell = currentRow.getCell(APP_SUPPLIES_COL);
    appSuppliesCell.dataValidation = YES_NO_VALIDATION;
    appSuppliesCell.value = 'N';
    if (hasAPPSupplies) appSuppliesCell.value = 'Y';

    // app ticket
    const appTicketCell = currentRow.getCell(APP_TICKET_COL);
    appTicketCell.dataValidation = YES_NO_VALIDATION;
    if (hasAPPTicket) appTicketCell.value = YES;

    // manner of release
    const mannerOfReleaseCell = currentRow.getCell(MANNER_OF_RELEASE_COL);
    mannerOfReleaseCell.value = releaseManner;
    mannerOfReleaseCell.dataValidation = MANNER_VALIDATION;

    // total obligation
    const obligationMonthStartCell = `${OBLIGATION_MONTH_START_COL}${currentRowIndex}`;
    const obligationMonthEndCell = `${OBLIGATION_MONTH_END_COL}${currentRowIndex}`;

    currentRow.getCell(TOTAL_OBLIGATION_COL).value = {
      formula: `SUM(${obligationMonthStartCell}:${obligationMonthEndCell})`,
    };

    const totalRef = {
      formula: `${TOTAL_COST_COL}${currentRowIndex}`,
    };

    if (!isFirstActivity) {
      ExpenditureMatrix.clearPreviousFinancialPrograms(currentRow);
    }

    // obligation month
    currentRow.getCell(OBLIGATION_MONTH_COL_INDEX + month).value = totalRef;

    // total disbursement
    const disbursementMonthStartCell = `${DISBURSEMENT_MONTH_START_COL}${currentRowIndex}`;
    const disbursementMonthEndCell = `${DISBURSEMENT_MONTH_END_COL}${currentRowIndex}`;

    currentRow.getCell(TOTAL_DISBURSEMENT_COL).value = {
      formula: `SUM(${disbursementMonthStartCell}:${disbursementMonthEndCell})`,
    };

    // disbursement month
    currentRow.getCell(DISBURSEMENT_MONTH_COL_INDEX + month).value = totalRef;
  }

  /**
   * Converts an array of budget estimates to an expenditure matrix.
   *
   * @param files {ExcelFile[]} The array of files to be converted.
   *
   * @returns {Promise<ArrayBuffer>} A promise that resolves to the array buffer of the Expenditure Matrix
   */
  async fromBudgetEstimates(files: ExcelFile[]): Promise<ArrayBuffer> {
    const sheet = this.getActiveSheet();

    await Promise.allSettled(files.map(file => this._addToActivities(file)));

    let currentRowIndex: number = EXPENDITURE_MATRIX.TARGET_ROW_INDEX;
    let isFirstActivity = true;
    let rank = 1;

    /**
     * Records the indices of rows of each activity that contains the total unit cost
     * to be used later to compute the grand total
     */
    const activityRows: number[] = [];

    const paxTevsExpenses: ExpenseItem[] = [];

    this.activities.sort(this._orderByProgram).forEach(activity => {
      const {
        info: { program, month },
        expenseItems,
      } = activity;

      // program
      let programRowIndex = currentRowIndex;

      if (isFirstActivity) {
        programRowIndex = EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX;
        this.programs.push(program);
      } else {
        if (!this.programs.includes(program)) {
          this._duplicateProgram(programRowIndex);
          this.programs.push(program);
          currentRowIndex += 1;
        }
      }

      const programRow = sheet.getRow(programRowIndex);
      programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;

      // output
      this._createOutputRow(currentRowIndex, activity, rank, isFirstActivity);

      rank += 1;

      if (!isFirstActivity) currentRowIndex += 1;

      // activity
      const activityRowIndex = this._createActivityRow(
        currentRowIndex,
        activity,
        isFirstActivity,
      );

      activityRows.push(activityRowIndex);

      if (!isFirstActivity) currentRowIndex += 1;

      // expense items
      this._duplicateExpenseItem(currentRowIndex, expenseItems.length);

      expenseItems.forEach(expense => {
        const { expenseItem, unitCost } = expense;

        if (expenseItem.includes('Participant')) {
          const reg = paxTevsExpenses.find(ex => ex.expenseItem == expenseItem);

          if (reg) {
            reg.unitCost += unitCost;
          } else {
            paxTevsExpenses.push(expense);
          }
        } else {
          this._createExpenseItemRow(
            currentRowIndex,
            expense,
            month,
            isFirstActivity,
          );
          currentRowIndex += 1;
        }
      });

      if (isFirstActivity) currentRowIndex -= 1;

      isFirstActivity = false;
    });
    // activity
    const tevPaxInfo: ActivityInfo = {
      program: 'BEC',
      output: 'Benefitted implementers',
      outputIndicator: 'No. of implementers benefitted',
      activityTitle:
        'Program Support Funds for Travelling Expenses of Participants',
      activityIndicator: 'No. of program support funds downloaded',
      month: 1,
      venue: '',
      totalPax: 16,
      outputPhysicalTarget: 16,
      activityPhysicalTarget: 1,
    };

    const tevPaxActivity: Activity = {
      info: tevPaxInfo,
      expenseItems: paxTevsExpenses,
    };

    const activityRowIndex = this._createActivityRow(
      currentRowIndex,
      tevPaxActivity,
      isFirstActivity,
    );

    activityRows.push(activityRowIndex);

    // if (!isFirstActivity) currentRowIndex += 1;

    // expense items
    this._duplicateExpenseItem(currentRowIndex, paxTevsExpenses.length);

    paxTevsExpenses.forEach(exp => {
      this._createExpenseItemRow(currentRowIndex, exp, 1, isFirstActivity);
      currentRowIndex += 1;
    });

    sheet.spliceRows(currentRowIndex, 2);

    const lastRowIndex = currentRowIndex + 1;

    const { TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL } =
      EXPENDITURE_MATRIX;
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

    const buffer = await this.wb.xlsx.writeBuffer();

    return buffer;
  }

  /**
   * Appends the activities of the specified Excel File to the activities array.
   *
   * @param file {ExcelFile} The Excel file containing the budget estimate
   *
   * @returns void
   */
  private async _addToActivities({ buffer }: ExcelFile): Promise<void> {
    const budgetEstimate =
      await BudgetEstimate.createAsync<BudgetEstimate>(buffer);
    const activities = budgetEstimate.getActivities();

    this.activities.push(...activities);
  }

  /**
   * Orders activities based on program and output.
   *
   * @param a {Activity} The first activity.
   * @param b {Activity} The other activity.
   *
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
}
