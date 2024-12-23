import { Workbook } from './workbook.js';
import {
  EXPENDITURE_MATRIX,
  FONT,
  MANNER_VALIDATION,
  MAX_MONTH,
  MONTHS_IN_A_YEAR,
  YES,
  YES_NO_VALIDATION,
} from './constants.js';
import type {
  Activity,
  ActivityInfo,
  ExcelFile,
  ExpenseItem,
  ExpenseItemRowContext,
  ExpenditureFile,
} from './types/globals.js';
import { BudgetEstimate } from './budgetEstimate.js';
import type { Cell, Row } from 'exceljs';
import {
  deepFreeze,
  extractProgramTitle,
  isCellFormulaValue,
} from './utils.js';

/**
 * Represents a specialized workbook for managing expenditure matrices.
 *
 * @class ExpenditureMatrix
 * @extends {Workbook<ExpenditureMatrix>}
 */
export class ExpenditureMatrix extends Workbook<ExpenditureMatrix> {
  /**
   * An array that stores the parsed activities.
   *
   * @protected
   * @type {Activity[]}
   */
  private activities: Activity[] = [];

  /**
   * The current output rank.
   *
   * @protected
   * @type {number}
   */
  private rank: number = 1;

  /**
   * The current row index being processed.
   */
  private currentRowIndex = 0;

  /**
   * The current program
   */
  private currentProgram = '';

  private currentOutput = '';

  /**
   * Status flag that indicates if the current activity
   * is the first activity being processed
   */
  private isFirstActivity = true;

  /**
   * The font settings
   */
  private readonly font = deepFreeze(FONT);

  /**
   * The PSF Activity
   */
  private PSF: Activity = {
    info: {
      program: 'Programs Support Funds (PSF)',
      output: 'Benefitted implementers',
      outputIndicator: 'No. of implementers benefitted',
      outputPhysicalTarget: 16,
      activityTitle:
        'Provision of Program Support Funds for the Travel Expenses of Field Participants',
      activityIndicator: 'No. of downloading activities conducted',
      activityPhysicalTarget: 1,
      month: 1,
      venue: '',
      totalPax: 16,
    },
    expenseItems: [],
    tevPSF: [],
  };

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
   * Duplicates a cell to another cell
   *
   * @param sourceCell The source cell to be copied
   * @param targetCell The target cell
   */
  private static _duplicateCell(sourceCell: Cell, targetCell: Cell): void {
    Object.assign(targetCell, {
      value: isCellFormulaValue(sourceCell.value)
        ? { sharedFormula: sourceCell.address }
        : sourceCell.value,
      style: sourceCell.style,
      dataValidation: sourceCell.dataValidation,
    });
  }

  /**
   * Clears all entries in the physical target columns.
   *
   * @private
   * @static
   *
   * @param row {Row} - The target row   *
   *
   * @returns {void}
   */
  private static _clearPreviousPhysicalTargets(row: Row): void {
    Array.from({ length: MONTHS_IN_A_YEAR }, (_, i) => {
      row.getCell(
        EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_COL_INDEX + i,
      ).value = '';
    });
  }

  /**
   * Clears all entries in the obligation and disbursement columns.
   *
   * @private
   * @static
   * @param row {Row} - The target row
   *
   * @returns {void}
   */
  private static _clearPreviousFinancialPrograms(row: Row): void {
    const { OBLIGATION_MONTH_COL_INDEX, DISBURSEMENT_MONTH_COL_INDEX } =
      EXPENDITURE_MATRIX;

    Array.from({ length: MONTHS_IN_A_YEAR }, (_, i) => {
      row.getCell(OBLIGATION_MONTH_COL_INDEX + i).value = '';
      row.getCell(DISBURSEMENT_MONTH_COL_INDEX + i).value = '';
    });
  }

  /**
   * Gets the letter part of the cell address
   *
   * @param cell The target cell
   * @returns The letter part of the cell address
   */
  private static _getCellCol(cell: Cell): string {
    return cell.address.replace(/\d+/, '');
  }

  /**
   * Increments the specified month by 1.
   *
   * @private
   * @static
   * @param month The target month
   *
   * @returns {number} The incremented month
   */
  private static _incrementMonth(month: number): number {
    if (month < MAX_MONTH) return month + 1;
    return month;
  }

  /**
   * Orders activities based on program and output.
   *
   * @static
   * @param a {Activity} The first activity.
   * @param b {Activity} The other activity.
   *
   * @returns A number indicating the order.
   */
  private static _orderByProgramAndOutput(
    this: void,
    a: Activity,
    b: Activity,
  ): number {
    const { program: programA, output: outputA } = a.info;
    const { program: programB, output: outputB } = b.info;

    const programComparison = programA.localeCompare(programB);

    if (programComparison === 0) {
      return outputA.localeCompare(outputB);
    }

    return programComparison;
  }

  /**
   * Converts an array of budget estimates to an expenditure matrix.
   *
   * @param budgetEstimates {ExcelFile[]} The array of files to be converted.
   *
   * @returns {Promise<ArrayBuffer>} A promise that resolves to the array buffer of the Expenditure Matrix
   */
  async fromBudgetEstimates(
    budgetEstimates: ExcelFile[],
  ): Promise<ExpenditureFile> {
    await this._loadActivities(budgetEstimates);
    this._prepareEM();

    const {
      PROGRAM_ROW_INDEX,
      OUTPUT_ROW_INDEX,
      ACTIVITY_ROW_INDEX,
      GAA_OBJECT_FORMULA_COL,
      GAA_OBJECT_FORMULA_CELL,
      TOTAL_COST_COL,
      TOTAL_OBLIGATION_COL,
      TOTAL_DISBURSEMENT_COL,
      MONTHLY_PROGRAM_NUM_ROWS,
      TOTAL_OBLIGATION_COL_INDEX,
      OVERHEAD_NUM_ROWS,
      PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL,
      PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX,
      CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL,
      CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX,
      OVERHEAD_TOTAL_ROW_MAPPINGS,
      EXPENSE_ITEM_ROW_INDEX,
      OBLIGATION_MONTH_COL_INDEX,
    } = EXPENDITURE_MATRIX;

    /**
     * Records the indices of rows of each activity that contains the total unit cost
     * to be used later to compute the grand total
     */
    const activityRows: number[] = [];

    const sheet = this.getActiveSheet();

    this.activities.sort(ExpenditureMatrix._orderByProgramAndOutput);

    // first activity needs special treatment because it will not duplicate rows
    // but instead overwrite the sample rows in the template
    const firstActivity = this.activities[0];

    if (!firstActivity) throw new Error('No activities found.');

    const { info, tevPSF } = firstActivity;
    const { program, output } = info;

    this._createProgram(program, PROGRAM_ROW_INDEX);
    this._createOutput(output, info, OUTPUT_ROW_INDEX);
    activityRows.push(ACTIVITY_ROW_INDEX);
    this._createActivity(firstActivity, ACTIVITY_ROW_INDEX);
    this._createExpenseItems(firstActivity, EXPENSE_ITEM_ROW_INDEX);
    this._aggregatePSF(tevPSF);

    // process the rest of the activities
    this.activities.slice(1).forEach(activity => {
      console.log('sliced');

      const { info, tevPSF } = activity;
      const { program, output } = info;

      // program
      this._createProgram(program);

      // output
      this._createOutput(output, info);

      // activity
      activityRows.push(this.currentRowIndex);
      this._createActivity(activity);

      // Expense Items
      this._createExpenseItems(activity);

      // Aggregate TEVs of participants for PSF
      this._aggregatePSF(tevPSF);

      console.log('activity created');
      console.log('currentrowindex:', this.currentRowIndex);

      this.isFirstActivity = false;
    });

    // PSF
    if (this.PSF.expenseItems.length > 0) {
      console.log('Creating PSF Activity...');

      // program
      const { info } = this.PSF;
      const { program, output } = info;
      this._createProgram(program);

      // output
      this._createOutput(output, info);

      // activity
      activityRows.push(this.currentRowIndex);
      this._createActivity(this.PSF);

      // expense items
      this._createExpenseItems(this.PSF);
    }

    sheet.spliceRows(this.currentRowIndex, 1);

    console.log('last row index:', this.currentRowIndex);

    const lastRowIndex = this.currentRowIndex + OVERHEAD_NUM_ROWS + 1;

    // Costing Grand Total
    const grandTotalRow = sheet.getRow(lastRowIndex);
    const setGrandTotalCell = (cell: string) => {
      const cellsWithTotals = activityRows.map(row => cell + row);

      const grandTotalCell = grandTotalRow.getCell(cell);
      grandTotalCell.value = {
        formula: `SUM(${cellsWithTotals.toString()})`,
      };

      const font = Object.assign(grandTotalCell.font, {
        italic: false,
        strike: false,
      });
      grandTotalCell.font = font;
    };

    [TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL].forEach(
      total => setGrandTotalCell(total),
    );

    // Monthly Program Grand Totals
    Array.from({ length: MONTHLY_PROGRAM_NUM_ROWS }, (_, i) => {
      const colIndex = TOTAL_OBLIGATION_COL_INDEX + i;
      const cell = grandTotalRow.getCell(colIndex);
      const cellsWithTotals = activityRows.map(
        row => cell.address.replace(/\d+/, '') + row,
      );

      cell.value = {
        formula: `SUM(${cellsWithTotals.toString()})`,
      };
    });

    // Overhead output rank
    this._setOutputRank(sheet.getRow(this.currentRowIndex + 1));

    // Overhead hidden columns formulas
    Array.from({ length: OVERHEAD_NUM_ROWS }, (_, i) => {
      const currentRow = this.currentRowIndex + i;

      // Expense Object
      // console.log(
      //   'setting overhead expense object formula at row:',
      //   currentRow,
      // );

      this._setExpenseObjectFormula(currentRow);

      // GAA Object
      // console.log('setting overhead gaa object formula at row:', currentRow);
      ExpenditureMatrix._duplicateCell(
        sheet.getCell(GAA_OBJECT_FORMULA_CELL),
        sheet.getRow(currentRow).getCell(GAA_OBJECT_FORMULA_COL),
      );

      // ISBLANK formulas
      // console.log('setting overhead isblank formula at row:', currentRow);
      this._setIsBlankFormulas(currentRow);
    });

    // Overhead Totals
    OVERHEAD_TOTAL_ROW_MAPPINGS.forEach(rowMap => {
      const { rowsToAdd, expenseItemsCount } = rowMap;
      const rowIndex = this.currentRowIndex + rowsToAdd;
      const currentRow = sheet.getRow(rowIndex);

      if (expenseItemsCount) {
        [TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL].forEach(
          col => {
            // console.log('setting overhead total formula at row:', rowIndex);

            currentRow.getCell(col).value = {
              formula: `SUM(${col}${rowIndex + 1}:${col}${
                rowIndex + expenseItemsCount
              })`,
            };
          },
        );

        // console.log(
        //   'setting overhead monthly program totals at row:',
        //   rowIndex,
        // );
        Array.from({ length: MONTHLY_PROGRAM_NUM_ROWS }, (_, i) => {
          const targetCell = currentRow.getCell(OBLIGATION_MONTH_COL_INDEX + i);
          const col = ExpenditureMatrix._getCellCol(targetCell);

          targetCell.value = {
            formula: `SUM(${col}${rowIndex + 1}:${col}${
              rowIndex + expenseItemsCount
            })`,
          };
        });
      }

      // Expense items
      if (expenseItemsCount) {
        [TOTAL_COST_COL, TOTAL_OBLIGATION_COL, TOTAL_DISBURSEMENT_COL].forEach(
          col => {
            Array.from({ length: expenseItemsCount }, (_, count) => {
              const targetRowIndex = rowIndex + 1 + count;

              // console.log(
              //   'setting overhead expense items total at row:',
              //   targetRowIndex,
              // );
              ExpenditureMatrix._duplicateCell(
                sheet.getCell(col + EXPENSE_ITEM_ROW_INDEX),
                sheet.getRow(targetRowIndex).getCell(col),
              );
            });
          },
        );
      }

      // Previous Year Physical Target
      // console.log(
      //   'setting overhead previous year physical target object at row:',
      //   currentRow,
      // );
      ExpenditureMatrix._duplicateCell(
        sheet.getCell(PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL),
        currentRow.getCell(PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX),
      );

      // Current Year Physical Target
      // console.log(
      //   'setting overhead current year physical target object at row:',
      //   currentRow,
      // );
      ExpenditureMatrix._duplicateCell(
        sheet.getCell(CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL),
        currentRow.getCell(CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX),
      );
    });

    const buffer = await this.wb.xlsx.writeBuffer();
    const programTitle = this._getProgram();

    return { programTitle, buffer };
  }

  /**
   * Duplicates the program row in the expenditure matrix.
   *
   * @private
   * @param targetRowIndex {number} The index where the duplicate program row will be inserted.
   *
   * @returns {void}
   */
  private _duplicateProgram(): void {
    this._duplicateRow(EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX);
  }

  /**
   * Duplicates the output row in the expenditure matrix.
   *
   * @private
   *
   * @returns {void}
   */
  private _duplicateOutput(): void {
    this._duplicateRow(EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX);
  }

  /**
   * Duplicates the activity row in a worksheet.
   *
   * @private
   *
   * @returns {void}
   */
  private _duplicateActivity(): void {
    this._duplicateRow(EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX);
  }

  /**
   * Duplicates an expense item at the specified target row index in the expenditure matrix.
   *
   * @private
   * @param {number} targetRowIndex - The row index where the expense item is to be duplicated.
   * @param {number|undefined} copies - The number of copies to create (default is 1).
   *
   * @returns {void}
   */
  private _duplicateExpenseItem(): void {
    this._duplicateRow(EXPENDITURE_MATRIX.EXPENSE_ITEM_ROW_INDEX);
  }

  /**
   * Duplicates a specified row.
   *
   * @private
   * @param range {RowCopyMap} Contains the indices of the target and source rows that will be duplicated
   *
   * @returns {void}
   */
  private _duplicateRow(srcRowIndex: number): void {
    const sheet = this.getActiveSheet();
    const srcRow = sheet.getRow(srcRowIndex);
    const newRow = sheet.insertRow(this.currentRowIndex, []);

    srcRow.font = Object.assign({}, this.font);

    srcRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
      const targetCell = newRow.getCell(colNumber);
      ExpenditureMatrix._duplicateCell(sourceCell, targetCell);
    });

    console.log(
      'duplicated row at index:',
      this.currentRowIndex,
      'source row index:',
      srcRowIndex,
    );

    // move to the next row
    this.currentRowIndex += 1;
    console.log('moved to the next row:', this.currentRowIndex);
  }

  /**
   * Sets the ISBLANK formula for the given row
   *
   * @param rowIndex The index of the target row
   */
  private _setIsBlankFormulas(rowIndex: number) {
    const {
      IS_BLANK_FORMULA_CELL1,
      IS_BLANK_FORMULA_CELL2,
      IS_BLANK_FORMULA_CELL3,
      IS_BLANK_FORMULA_START_COL,
    } = EXPENDITURE_MATRIX;
    const sheet = this.getActiveSheet();

    interface FormulaCells {
      formulaCell: string;
      count: number;
    }

    const formulaCells: FormulaCells[] = [
      { formulaCell: IS_BLANK_FORMULA_CELL1, count: 2 },
      { formulaCell: IS_BLANK_FORMULA_CELL2, count: 2 },
      { formulaCell: IS_BLANK_FORMULA_CELL3, count: 1 },
    ];

    let colIndex = IS_BLANK_FORMULA_START_COL;

    formulaCells.forEach(cell => {
      const sourceCell = sheet.getCell(cell.formulaCell);

      Array.from({ length: cell.count }, () => {
        const targetCell = sheet.getRow(rowIndex).getCell(colIndex);

        ExpenditureMatrix._duplicateCell(sourceCell, targetCell);
        colIndex += 1;
      });
    });
  }

  /**
   * Sets the Expense Object Formula for the given row.
   *
   * @param rowIndex The index of the target row
   */
  private _setExpenseObjectFormula(rowIndex: number) {
    const sheet = this.getActiveSheet();

    const targetCell = sheet
      .getRow(rowIndex)
      .getCell(EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_COL);

    ExpenditureMatrix._duplicateCell(
      sheet.getCell(EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_CELL),
      targetCell,
    );
  }

  /**
   * Creates or duplicates an activity row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the activity row will be inserted.
   * @param activity {Activity} The activity information.
   *
   * @returns number - The index of the activity row
   */
  private _createActivityRow(targetRowIndex: number, activity: Activity): void {
    const {
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

    const {
      info: { activityTitle, activityIndicator, activityPhysicalTarget, month },
      expenseItems,
    } = activity;

    const sumFormula = (cell: string) => ({
      formula: `SUM(${cell}${targetRowIndex + 1}:${cell}${
        targetRowIndex + expenseItems.length
      })`,
    });

    const activityRow = this.getActiveSheet().getRow(targetRowIndex);

    // activity title
    activityRow.getCell(ACTIVITIES_COL).value = activityTitle;

    // activity indicator
    activityRow.getCell(PERFORMANCE_INDICATOR_COL).value = activityIndicator;

    ExpenditureMatrix._clearPreviousPhysicalTargets(activityRow);

    // activity physical target
    const targetMonth = ExpenditureMatrix._incrementMonth(month);
    const physTargetCell = activityRow.getCell(
      PHYSICAL_TARGET_MONTH_COL_INDEX + targetMonth,
    );
    physTargetCell.value = activityPhysicalTarget;
    Object.assign(physTargetCell.font, { bold: false });

    // costing grand total
    const totalCostCell = activityRow.getCell(TOTAL_COST_COL);
    totalCostCell.value = sumFormula(TOTAL_COST_COL);
    totalCostCell.font.bold = true;

    // physical target grand total
    const physicalTargetMonthStartCell = `${PHYSICAL_TARGET_MONTH_START_COL_INDEX}${targetRowIndex}`;
    const physicalTargetMonthEndCell = `${PHYSICAL_TARGET_MONTH_END_COL_INDEX}${targetRowIndex}`;
    activityRow.getCell(PHYSICAL_TARGET_TOTAL_COL).value = {
      formula: `SUM(${physicalTargetMonthStartCell}:${physicalTargetMonthEndCell})`,
    };

    // obligation and disbursement grand total per month
    [OBLIGATION_MONTH_COL_INDEX, DISBURSEMENT_MONTH_COL_INDEX].forEach(
      programIndex => {
        Array.from({ length: MONTHS_IN_A_YEAR }, (_, monthPointer) => {
          const monthCell = activityRow.getCell(programIndex + monthPointer);
          const cellLetterPart = ExpenditureMatrix._getCellCol(monthCell);

          monthCell.value = sumFormula(cellLetterPart);
        });
      },
    );

    // obligation grand total
    activityRow.getCell(TOTAL_OBLIGATION_COL).value =
      sumFormula(TOTAL_OBLIGATION_COL);

    // disbursement grand total
    activityRow.getCell(TOTAL_DISBURSEMENT_COL).value = sumFormula(
      TOTAL_DISBURSEMENT_COL,
    );
  }

  /**
   * Sets the current output rank for a given row.
   *
   * @param row The target row
   */
  private _setOutputRank(row: Row) {
    row.getCell(EXPENDITURE_MATRIX.RANK_COL).value = this.rank;
  }

  /**
   * Creates or duplicates an output row in the expenditure matrix.
   *
   * @param targetRowIndex {number} The index where the output row will be inserted.
   * @param activityInfo {Activity} The activity information.
   * @param rank {number} The rank of the output.
   *
   * @returns {void}
   */
  private _createOutputRow(
    targetRowIndex: number,
    activityInfo: ActivityInfo,
  ): void {
    const sheet = this.getActiveSheet();

    const {
      OUTPUT_COL,
      PERFORMANCE_INDICATOR_COL,
      PHYSICAL_TARGET_MONTH_COL_INDEX,
      PHYSICAL_TARGET_TOTAL_COL,
      PHYSICAL_TARGET_MONTH_START_COL_INDEX,
      PHYSICAL_TARGET_MONTH_END_COL_INDEX,
    } = EXPENDITURE_MATRIX;

    const { output, outputIndicator, outputPhysicalTarget, month } =
      activityInfo;

    console.table(activityInfo);
    console.log('rowindex:', targetRowIndex);

    // output
    const outputRow = sheet.getRow(targetRowIndex);

    console.log('outputrow:', outputRow.number);

    outputRow.getCell(OUTPUT_COL).value = output;

    // output rank
    this._setOutputRank(outputRow);

    // output indicator
    outputRow.getCell(PERFORMANCE_INDICATOR_COL).value = outputIndicator;

    ExpenditureMatrix._clearPreviousPhysicalTargets(outputRow);

    // output physical target
    const targetMonth = ExpenditureMatrix._incrementMonth(month);
    const physTargetCell = outputRow.getCell(
      PHYSICAL_TARGET_MONTH_COL_INDEX + targetMonth,
    );

    physTargetCell.value = outputPhysicalTarget;
    physTargetCell.font.bold = false;

    // physical target grand total
    const physicalTargetMonthStartCell = `${PHYSICAL_TARGET_MONTH_START_COL_INDEX}${targetRowIndex}`;
    const physicalTargetMonthEndCell = `${PHYSICAL_TARGET_MONTH_END_COL_INDEX}${targetRowIndex}`;

    outputRow.getCell(PHYSICAL_TARGET_TOTAL_COL).value = {
      formula: `SUM(${physicalTargetMonthStartCell}:${physicalTargetMonthEndCell})`,
    };
  }

  /**
   * Sets the GAA Object formula of a row.
   *
   * @param rowIndex The index of the target row
   */
  private _setGAAObjFormula(rowIndex: number) {
    const targetCell = this.getActiveSheet()
      .getRow(rowIndex)
      .getCell(Number(EXPENDITURE_MATRIX.GAA_OBJECT_FORMULA_COL));

    Object.assign(targetCell, {
      value: {
        formula: `VLOOKUP(L${rowIndex},'links'!K$2:L$222,2,FALSE())`,
      },
    });
  }

  /**
   * Creates an expense item row in the expenditure matrix.
   *
   * @private
   * @param context {ExpenseItemRowContext} The context of the expense item row will be inserted.
   *
   * @returns {void}
   */
  private _createExpenseItemRow(
    rowIndex: number,
    { expense, month }: ExpenseItemRowContext,
  ): void {
    console.log('creating expense item row at index', rowIndex);
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

    // Expense Object formula
    this._setExpenseObjectFormula(rowIndex);

    // GAA Object formula
    this._setGAAObjFormula(rowIndex);

    // Is blank formulas
    this._setIsBlankFormulas(rowIndex);

    console.log('currentrowindex:', rowIndex);

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
    const obligationMonthStartCell = `${OBLIGATION_MONTH_START_COL}${rowIndex}`;
    const obligationMonthEndCell = `${OBLIGATION_MONTH_END_COL}${rowIndex}`;

    currentRow.getCell(TOTAL_OBLIGATION_COL).value = {
      formula: `SUM(${obligationMonthStartCell}:${obligationMonthEndCell})`,
    };

    const totalRef = {
      formula: `${TOTAL_COST_COL}${rowIndex}`,
    };

    ExpenditureMatrix._clearPreviousFinancialPrograms(currentRow);

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
   * Writes the name of the program to the specified/current row.
   *
   * @param program The name of the program
   * @param rowIndex Index of the target row
   */
  private _createProgram(program: string, rowIndex?: number): void {
    if (program !== this.currentProgram) {
      let programIndex: number;

      if (rowIndex) {
        programIndex = rowIndex;
        this.currentRowIndex = programIndex;
      } else {
        programIndex = this.currentRowIndex;
        this._duplicateProgram();
      }

      this.getActiveSheet()
        .getRow(programIndex)
        .getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;
      this.currentProgram = program;

      console.log(`created program at row ${rowIndex}, program: ${program}`);
    }
  }

  /**
   * Writes the output to the specified/current row.
   *
   * @param output The output text to be written
   * @param activityInfo The activity info to be added to the output
   * @param rowIndex Index of the target row
   */
  private _createOutput(
    output: string,
    activityInfo: ActivityInfo,
    rowIndex?: number,
  ): void {
    if (output !== this.currentOutput) {
      let outputIndex: number;

      if (rowIndex) {
        outputIndex = rowIndex;
        this.currentRowIndex = outputIndex;
      } else {
        outputIndex = this.currentRowIndex;
        this._duplicateOutput();
      }

      this._createOutputRow(outputIndex, activityInfo);
      this.currentOutput = output;
      this.rank += 1;

      console.log(`created output at row ${outputIndex}, output: ${output}`);
    }
  }

  /**
   * Writes the activity to the specified/current row.
   *
   * @param activity The activity to be created
   * @param rowIndex Index of the target row
   */
  private _createActivity(activity: Activity, rowIndex?: number): void {
    let activityIndex: number;

    if (rowIndex) {
      activityIndex = rowIndex;
      this._setExpenseObjectFormula(activityIndex);
      this._setIsBlankFormulas(activityIndex);
      this.currentRowIndex = activityIndex;
    } else {
      activityIndex = this.currentRowIndex;
      this._duplicateActivity();
    }

    this._createActivityRow(activityIndex, activity);

    console.log(
      `created activity at row ${activityIndex}, activity title: ${activity.info.activityTitle}`,
    );
  }

  /**
   * Parses a list of files to get the activities.
   *
   * @param budgetEstimates Budget Estimate Files to load
   */
  private async _loadActivities(
    budgetEstimates: ExcelFile[],
  ): Promise<void | never> {
    await Promise.all(
      budgetEstimates.map(budgetEstimate =>
        this._addToActivities(budgetEstimate),
      ),
    );

    if (this.activities.length === 0) throw new Error('No activities found.');
  }

  /**
   * Removes excess rows and fixes the header fonts of the Expenditure template
   */
  private _prepareEM() {
    const { HEADER_FIRST_ROW_INDEX, HEADER_LAST_ROW_INDEX } =
      EXPENDITURE_MATRIX;

    this._removeExtraRows();
    this._fixFonts(HEADER_FIRST_ROW_INDEX, HEADER_LAST_ROW_INDEX);
  }

  /**
   * Calculates the sum of the TEVS for each region.
   *
   * @param TEVs The TEVs per region
   */
  private _aggregatePSF(TEVs: ExpenseItem[]) {
    TEVs.forEach(tev => {
      const existingPSF = this.PSF.expenseItems.find(
        i => i.expenseItem === tev.expenseItem,
      );

      if (existingPSF) {
        existingPSF.unitCost += tev.unitCost * tev.quantity;
      } else {
        this.PSF.expenseItems.push({
          ...tev,
          quantity: 1,
          unitCost: tev.unitCost * tev.quantity,
        });
      }
    });
  }

  /**
   * Creates the expense items at the specified row on the active sheet.
   *
   * @private
   * @param context {ActivityContext} The context of the Activity
   *
   */
  private _createExpenseItems(activity: Activity, rowIndex?: number): void {
    const { info, expenseItems } = activity;
    const { month } = info;

    let items = expenseItems;

    if (rowIndex) {
      this.currentRowIndex = rowIndex;

      const firstItem = expenseItems[0];

      if (firstItem && this.isFirstActivity) {
        this._createExpenseItemRow(rowIndex, { expense: firstItem, month });
        this.currentRowIndex += 1;
        items = expenseItems.slice(1);
      }
    }

    items.forEach(expense => {
      console.dir(expense);

      const expenseIndex = this.currentRowIndex;

      console.log('duplicating expense item at row', expenseIndex);
      this._duplicateExpenseItem();

      this._createExpenseItemRow(expenseIndex, { expense, month });
      console.log('created expense item at row', expenseIndex);
    });
  }

  /**
   * Appends the activities of the specified Excel File to the activities array.
   *
   * @param file {ExcelFile} The Excel file containing the budget estimate
   *
   * @returns void
   */
  private async _addToActivities(file: ExcelFile): Promise<void> {
    const budgetEstimate =
      await BudgetEstimate.createAsync<BudgetEstimate>(file);
    const activities = budgetEstimate.getActivities();

    this.activities.push(...activities);
  }

  /**
   * Removes the excess rows in the Expenditure template
   */
  private _removeExtraRows() {
    const {
      EXTRA_ROWS_START_INDEX,
      EXTRA_ROWS_NUM_ROWS,
      MILESTONES_START_ROW,
      MILESTONES_NUM_ROWS,
    } = EXPENDITURE_MATRIX;
    const sheet = this.getActiveSheet();

    sheet.spliceRows(EXTRA_ROWS_START_INDEX, EXTRA_ROWS_NUM_ROWS);
    sheet.spliceRows(MILESTONES_START_ROW, MILESTONES_NUM_ROWS);
  }

  /**
   * Fixes the font styles of the given row(s).
   *
   * @param startRowIndex Index of the start row
   * @param count Number of rows to fix
   */
  private _fixFonts(startRowIndex: number, count: number) {
    Array.from({ length: count }, (_, i) => {
      const row = this.getActiveSheet().getRow(startRowIndex + i);
      const cell = row.getCell(1);
      cell.font = Object.assign(cell.font, {
        italic: false,
        strike: false,
      });
    });
  }

  /**
   * Parses the program title from the UACS cell of the Expenditure template.
   *
   * @returns The program title or an empty string if the UACS cell is empty
   */
  private _getProgram(): string | undefined {
    const program = this.getActiveSheet().getCell(
      EXPENDITURE_MATRIX.UACS_CELL,
    ).text;

    return extractProgramTitle(program);
  }
}
