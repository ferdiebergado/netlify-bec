import { Workbook } from './workbook.js';
import {
  EXPENDITURE_MATRIX,
  MANNER_VALIDATION,
  YES,
  YES_NO_VALIDATION,
} from './constants.js';
import type {
  Activity,
  ActivityRowMap,
  ExcelFile,
  ExpenseItemRowContext,
  RowCopyMap,
} from './types/globals.js';
import { BudgetEstimate } from './budgetEstimate.js';
import type { Cell, Row } from 'exceljs';
import { isCellFormulaValue } from './utils.js';

// 0-based index of the last month in a year (December)
const MAX_MONTH = 11;

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
  protected activities: Activity[] = [];

  /**
   * The index of the activity being processed.
   *
   * @protected
   * @type {number}
   */
  protected activityIndex: number = 0;

  /**
   * The current output rank.
   *
   * @protected
   * @type {number}
   */
  protected rank: number = 1;

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
   * @private
   * @param range {RowCopyMap} Contains the indices of the target and source rows that will be duplicated
   *
   * @returns {void}
   */
  private _duplicateRow({
    targetRowIndex,
    srcRowIndex,
    numRows,
  }: RowCopyMap): void {
    const sheet = this.getActiveSheet();
    const srcRow = sheet.getRow(srcRowIndex);
    srcRow.font = {
      bold: false,
      italic: false,
      strike: false,
      underline: 'none',
      size: 11,
      color: {
        argb: 'FF000000',
      },
      name: 'Calibri',
    };

    Array.from({ length: numRows }).forEach((_, j) => {
      const newRow = sheet.insertRow(targetRowIndex + j, []);

      srcRow.eachCell({ includeEmpty: true }, (sourceCell, colNumber) => {
        const targetCell = newRow.getCell(colNumber);
        ExpenditureMatrix._duplicateCell(sourceCell, targetCell);
      });
    });
  }

  private static _duplicateCell(sourceCell: Cell, targetCell: Cell) {
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
    for (let i = 0; i < 12; i += 1) {
      row.getCell(
        EXPENDITURE_MATRIX.PHYSICAL_TARGET_MONTH_COL_INDEX + i,
      ).value = '';
    }
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

    for (let i = 0; i < 12; i += 1) {
      row.getCell(OBLIGATION_MONTH_COL_INDEX + i).value = '';
      row.getCell(DISBURSEMENT_MONTH_COL_INDEX + i).value = '';
    }
  }

  private _fixFonts() {
    for (let i = 1; i <= 12; i++) {
      const row = this.getActiveSheet().getRow(i);

      row.font = {
        size: row.getCell(1).font.size,
        bold: true,
        italic: false,
        strike: false,
        underline: 'none',
        color: {
          argb: 'FF000000',
        },
        name: 'Calibri',
      };
    }
  }

  /**
   * Duplicates the program row in the expenditure matrix.
   *
   * @private
   * @param targetRowIndex {number} The index where the duplicate program row will be inserted.
   *
   * @returns {void}
   */
  private _duplicateProgram(targetRowIndex: number): void {
    const range: RowCopyMap = {
      targetRowIndex,
      srcRowIndex: EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX,
      numRows: 1,
    };

    this._duplicateRow(range);
  }

  /**
   * Duplicates the output row in the expenditure matrix.
   *
   * @private
   * @param targetRowIndex {number} The index where the duplicate output row will be inserted.
   *
   * @returns {void}
   */
  private _duplicateOutput(targetRowIndex: number): void {
    const range: RowCopyMap = {
      targetRowIndex,
      srcRowIndex: EXPENDITURE_MATRIX.OUTPUT_ROW_INDEX,
      numRows: 1,
    };

    this._duplicateRow(range);
  }

  /**
   * Duplicates the activity row in a worksheet.
   *
   * @private
   * @param targetRowIndex {number} The index where the duplicate activity row will be inserted.
   *
   * @returns {void}
   */
  private _duplicateActivity(targetRowIndex: number): void {
    const range: RowCopyMap = {
      targetRowIndex,
      srcRowIndex: EXPENDITURE_MATRIX.ACTIVITY_ROW_INDEX,
      numRows: 1,
    };

    this._duplicateRow(range);
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
  private _duplicateExpenseItem(targetRowIndex: number, copies?: number): void {
    const range: RowCopyMap = {
      targetRowIndex,
      srcRowIndex: EXPENDITURE_MATRIX.EXPENSE_ITEM_ROW_INDEX,
      numRows: copies || 1,
    };

    this._duplicateRow(range);
  }

  /**
   * Checks if the current activity is the first activity.
   *
   * @returns {boolean}
   */
  private _hasFirstActivity(): boolean {
    return this.activityIndex === 0;
  }

  private _setIsBlankFormula(rowIndex: number) {
    const sheet = this.getActiveSheet();

    interface FormulaCells {
      formulaCell: string;
      count: number;
    }

    const formulaCells: FormulaCells[] = [
      { formulaCell: EXPENDITURE_MATRIX.IS_BLANK_FORMULA_CELL1, count: 2 },
      { formulaCell: EXPENDITURE_MATRIX.IS_BLANK_FORMULA_CELL2, count: 2 },
      { formulaCell: EXPENDITURE_MATRIX.IS_BLANK_FORMULA_CELL3, count: 1 },
    ];

    let colNum = EXPENDITURE_MATRIX.IS_BLANK_FORMULA_START_COL;
    formulaCells.forEach(cell => {
      const sourceCell = sheet.getCell(cell.formulaCell);

      for (let i = 0; i < cell.count; i++) {
        const targetCell = sheet.getRow(rowIndex).getCell(colNum);

        ExpenditureMatrix._duplicateCell(sourceCell, targetCell);

        colNum += 1;
      }
    });
  }

  /**
   * Creates or duplicates an activity row in the expenditure matrix.
   *
   * @param targetRow {number} The index where the activity row will be inserted.
   * @param activity {Activity} The activity information.
   *
   * @returns number - The index of the activity row
   */
  private _createActivityRow(targetRow: number, activity: Activity): number {
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

    if (this._hasFirstActivity()) {
      activityRowIndex = ACTIVITY_ROW_INDEX;
      const activityRow = sheet.getRow(activityRowIndex);

      // Expense object formula
      let targetCell = activityRow.getCell(
        EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_COL,
      );

      ExpenditureMatrix._duplicateCell(
        sheet.getCell(EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_CELL),
        targetCell,
      );

      // Isblank formula
      this._setIsBlankFormula(activityRowIndex);
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

    if (!this._hasFirstActivity()) {
      ExpenditureMatrix._clearPreviousPhysicalTargets(activityRow);
    }

    // activity physical target
    const targetMonth = ExpenditureMatrix._incrementMonth(month);

    activityRow.getCell(PHYSICAL_TARGET_MONTH_COL_INDEX + targetMonth).value =
      activityPhysicalTarget;

    // costing grand total
    const totalCostCell = activityRow.getCell(TOTAL_COST_COL);
    totalCostCell.value = sumFormula(TOTAL_COST_COL);
    totalCostCell.font.bold = true;

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
   *
   * @returns {void}
   */
  private _createOutputRow(targetRowIndex: number, activity: Activity): void {
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

    if (this._hasFirstActivity()) {
      outputRowIndex = OUTPUT_ROW_INDEX;
    } else {
      this._duplicateOutput(outputRowIndex);
    }

    const { output, outputIndicator, outputPhysicalTarget, month } =
      activity.info;

    console.table(activity.info);
    console.log('outputrowindex:', outputRowIndex);

    // output
    const outputRow = sheet.getRow(outputRowIndex);

    console.log('outputrow:', outputRow.number);

    outputRow.getCell(OUTPUT_COL).value = output;
    outputRow.getCell(RANK_COL).value = this.rank;

    // output indicator
    outputRow.getCell(PERFORMANCE_INDICATOR_COL).value = outputIndicator;

    if (!this._hasFirstActivity()) {
      ExpenditureMatrix._clearPreviousPhysicalTargets(outputRow);
    }

    // output physical target
    const targetMonth = ExpenditureMatrix._incrementMonth(month);

    outputRow.getCell(PHYSICAL_TARGET_MONTH_COL_INDEX + targetMonth).value =
      outputPhysicalTarget;

    // physical target grand total
    const physicalTargetMonthStartCell = `${PHYSICAL_TARGET_MONTH_START_COL_INDEX}${outputRowIndex}`;
    const physicalTargetMonthEndCell = `${PHYSICAL_TARGET_MONTH_END_COL_INDEX}${outputRowIndex}`;

    outputRow.getCell(PHYSICAL_TARGET_TOTAL_COL).value = {
      formula: `SUM(${physicalTargetMonthStartCell}:${physicalTargetMonthEndCell})`,
    };
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
   * Creates an expense item row in the expenditure matrix.
   *
   * @private
   * @param context {ExpenseItemRowContext} The context of the expense item row will be inserted.
   *
   * @returns {void}
   */
  private _createExpenseItemRow({
    targetRowIndex,
    expense,
    month,
  }: ExpenseItemRowContext): void {
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

    if (this._hasFirstActivity()) {
      currentRowIndex -= 1;
      const currentRow = sheet.getRow(currentRowIndex);

      // Expense object formula
      let targetCell = currentRow.getCell(
        Number(EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_COL),
      );

      ExpenditureMatrix._duplicateCell(
        sheet.getCell(EXPENDITURE_MATRIX.EXPENSE_OBJECT_FORMULA_CELL),
        targetCell,
      );

      // GAA Object formula
      targetCell = currentRow.getCell(
        Number(EXPENDITURE_MATRIX.GAA_OBJECT_FORMULA_COL),
      );

      Object.assign(targetCell, {
        value: {
          formula: `VLOOKUP(L${currentRowIndex},'links'!K$2:L$222,2,FALSE())`,
        },
      });

      // Is blank formulas
      this._setIsBlankFormula(currentRowIndex);
    }

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

    if (!this._hasFirstActivity()) {
      ExpenditureMatrix._clearPreviousFinancialPrograms(currentRow);
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
    this._removeExtraRows();
    this._fixFonts();

    const sheet = this.getActiveSheet();

    await Promise.all(files.map(file => this._addToActivities(file)));

    if (this.activities.length === 0) throw new Error('No activities found.');

    this.activities.sort(ExpenditureMatrix._orderByProgramAndOutput);

    let currentRowIndex: number = EXPENDITURE_MATRIX.TARGET_ROW_INDEX;
    let currentOutput = '';
    let currentProgram = '';

    /**
     * Records the indices of rows of each activity that contains the total unit cost
     * to be used later to compute the grand total
     */
    const activityRows: number[] = [];

    const psf: Activity = {
      info: {
        program: 'Programs Support Funds (PSF)',
        output: 'Benefitted implementers',
        outputIndicator: 'No. of implementers benefitted',
        outputPhysicalTarget: 16,
        activityTitle: 'Provision of Program Support Funds',
        activityIndicator: 'No. of downloading activities conducted',
        activityPhysicalTarget: 1,
        month: 1,
        venue: '',
        totalPax: 16,
      },
      expenseItems: [],
      tevPSF: [],
    };

    this.activities.forEach(activity => {
      const {
        info: { program, output },
        tevPSF,
      } = activity;

      // program
      let programRowIndex;

      if (this._hasFirstActivity()) {
        programRowIndex = EXPENDITURE_MATRIX.PROGRAM_ROW_INDEX;
        console.log(
          'writing program at row',
          programRowIndex,
          'program:',
          program,
        );

        const programRow = sheet.getRow(programRowIndex);
        programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;
      } else {
        programRowIndex = currentRowIndex;

        console.log('program:', program, 'currentProgram', currentProgram);
        if (program !== currentProgram) {
          console.log(
            'creating program at index',
            programRowIndex,
            'program:',
            program,
          );

          this._duplicateProgram(programRowIndex);

          const programRow = sheet.getRow(programRowIndex);
          programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;
        } else {
          currentRowIndex -= 1;
          console.log('skipping program row');
        }
      }

      currentProgram = program;

      if (!this._hasFirstActivity()) {
        currentRowIndex += 1;
        console.log('moved current row to', currentRowIndex);
      }

      // output
      console.log('output:', output, 'currentOutput:', currentOutput);

      if (output !== currentOutput) {
        console.log(
          'creating output row at index',
          currentRowIndex,
          'hasEmptyActivities',
          this._hasFirstActivity(),
          'output',
          output,
        );
        this._createOutputRow(currentRowIndex, activity);
        this.rank += 1;
        console.log('moved current row to', currentRowIndex);
      } else {
        currentRowIndex -= 1;
        console.log('skipping output row');
      }

      currentOutput = output;

      if (!this._hasFirstActivity()) {
        currentRowIndex += 1;
        console.log('moved current row to', currentRowIndex);
      }

      // activity
      console.log(
        'creating activity row at index',
        currentRowIndex,
        'activity:',
        activity,
      );

      const activityRowIndex = this._createActivityRow(
        currentRowIndex,
        activity,
      );

      activityRows.push(activityRowIndex);

      if (!this._hasFirstActivity()) {
        currentRowIndex += 1;
        console.log('moved current row to', currentRowIndex);
      }

      console.log(
        'hasFirstActivity:',
        this._hasFirstActivity(),
        'current row index:',
        currentRowIndex,
      );

      // Expense Items
      const activityRowMap: ActivityRowMap = {
        activity,
        rowIndex: currentRowIndex,
      };

      currentRowIndex = this._createExpenseItems(activityRowMap);

      // Aggregate TEVs of participants for PSF
      tevPSF.forEach(tev => {
        const existingPSF = psf.expenseItems.find(
          i => i.expenseItem === tev.expenseItem,
        );

        if (existingPSF) {
          existingPSF.unitCost += tev.unitCost * tev.quantity;
        } else {
          psf.expenseItems.push({
            ...tev,
            quantity: 1,
            unitCost: tev.unitCost * tev.quantity,
          });
        }
      });

      if (this._hasFirstActivity()) {
        currentRowIndex -= 1;
      }

      this.activityIndex += 1;

      console.log('activity created');
      console.log('currentrowindex:', currentRowIndex);
    });

    // PSF
    if (psf.expenseItems.length > 0) {
      console.log('Creating PSF Activity...');

      // program
      const { program } = psf.info;

      this._duplicateProgram(currentRowIndex);

      const programRow = sheet.getRow(currentRowIndex);
      programRow.getCell(EXPENDITURE_MATRIX.PROGRAM_COL).value = program;

      currentRowIndex += 1;
      console.log('moved current row to', currentRowIndex);

      // output
      this._createOutputRow(currentRowIndex, psf);

      currentRowIndex += 1;
      console.log('moved current row to', currentRowIndex);

      // activity
      const activityRowIndex = this._createActivityRow(currentRowIndex, psf);

      activityRows.push(activityRowIndex);

      currentRowIndex += 1;
      console.log('moved current row to', currentRowIndex);

      const activityRowMap: ActivityRowMap = {
        activity: psf,
        rowIndex: currentRowIndex,
      };

      // expense items
      currentRowIndex = this._createExpenseItems(activityRowMap);
    }

    sheet.spliceRows(currentRowIndex, 2);

    const lastRowIndex = currentRowIndex + 1;

    // GRAND TOTAL
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
   * Creates the expense items at the specified row on the active sheet.
   *
   * @private
   * @param context {ActivityContext} The context of the Activity
   *
   * @returns {number} - The new index for the current row
   */
  private _createExpenseItems({ activity, rowIndex }: ActivityRowMap): number {
    const { expenseItems } = activity;

    const finalRowIndex = expenseItems.reduce((rowIndex, expense, index) => {
      console.dir(expense);
      console.log(
        'current expense item index:',
        index,
        'expenseitems.length:',
        expenseItems.length,
      );

      console.log('duplicating expense item at row', rowIndex);
      this._duplicateExpenseItem(rowIndex);

      console.log('creating expense row at index', rowIndex);

      const context: ExpenseItemRowContext = {
        targetRowIndex: rowIndex,
        expense,
        month: activity.info.month,
      };

      this._createExpenseItemRow(context);

      console.log('expenses item created.');
      return rowIndex + 1;
    }, rowIndex);

    return finalRowIndex;
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

  private _removeExtraRows() {
    const {
      SPLICE_START_ROW,
      SPLICE_NUM_ROWS,
      MILESTONES_START_ROW,
      MILESTONES_NUM_ROWS,
    } = EXPENDITURE_MATRIX;
    const sheet = this.getActiveSheet();
    sheet.spliceRows(SPLICE_START_ROW, SPLICE_NUM_ROWS);
    sheet.spliceRows(MILESTONES_START_ROW, MILESTONES_NUM_ROWS);
  }
}
