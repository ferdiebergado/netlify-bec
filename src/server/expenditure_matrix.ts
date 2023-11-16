import Worksheet from './worksheet';
import BudgetEstimate from './budget_estimate';
import type { ActivityInfo, ExpenseItem, CellData } from './types';
import { EXPENDITURE_MATRIX } from './constants';

class ExpenditureMatrix extends Worksheet {
  beMonth?: number;

  // eslint-disable-next-line @typescript-eslint/no-useless-constructor
  constructor(xls: ArrayBuffer, sheet: string = EXPENDITURE_MATRIX.SHEET_NAME) {
    super(xls, sheet);
  }

  _addInfo(info: ActivityInfo): void {
    if (!this.ws) throw new Error(ExpenditureMatrix.LOAD_ERROR_MSG);

    const { program, output, outputIndicator, activity, activityIndicator } =
      info;
    const data: CellData[] = [
      {
        // program
        cell: EXPENDITURE_MATRIX.CELL_PROGRAM,
        value: program,
      },
      {
        // output
        cell: EXPENDITURE_MATRIX.CELL_OUTPUT,
        value: output,
      },
      {
        // output indicator
        cell: EXPENDITURE_MATRIX.CELL_OUTPUT_INDICATOR,
        value: outputIndicator,
      },
      {
        // activity
        cell: EXPENDITURE_MATRIX.CELL_ACTIVITY,
        value: activity,
      },
      {
        // activity indicator
        cell: EXPENDITURE_MATRIX.CELL_ACTIVITY_INDICATOR,
        value: activityIndicator,
      },
    ];

    // set the respective info
    for (const d of data) {
      const { cell, value } = d;

      this.ws.getCell(cell).value = value;
    }

    // track the activity month
    this.beMonth = info.month;
  }

  _addExpenseItems(expenseItems: ExpenseItem[]): void {
    if (!this.ws) throw new Error(ExpenditureMatrix.LOAD_ERROR_MSG);

    const startRow = EXPENDITURE_MATRIX.ROW_EXPENSE_ITEM_START_ROW;
    const nRows =
      expenseItems.length - EXPENDITURE_MATRIX.ROW_EXISTING_EXPENSE_ITEM_ROWS;
    let currentRow = startRow;

    const expenseGrpValidation = this.ws.getCell(
      EXPENDITURE_MATRIX.COL_EXPENSE_GROUP + startRow,
    ).dataValidation;

    const gaaObjValidation = this.ws.getCell(
      EXPENDITURE_MATRIX.COL_GAA_OBJECT + startRow,
    ).dataValidation;

    // add more rows for the expense items
    this.ws.duplicateRow(startRow, nRows, true);

    // iterate through each expense item
    for (const item of expenseItems) {
      const {
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
      } = item;

      // set the physical target at the activity month column
      if (this.beMonth)
        this.ws
          .getRow(startRow - 1)
          .getCell(
            EXPENDITURE_MATRIX.COL_PHYSICAL_TARGET_MONTH_START_INDEX +
              this.beMonth,
          ).value = 1;

      const totalCost = {
        formula: `${EXPENDITURE_MATRIX.COL_COSTING_QUANTITY}${currentRow}*${EXPENDITURE_MATRIX.COL_COSTING_UNIT_COST}${currentRow}*${EXPENDITURE_MATRIX.COL_COSTING_FREQUENCY}${currentRow}`,
      };
      const totalObligation = {
        formula: `SUM(${EXPENDITURE_MATRIX.COL_OBLIGATION_MONTH_START}${currentRow}:${EXPENDITURE_MATRIX.COL_OBLIGATION_MONTH_END}${currentRow})`,
      };
      const totalDisbursement = {
        formula: `SUM(${EXPENDITURE_MATRIX.COL_DISBURSEMENT_MONTH_START}${currentRow}:${EXPENDITURE_MATRIX.COL_DISBURSEMENT_MONTH_END}${currentRow})`,
      };

      const data: CellData[] = [
        {
          // expense group
          cell: EXPENDITURE_MATRIX.COL_EXPENSE_GROUP as string,
          value: expenseGroup,
          dataValidation: expenseGrpValidation,
        },
        {
          // gaa object
          cell: EXPENDITURE_MATRIX.COL_GAA_OBJECT as string,
          value: gaaObject,
          dataValidation: gaaObjValidation,
        },
        {
          // expense item
          cell: EXPENDITURE_MATRIX.COL_EXPENSE_ITEM as string,
          value: expenseItem,
        },
        {
          // quantity
          cell: EXPENDITURE_MATRIX.COL_QUANTITY as string,
          value: quantity,
        },
        {
          // frequency
          cell: EXPENDITURE_MATRIX.COL_FREQUENCY as string,
          value: freq,
        },
        {
          // unit cost
          cell: EXPENDITURE_MATRIX.COL_UNIT_COST as string,
          value: unitCost,
        },
        {
          // total cost
          cell: EXPENDITURE_MATRIX.COL_TOTAL_COST as string,
          value: totalCost,
        },
        {
          // ppmp
          cell: EXPENDITURE_MATRIX.COL_PPMP as string,
          value: ppmp,
        },
        {
          // app supplies
          cell: EXPENDITURE_MATRIX.COL_APP_SUPPLIES as string,
          value: appSupplies,
        },
        {
          // app ticket
          cell: EXPENDITURE_MATRIX.COL_APP_TICKET as string,
          value: appTicket,
        },
        {
          // manner of release
          cell: EXPENDITURE_MATRIX.COL_MANNER_OF_RELEASE as string,
          value: mannerOfRelease,
        },
        {
          // total obligation
          cell: EXPENDITURE_MATRIX.COL_TOTAL_OBLIGATION as string,
          value: totalObligation,
        },
        {
          // total disbursement
          cell: EXPENDITURE_MATRIX.COL_TOTAL_DISBURSEMENT as string,
          value: totalDisbursement,
        },
      ];

      // set each cell to the respective values
      for (const d of data) {
        const { cell, value } = d;
        const currentCell = this.ws.getCell(cell + currentRow);

        currentCell.value = value;

        if (d.dataValidation) currentCell.dataValidation = d.dataValidation;
      }

      // create a reference the total cost at the corresponding obligation and disbursement month column
      if (this.beMonth) {
        const cols = [
          EXPENDITURE_MATRIX.COL_OBLIGATION_MONTH_START_INDEX as number,
          EXPENDITURE_MATRIX.COL_DISBURSEMENT_MONTH_START_INDEX as number,
        ];
        const row = this.ws.getRow(currentRow);

        for (const col of cols) {
          const monthCol = col + this.beMonth;

          row.getCell(monthCol).value = {
            formula: EXPENDITURE_MATRIX.COL_TOTAL_COST + currentRow,
          };
        }
      }

      // move to the next row
      currentRow++;
    }

    // Update the grand total columns
    this._updateGrandTotals(startRow, currentRow);
  }

  _updateGrandTotals(startRow: number, endRow: number): void {
    const grandCols = [
      EXPENDITURE_MATRIX.COL_TOTAL_COST as string,
      EXPENDITURE_MATRIX.COL_TOTAL_OBLIGATION as string,
      EXPENDITURE_MATRIX.COL_TOTAL_DISBURSEMENT as string,
    ];

    for (const col of grandCols) {
      const grandTotal = {
        formula: `SUM(${col}${startRow}:${col}${endRow - 1})`,
      };

      this.ws!.getCell(col + (startRow - 1)).value = grandTotal;
    }
  }

  apply(be: BudgetEstimate): void {
    const { activityInfo, expenseItems } = be;

    this._addInfo(activityInfo);
    this._addExpenseItems(expenseItems);
  }

  async save(): Promise<ArrayBuffer> {
    const buffer = await this.wb.xlsx.writeBuffer();

    return buffer;
  }
}

export default ExpenditureMatrix;
