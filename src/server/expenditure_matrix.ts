import Worksheet from './worksheet';
import BudgetEstimate from './budget_estimate';
import { ActivityInfo, ExpenseItem, CellData } from './types';
import { EXPENDITURE_MATRIX } from './constants';

class ExpenditureMatrix extends Worksheet {
  beMonth?: number;

  // eslint-disable-next-line @typescript-eslint/no-useless-constructor
  constructor(xls: ArrayBuffer, sheet: string) {
    super(xls, sheet);
  }

  addInfo(info: ActivityInfo) {
    if (this.ws) {
      const { program, output, outputIndicator, activity, activityIndicator } =
        info;
      const data: CellData[] = [
        {
          cell: EXPENDITURE_MATRIX.CELL_PROGRAM,
          value: program,
        },
        {
          cell: EXPENDITURE_MATRIX.CELL_OUTPUT,
          value: output,
        },
        {
          cell: EXPENDITURE_MATRIX.CELL_OUTPUT_INDICATOR,
          value: outputIndicator,
        },
        { cell: EXPENDITURE_MATRIX.CELL_ACTIVITY, value: activity },
        {
          cell: EXPENDITURE_MATRIX.CELL_ACTIVITY_INDICATOR,
          value: activityIndicator,
        },
      ];

      for (const d of data) {
        const { cell, value } = d;

        this.ws.getCell(cell).value = value;
      }
    }

    this.beMonth = info.month;
  }

  addExpenseItems(expenseItems: ExpenseItem[]) {
    const updateGrandTotals = (startRow: number, endRow: number) => {
      const grandCols = [
        EXPENDITURE_MATRIX.COL_TOTAL_COST as string,
        EXPENDITURE_MATRIX.COL_TOTAL_OBLIGATION as string,
        EXPENDITURE_MATRIX.COL_TOTAL_DISBURSEMENT as string,
      ];

      for (const col of grandCols) {
        const grandTotal = {
          formula: `sum(${col}${startRow}:${col}${endRow - 1})`,
        };

        if (this.ws) this.ws.getCell(col + (startRow - 1)).value = grandTotal;
      }
    };

    if (this.ws) {
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

      this.ws.duplicateRow(startRow, nRows, true);

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
          formula: `sum(${EXPENDITURE_MATRIX.COL_OBLIGATION_MONTH_START}${currentRow}:${EXPENDITURE_MATRIX.COL_OBLIGATION_MONTH_END}${currentRow})`,
        };
        const totalDisbursement = {
          formula: `sum(${EXPENDITURE_MATRIX.COL_DISBURSEMENT_MONTH_START}${currentRow}:${EXPENDITURE_MATRIX.COL_DISBURSEMENT_MONTH_END}${currentRow})`,
        };

        const data: CellData[] = [
          {
            cell: EXPENDITURE_MATRIX.COL_EXPENSE_GROUP as string,
            value: expenseGroup,
            dataValidation: expenseGrpValidation,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_GAA_OBJECT as string,
            value: gaaObject,
            dataValidation: gaaObjValidation,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_EXPENSE_ITEM as string,
            value: expenseItem,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_QUANTITY as string,
            value: quantity,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_FREQUENCY as string,
            value: freq,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_UNIT_COST as string,
            value: unitCost,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_TOTAL_COST as string,
            value: totalCost,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_PPMP as string,
            value: ppmp,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_APP_SUPPLIES as string,
            value: appSupplies,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_APP_TICKET as string,
            value: appTicket,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_MANNER_OF_RELEASE as string,
            value: mannerOfRelease,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_TOTAL_OBLIGATION as string,
            value: totalObligation,
          },
          {
            cell: EXPENDITURE_MATRIX.COL_TOTAL_DISBURSEMENT as string,
            value: totalDisbursement,
          },
        ];

        for (const d of data) {
          const { cell, value } = d;
          const currentCell = this.ws.getCell(cell + currentRow);

          currentCell.value = value;

          if (d.dataValidation) currentCell.dataValidation = d.dataValidation;
        }

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

        currentRow++;
      }

      updateGrandTotals(startRow, currentRow);
    }
  }

  apply(be: BudgetEstimate) {
    const { activityInfo, expenseItems } = be;

    this.addInfo(activityInfo);
    this.addExpenseItems(expenseItems);
  }

  async save() {
    const buffer = await this.wb.xlsx.writeBuffer();

    return buffer;
  }
}

export default ExpenditureMatrix;
