import Worksheet from "./worksheet";
import BudgetEstimate from "./budget_estimate";
import {
  ActivityInfo,
  ExpenditureMatrixCell,
  ExpenditureMatrixRow,
  ExpenditureMatrixCol,
  ExpenseItem,
  CellData,
} from "./types";

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
          cell: ExpenditureMatrixCell.PROGRAM,
          value: program,
        },
        {
          cell: ExpenditureMatrixCell.OUTPUT,
          value: output,
        },
        {
          cell: ExpenditureMatrixCell.OUTPUT_INDICATOR,
          value: outputIndicator,
        },
        { cell: ExpenditureMatrixCell.ACTIVITY, value: activity },
        {
          cell: ExpenditureMatrixCell.ACTIVITY_INDICATOR,
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
        ExpenditureMatrixCol.TOTAL_COST as string,
        ExpenditureMatrixCol.TOTAL_OBLIGATION as string,
        ExpenditureMatrixCol.TOTAL_DISBURSEMENT as string,
      ];

      for (const col of grandCols) {
        const grandTotal = {
          formula: `sum(${col}${startRow}:${col}${endRow - 1})`,
        };

        if (this.ws) this.ws.getCell(col + (startRow - 1)).value = grandTotal;
      }
    };

    if (this.ws) {
      const startRow = ExpenditureMatrixRow.EXPENSE_ITEM_START_ROW;
      const nRows =
        expenseItems.length - ExpenditureMatrixRow.EXISTING_EXPENSE_ITEM_ROWS;
      let currentRow = startRow;

      const expenseGrpValidation = this.ws.getCell(
        ExpenditureMatrixCol.EXPENSE_GROUP + startRow
      ).dataValidation;

      const gaaObjValidation = this.ws.getCell(
        ExpenditureMatrixCol.GAA_OBJECT + startRow
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
              ExpenditureMatrixCol.PHYSICAL_TARGET_MONTH_START_INDEX +
                this.beMonth
            ).value = 1;

        const totalCost = {
          formula: `O${currentRow}*P${currentRow}*Q${currentRow}`,
        };
        const totalObligation = {
          formula: `sum(AS${currentRow}:BD${currentRow})`,
        };
        const totalDisbursement = {
          formula: `sum(BF${currentRow}:BQ${currentRow})`,
        };

        const data: CellData[] = [
          {
            cell: ExpenditureMatrixCol.EXPENSE_GROUP as string,
            value: expenseGroup,
            dataValidation: expenseGrpValidation,
          },
          {
            cell: ExpenditureMatrixCol.GAA_OBJECT as string,
            value: gaaObject,
            dataValidation: gaaObjValidation,
          },
          {
            cell: ExpenditureMatrixCol.EXPENSE_ITEM as string,
            value: expenseItem,
          },
          {
            cell: ExpenditureMatrixCol.QUANTITY as string,
            value: quantity,
          },
          {
            cell: ExpenditureMatrixCol.FREQUENCY as string,
            value: freq,
          },
          {
            cell: ExpenditureMatrixCol.UNIT_COST as string,
            value: unitCost,
          },
          {
            cell: ExpenditureMatrixCol.TOTAL_COST as string,
            value: totalCost,
          },
          {
            cell: ExpenditureMatrixCol.PPMP as string,
            value: ppmp,
          },
          {
            cell: ExpenditureMatrixCol.APP_SUPPLIES as string,
            value: appSupplies,
          },
          {
            cell: ExpenditureMatrixCol.APP_TICKET as string,
            value: appTicket,
          },
          {
            cell: ExpenditureMatrixCol.MANNER_OF_RELEASE as string,
            value: mannerOfRelease,
          },
          {
            cell: ExpenditureMatrixCol.TOTAL_OBLIGATION as string,
            value: totalObligation,
          },
          {
            cell: ExpenditureMatrixCol.TOTAL_DISBURSEMENT as string,
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
            ExpenditureMatrixCol.OBLIGATION_MONTH_START_INDEX as number,
            ExpenditureMatrixCol.DISBURSEMENT_MOTH_START_INDEX as number,
          ];
          const row = this.ws.getRow(currentRow);

          for (const col of cols) {
            const monthCol = col + this.beMonth;

            row.getCell(monthCol).value = {
              formula: ExpenditureMatrixCol.TOTAL_COST + currentRow,
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
    // const filename = `expenditure-${timestamp()}.xlsx`;
    // const outputFile = path.join(this.workDir, filename);

    // await this.wb.xlsx.writeFile(outputFile);

    return await this.wb.xlsx.writeBuffer();
  }
}

export default ExpenditureMatrix;
