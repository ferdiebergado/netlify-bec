import { DataValidation, Workbook, Worksheet } from 'exceljs';
import {
  BUDGET_ESTIMATE,
  EXPENDITURE_MATRIX,
  EXPENSE_GROUP,
  GAA_OBJECT,
  MANNER_OF_RELEASE,
} from './constants';
import {
  ExpenseGroup,
  GAAObject,
  ExpenseItem,
  MannerOfRelease,
  Activity,
} from './types';
import { extractResult } from './utils';
import config from './config';

const QUANTITY_CELL_INDEX = 8;
const FREQ_CELL_INDEX = 10;
const UNIT_COST_CELL_INDEX = 11;
const PROGRAM_ROW_INDEX = 13;
const OUTPUT_ROW_INDEX = 14;
const ACTIVITY_ROW_INDEX = 15;
const EXPENSE_ITEM_ROW_INDEX = 16;
const AUXILLIARY_SHEETS = ['ContingencyMatrix', 'Venues', 'Honorarium'];

function getCellValueAsNumber(cellValue: string): number {
  const numericValue = +cellValue;
  return isNaN(numericValue) ? 0 : numericValue;
}

function createExpenseItem(
  expenseGroup: ExpenseGroup,
  gaaObject: GAAObject,
  expenseItem: string,
  quantity: number,
  freq: number,
  unitCost: number,
  mannerOfRelease: MannerOfRelease,
  tevLocation: string,
  ppmp: boolean,
  appSupplies: boolean,
  appTicket: boolean,
): ExpenseItem {
  return {
    expenseGroup,
    gaaObject,
    expenseItem,
    quantity,
    freq,
    unitCost,
    mannerOfRelease,
    tevLocation,
    ppmp,
    appSupplies,
    appTicket,
  };
}

function getExpenseItems(
  sheet: Worksheet,
  startRowIndex: number,
  startColIndex: number,
  numRows: number,
  prefix: string,
  mannerOfRelease: MannerOfRelease = MANNER_OF_RELEASE.DIRECT_PAYMENT,
) {
  const items: ExpenseItem[] = [];

  for (let i = 0; i < numRows; i++) {
    const row = sheet.getRow(startRowIndex);

    const quantity = getCellValueAsNumber(
      row.getCell(QUANTITY_CELL_INDEX).text,
    );
    if (quantity === 0) continue;

    const item = row.getCell(startColIndex).text;
    let expenseGroup: ExpenseGroup =
      EXPENSE_GROUP.TRAINING_SCHOLARSHIPS_EXPENSES;
    let gaaObject: GAAObject = GAA_OBJECT.TRAINING_EXPENSES;
    let tevLocation = '';
    let ppmp = false;
    let appSupplies = false;
    let appTicket = false;

    if (item.toLowerCase().includes('supplies')) {
      expenseGroup = EXPENSE_GROUP.SUPPLIES_EXPENSES;
      gaaObject = GAA_OBJECT.OTHER_SUPPLIES;
      appSupplies = true;
    }

    const expenseItem = `${prefix} ${item}`;

    if (
      expenseItem.toLowerCase().includes('travel') &&
      expenseItem.toLowerCase().includes('participants')
    )
      tevLocation = item;

    const freq = getCellValueAsNumber(row.getCell(FREQ_CELL_INDEX).text || '1');
    const unitCost = parseFloat(row.getCell(UNIT_COST_CELL_INDEX).text);

    items.push(
      createExpenseItem(
        expenseGroup,
        gaaObject,
        expenseItem,
        quantity,
        freq,
        unitCost,
        mannerOfRelease,
        tevLocation,
        ppmp,
        appSupplies,
        appTicket,
      ),
    );

    startRowIndex++;
  }

  return items;
}

function boardLodging(sheet: Worksheet) {
  const blPrefix = 'Board and Lodging of';
  const blMannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_BOARD;
  const bl = getExpenseItems(sheet, 17, 3, 4, blPrefix, blMannerOfRelease);
  const blOthers = getExpenseItems(
    sheet,
    24,
    3,
    1,
    blPrefix,
    blMannerOfRelease,
  );
  return [...bl, ...blOthers];
}

function travelExpenses(sheet: Worksheet) {
  const tevPrefix = 'Travel Expenses of Participants from';
  const tevMannerOfRelease = MANNER_OF_RELEASE.FOR_DOWNLOAD_PSF;
  const tevPax = getExpenseItems(
    sheet,
    29,
    4,
    18,
    tevPrefix,
    tevMannerOfRelease,
  );

  const tevPrefixNonPax = 'Travel Expenses of';
  const tevNonPax = getExpenseItems(sheet, 48, 4, 3, tevPrefixNonPax);
  const tevNonPaxOther = getExpenseItems(sheet, 54, 4, 1, tevPrefixNonPax);

  return [...tevPax, ...tevNonPax, ...tevNonPaxOther];
}

function honorarium(sheet: Worksheet) {
  const honorariumPrefix = 'Honorarium of';
  return getExpenseItems(sheet, 58, 3, 2, honorariumPrefix);
}

function otherExpenses(sheet: Worksheet) {
  return getExpenseItems(sheet, 60, 2, 3, '', MANNER_OF_RELEASE.CASH_ADVANCE);
}

function parseActivity(ws: Worksheet) {
  const stDate = ws.getCell(BUDGET_ESTIMATE.CELL_START_DATE).text;
  const month = new Date(stDate).getMonth();
  const bl = boardLodging(ws);
  const tev = travelExpenses(ws);
  const hon = honorarium(ws);
  const others = otherExpenses(ws);

  const info: Activity = {
    // program
    program: ws.getCell(BUDGET_ESTIMATE.CELL_PROGRAM).text,

    // output
    output: ws.getCell(BUDGET_ESTIMATE.CELL_OUTPUT).text,

    // output indicator
    outputIndicator: ws.getCell(BUDGET_ESTIMATE.CELL_OUTPUT_INDICATOR).text,

    // activity
    activityTitle: ws.getCell(BUDGET_ESTIMATE.CELL_ACTIVITY).text,

    // activity indicator
    activityIndicator: ws.getCell(BUDGET_ESTIMATE.CELL_ACTIVITY_INDICATOR).text,

    // month
    month,

    // venue
    venue: ws.getCell(BUDGET_ESTIMATE.CELL_VENUE).text,

    // total pax
    totalPax: extractResult(ws.getCell(BUDGET_ESTIMATE.CELL_TOTAL_PAX).value),

    // output physical target
    outputPhysicalTarget: +ws.getCell(
      BUDGET_ESTIMATE.CELL_ACTIVITY_PHYSICAL_TARGET,
    ).text,

    // activity physical target
    activityPhysicalTarget: +ws.getCell(
      BUDGET_ESTIMATE.CELL_ACTIVITY_PHYSICAL_TARGET,
    ).text,

    // expense items
    expenseItems: [...bl, ...tev, ...hon, ...others],
  };

  return info;
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
function duplicateRows(
  ws: Worksheet,
  targetRowIndex: number,
  srcRowIndex: number,
  numRows: number,
) {
  for (let j = 0; j < numRows; j++) {
    const newRow = ws.insertRow(targetRowIndex, []);
    const srcRow = ws.getRow(srcRowIndex);

    srcRow.eachCell({ includeEmpty: true }, (cell, colNumber) => {
      const targetCell = newRow.getCell(colNumber);

      targetCell.value = cell.value;
      targetCell.style = cell.style;
      targetCell.dataValidation = cell.dataValidation;
    });

    targetRowIndex++;
    srcRowIndex++;
  }
}

function duplicateProgram(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, PROGRAM_ROW_INDEX, 1);
}

function duplicateOutput(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, OUTPUT_ROW_INDEX, 1);
}

function duplicateActivity(ws: Worksheet, targetRow: number) {
  duplicateRows(ws, targetRow, ACTIVITY_ROW_INDEX, 1);
}

function duplicateExpenseItem(ws: Worksheet, targetRow: number, count: number) {
  duplicateRows(ws, targetRow, EXPENSE_ITEM_ROW_INDEX, count);
}

function createActivityRow(
  ws: Worksheet,
  targetRow: number,
  activity: Activity,
  isFirst: boolean = false,
) {
  let activityRowIndex = targetRow;

  if (isFirst) {
    activityRowIndex = ACTIVITY_ROW_INDEX;
  } else {
    duplicateActivity(ws, targetRow);
  }

  const {
    activityTitle,
    activityIndicator,
    activityPhysicalTarget,
    month,
    expenseItems,
  } = activity;

  function sumFormula(cell: string) {
    return {
      formula: `SUM(${cell}${activityRowIndex + 1}:${cell}${
        activityRowIndex + expenseItems.length
      })`,
    };
  }
  const activityRow = ws.getRow(activityRowIndex);

  // activity
  activityRow.getCell('G').value = activityTitle;

  // activity indicator
  activityRow.getCell('H').value = activityIndicator;

  // activity physical target
  activityRow.getCell(
    EXPENDITURE_MATRIX.COL_PHYSICAL_TARGET_MONTH_START_INDEX + month,
  ).value = activityPhysicalTarget;

  // costing grand total
  activityRow.getCell('R').value = sumFormula('R');

  // physical target grand total
  activityRow.getCell('AE').value = {
    formula: `SUM(AF${activityRowIndex}:AQ${activityRowIndex})`,
  };

  // obligation and disbursement grand total per month
  [45, 58].forEach(c => {
    for (let i = 0; i < 12; i++) {
      const cell = activityRow.getCell(c + i);
      const col = cell.$col$row;
      const re = /\$([A-Z]+)\$/;
      const matches = re.exec(col);

      cell.value = sumFormula(matches![1]);
    }
  });

  // obligation grand total
  activityRow.getCell('AR').value = sumFormula('AR');

  // disbursement grand total
  activityRow.getCell('BE').value = sumFormula('BE');
}

function createOutputRow(
  ws: Worksheet,
  targetRow: number,
  activity: Activity,
  rank: number,
  isFirst = false,
) {
  let outputRowIndex = targetRow;

  if (isFirst) {
    outputRowIndex = OUTPUT_ROW_INDEX;
  } else {
    duplicateOutput(ws, targetRow);
  }

  const { output, outputIndicator, outputPhysicalTarget, month } = activity;

  // output
  const outputRow = ws.getRow(outputRowIndex);
  outputRow.getCell('D').value = output;
  outputRow.getCell('E').value = rank;

  // output indicator
  outputRow.getCell('H').value = outputIndicator;

  // output physical target
  outputRow.getCell(
    EXPENDITURE_MATRIX.COL_PHYSICAL_TARGET_MONTH_START_INDEX + month,
  ).value = outputPhysicalTarget;

  // physical target grand total
  outputRow.getCell('AE').value = {
    formula: `SUM(AF${outputRowIndex}:AQ${outputRowIndex})`,
  };
}

export default async function convert(buffers: Buffer[]): Promise<ArrayBuffer> {
  const em = new Workbook();
  await em.xlsx.readFile(config.paths.emTemplate);
  const emWs = em.getWorksheet(1);

  if (!emWs) throw new Error('Sheet not found');

  const Y = 'Y';

  const ynValidation: DataValidation = {
    type: 'list',
    formulae: ['links!$P$1:$P$2'],
  };

  const mannerValidation: DataValidation = {
    type: 'list',
    formulae: ['links!$O$1:$O$5'],
  };

  const programs: string[] = [];
  const activities: Activity[] = [];

  let targetRow = 17;
  let isFirst = true;
  let rank = 1;

  for (const buffer of buffers) {
    const be = new Workbook();
    await be.xlsx.load(buffer);

    be.eachSheet(sheet => {
      const { name } = sheet;

      // skip auxilliary sheets
      if (AUXILLIARY_SHEETS.includes(name)) {
        console.log('skipping', name);
        return;
      }

      // skip if not a budget estimate template
      if (sheet.getCell('C4').text !== 'PROGRAM:') {
        console.log('skipping non budget estimate sheet:', name);
        return;
      }

      console.log('processing', name);

      const activity = parseActivity(sheet);

      activities.push(activity);
    });
  }

  activities
    .sort((a, b) => {
      if (a.program < b.program) {
        return -1;
      } else if (a.program > b.program) {
        return 1;
      } else {
        if (a.output < b.output) {
          return -1;
        } else {
          return 1;
        }
      }
    })
    .forEach(activity => {
      const {
        program,
        output,
        outputIndicator,
        activityTitle,
        activityIndicator,
        month,
        outputPhysicalTarget,
        activityPhysicalTarget,
        expenseItems,
      } = activity;

      let programRowIndex = targetRow;
      // let outputRowIndex = targetRow + 1;

      // program
      if (isFirst) {
        programRowIndex = PROGRAM_ROW_INDEX;
        // outputRowIndex = OUTPUT_ROW_INDEX;
        programs.push(program);
      } else {
        if (!programs.includes(program)) {
          duplicateProgram(emWs, targetRow);
          programs.push(program);
          targetRow++;
        }
      }

      const programRow = emWs.getRow(programRowIndex);
      programRow.getCell('C').value = program;

      createOutputRow(emWs, targetRow, activity, rank, isFirst);

      rank++;

      if (!isFirst) targetRow++;

      createActivityRow(emWs, targetRow, activity, isFirst);

      if (!isFirst) targetRow++;

      // expense items
      duplicateExpenseItem(emWs, targetRow, expenseItems.length);

      for (const expense of expenseItems) {
        let rowIndex = targetRow;

        if (isFirst) rowIndex = targetRow - 1;

        const currentRow = emWs.getRow(rowIndex);

        const {
          expenseGroup,
          gaaObject,
          expenseItem,
          quantity,
          unitCost,
          freq,
          tevLocation,
          ppmp,
          appSupplies,
          appTicket,
          mannerOfRelease,
        } = expense;

        // expense group
        const expenseGroupCell = currentRow.getCell('J');
        expenseGroupCell.value = expenseGroup;

        // gaa object
        const gaaObjectCell = currentRow.getCell('L');
        gaaObjectCell.value = gaaObject;

        // expense item
        currentRow.getCell('N').value = expenseItem;

        // quantity
        currentRow.getCell('O').value = quantity;

        // unit cost
        currentRow.getCell('P').value = unitCost;

        // frequency
        currentRow.getCell('Q').value = freq || 1;

        // total amount
        currentRow.getCell('R').value = {
          formula: `O${rowIndex}*P${rowIndex}*Q${rowIndex}`,
        };

        // tev location
        currentRow.getCell('S').value = tevLocation;

        // ppmp
        const ppmpCell = currentRow.getCell('T');
        ppmpCell.dataValidation = ynValidation;
        if (ppmp) ppmpCell.value = Y;

        // app supplies
        const appSuppliesCell = currentRow.getCell('U');
        appSuppliesCell.dataValidation = ynValidation;
        if (appSupplies) appSuppliesCell.value = Y;

        // app ticket
        const appTicketCell = currentRow.getCell('V');
        appTicketCell.dataValidation = ynValidation;
        if (appTicket) appTicketCell.value = Y;

        // manner of release
        const mannerOfReleaseCell = currentRow.getCell('W');
        mannerOfReleaseCell.value = mannerOfRelease;
        mannerOfReleaseCell.dataValidation = mannerValidation;

        // total obligation
        currentRow.getCell('AR').value = {
          formula: `SUM(AS${rowIndex}:BD${rowIndex})`,
        };

        const totalRef = { formula: `R${rowIndex}` };

        // obligation month
        currentRow.getCell(45 + month).value = totalRef;

        // total disbursement
        currentRow.getCell('BE').value = {
          formula: `SUM(BF${rowIndex}:BQ${rowIndex})`,
        };

        // disbursement month
        currentRow.getCell(58 + month).value = totalRef;

        targetRow++;
      }

      isFirst = false;
      targetRow--;
    });

  emWs.spliceRows(targetRow + 1, 2);

  console.log('writing to buffer...');

  const outBuff = await em.xlsx.writeBuffer();

  return outBuff;
}
