import { CellValue, DataValidation } from 'exceljs';

enum Defaults {
  BE_SHEET_NAME = 'BE-001',
  EM_SHEET_NAME = 'Expenditure Form',
}

enum BudgetEstimateCell {
  PROGRAM = 'F4',
  OUTPUT = 'F5',
  OUTPUT_INDICATOR = 'F6',
  ACTIVITY = 'F7',
  ACTIVITY_INDICATOR = 'F8',
  START_DATE = 'G13',
  VENUE = 'O13',
  TOTAL_PAX = 'H16',
}

enum BudgetEstimateRow {
  BOARD_LODGING_START = 17,
  BOARD_LODGING_END = 20,
  BOARD_LODGING_OTHER = 24,
  TRAVEL_REGION_START = 29,
  TRAVEL_REGION_END = 47,
  TRAVEL_CO_START = 48,
  TRAVEL_CO_END = 50,
  TRAVEL_OTHER = 54,
  HONORARIUM_START = 58,
  HONORARIUM_END = 59,
  SUPPLIES_CONTINGENCY_START = 60,
  SUPPLIES_CONTINGENCY_END = 61,
}

enum BudgetEstimateCol {
  BOARD_LODGING = 'C',
  TRAVEL_REGION = 'D',
  TRAVEL_CO = 'C',
  TRAVEL_OTHER = 'C',
  HONORARIUM = 'C',
  SUPPLIES_CONTINGENCY = 'B',
}

enum ExpenditureMatrixCell {
  PROGRAM = 'C13',
  OUTPUT = 'D14',
  OUTPUT_INDICATOR = 'H14',
  ACTIVITY = 'G17',
  ACTIVITY_INDICATOR = 'H17',
}

enum ExpenditureMatrixCol {
  EXPENSE_GROUP = 'J',
  GAA_OBJECT = 'L',
  EXPENSE_ITEM = 'N',
  QUANTITY = 'O',
  FREQUENCY = 'Q',
  UNIT_COST = 'P',
  TOTAL_COST = 'R',
  PPMP = 'T',
  APP_SUPPLIES = 'U',
  APP_TICKET = 'V',
  MANNER_OF_RELEASE = 'W',
  TOTAL_OBLIGATION = 'AR',
  TOTAL_DISBURSEMENT = 'BE',
  OBLIGATION_MONTH_START_INDEX = 45,
  DISBURSEMENT_MOTH_START_INDEX = 58,
  PHYSICAL_TARGET_MONTH_START_INDEX = 32,
}

enum ExpenditureMatrixRow {
  EXPENSE_ITEM_START_ROW = 18,
  EXISTING_EXPENSE_ITEM_ROWS = 4,
}

enum ExpensePrefix {
  BOARD_LODGING = 'Board and Lodging of ',
  TRAVEL = 'Travel Expenses of ',
  HONORARIUM = 'Honorarium of ',
}

enum ExpenseGroup {
  TRAINING_SCHOLARSHIPS_EXPENSES = 'Training and Scholarship Expenses',
  SUPPLIES_EXPENSES = 'Supplies and Materials Expenses',
}

enum GAAObject {
  TRAINING_EXPENSES = 'Training Expenses',
  OTHER_SUPPLIES = 'Other Supplies and Materials Expenses',
}

enum MannerOfRelease {
  FOR_DOWNLOAD_BOARD = 'For Downloading (Board and Lodging)',
  FOR_DOWNLOAD_PSF = 'For Downloading (Program Support Funds)',
  DIRECT_PAYMENT = 'Direct Payment',
  CASH_ADVANCE = 'Cash Advance',
}

interface ExpenseItem {
  expenseGroup?: ExpenseGroup;
  gaaObject?: GAAObject;
  expenseItem: string;
  quantity: number;
  freq: number;
  unitCost: number;
  ppmp?: YesNo;
  appSupplies: YesNo;
  appTicket: YesNo;
  mannerOfRelease: MannerOfRelease;
}

interface ActivityInfo {
  program?: string;
  output?: string;
  outputIndicator?: string;
  activity?: string;
  activityIndicator?: string;
  month?: number;
}

interface CellData {
  cell: string;
  value: CellValue;
  dataValidation?: DataValidation;
}

type YesNo = 'Y' | 'N';

export {
  Defaults,
  BudgetEstimateCell,
  BudgetEstimateRow,
  BudgetEstimateCol,
  ExpensePrefix,
  YesNo,
  ExpenseGroup,
  GAAObject,
  MannerOfRelease,
  ExpenseItem,
  ExpenditureMatrixCell,
  ExpenditureMatrixRow,
  ExpenditureMatrixCol,
  ActivityInfo,
  CellData,
};
