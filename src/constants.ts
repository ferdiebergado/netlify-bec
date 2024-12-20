import { DataValidation } from 'exceljs';

/**
 * Constant representing the value 'Y'.
 */
const YES = 'Y';

/**
 * MIME type for Excel files.
 */
const EXCEL_MIMETYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

/**
 * List of auxiliary sheets that should be skipped during processing.
 */
const AUXILLIARY_SHEETS = ['ContingencyMatrix', 'Venues', 'Honorarium'];

const BOARD_LODGING_EXPENSE_PREFIX = 'Board and Lodging of';
const HONORARIUM_EXPENSE_PREFIX = 'Honorarium of';
const TRAVEL_EXPENSE_PREFIX = 'Travel Expenses of';

/**
 * Constants related to the budget estimate structure.
 */
const BUDGET_ESTIMATE = {
  PROGRAM_HEADING_CELL: 'C4',
  PROGRAM_CELL: 'F4',
  OUTPUT_CELL: 'F5',
  OUTPUT_INDICATOR_CELL: 'F6',
  ACTIVITY_CELL: 'F7',
  ACTIVITY_INDICATOR_CELL: 'F8',
  START_DATE_CELL: 'G13',
  NUM_DAYS_CELL: 'M13',
  VENUE_CELL: 'O13',
  TOTAL_PAX_CELL: 'H16',
  BOARD_LODGING_TOTAL_PAX_CELL: 'H16',
  BOARD_LODGING_UNIT_COST_CELL: 'K17',
  OUTPUT_PHYSICAL_TARGET_CELL: 'O6',
  ACTIVITY_PHYSICAL_TARGET_CELL: 'O8',
  EXPENSE_ITEM_COL_INDEX: 2,
  EXPENSE_ITEM_FIRST_COL_INDEX: 3,
  EXPENSE_ITEM_SECOND_COL_INDEX: 4,
  QUANTITY_CELL_INDEX: 8,
  FREQ_CELL_INDEX: 10,
  UNIT_COST_CELL_INDEX: 11,
  BOARD_LODGING_START_ROW_INDEX: 17,
  BOARD_LODGING_OTHER_ROW_INDEX: 24,
  TRAVEL_REGION_ROW_INDEX: 29,
  TRAVEL_CO_ROW_INDEX: 48,
  TRAVEL_OTHER_ROW_INDEX: 54,
  HONORARIUM_ROW_INDEX: 58,
  MEAL_EXPENSES_ROW_INDEX: 60,
} as const;

/**
 * Constants related to the expenditure matrix structure.
 */
const EXPENDITURE_MATRIX = {
  SPLICE_START_ROW: 20,
  SPLICE_NUM_ROWS: 20,
  MILESTONES_START_ROW: 15,
  MILESTONES_NUM_ROWS: 2,
  EXPENSE_OBJECT_FORMULA_COL: 11,
  EXPENSE_OBJECT_FORMULA_CELL: 'K13',
  GAA_OBJECT_FORMULA_COL: '13',
  GAA_OBJECT_FORMULA_CELL: 'M16',
  IS_BLANK_FORMULA_CELL1: 'BY13',
  IS_BLANK_FORMULA_CELL2: 'CA13',
  IS_BLANK_FORMULA_CELL3: 'CC13',
  IS_BLANK_FORMULA_START_COL: 77,
  OVERHEAD_NUM_ROWS: 32,
  PROGRAM_COL: 'C',
  OUTPUT_COL: 'D',
  RANK_COL: 'E',
  ACTIVITIES_COL: 'G',
  PERFORMANCE_INDICATOR_COL: 'H',
  EXPENSE_GROUP_COL: 'J',
  GAA_OBJECT_COL: 'L',
  EXPENSE_ITEM_COL: 'N',
  QUANTITY_COL: 'O',
  UNIT_COST_COL: 'P',
  FREQUENCY_COL: 'Q',
  TOTAL_COST_COL: 'R',
  TEV_LOCATION_COL: 'S',
  PPMP_COL: 'T',
  APP_SUPPLIES_COL: 'U',
  APP_TICKET_COL: 'V',
  MANNER_OF_RELEASE_COL: 'W',
  PHYSICAL_TARGET_TOTAL_COL: 'AE',
  PHYSICAL_TARGET_MONTH_START_COL_INDEX: 'AF',
  PHYSICAL_TARGET_MONTH_END_COL_INDEX: 'AQ',
  TOTAL_OBLIGATION_COL: 'AR',
  OBLIGATION_MONTH_START_COL: 'AS',
  OBLIGATION_MONTH_END_COL: 'BD',
  TOTAL_DISBURSEMENT_COL: 'BE',
  DISBURSEMENT_MONTH_START_COL: 'BF',
  DISBURSEMENT_MONTH_END_COL: 'BQ',
  PROGRAM_ROW_INDEX: 13,
  OUTPUT_ROW_INDEX: 14,
  ACTIVITY_ROW_INDEX: 15,
  EXPENSE_ITEM_ROW_INDEX: 16,
  TARGET_ROW_INDEX: 17,
  PHYSICAL_TARGET_MONTH_COL_INDEX: 32,
  OBLIGATION_MONTH_COL_INDEX: 45,
  DISBURSEMENT_MONTH_COL_INDEX: 58,
} as const;

/**
 * Data validation settings for validating 'yes' or 'no'.
 */
const YES_NO_VALIDATION: DataValidation = {
  type: 'list',
  formulae: ['links!$P$1:$P$2'],
};

/**
 * Data validation settings for validating manner of release.
 */
const MANNER_VALIDATION: DataValidation = {
  type: 'list',
  formulae: ['links!$O$1:$O$5'],
};

/**
 * List of venues accessible by air travel.
 */
const VENUES_BY_AIR = [
  'BACOLOD',
  'BORACAY',
  'BUTUAN',
  'CAG. DE ORO',
  'CEBU CITY',
  'DAVAO CITY',
  'DUMAGUETE',
  'GENERAL SANTOS',
  'ILOILO CITY',
  'KORONADAL',
  'LEGASPI',
  'NEGROS ISLAND',
  'PALAWAN',
  'PANAY ISLANDS',
  'TACLOBAN',
  'TAGBILARAN/ BOHOL',
  'TUGUEGARAO',
  'VIGAN CITY',
  'ZAMBOANGA',
];

/**
 * Values related to expense groups.
 */
enum ExpenseGroup {
  TRAINING_SCHOLARSHIPS_EXPENSES = 'Training and Scholarship Expenses',
  SUPPLIES_EXPENSES = 'Supplies and Materials Expenses',
}

/**
 * Values related to GAA objects.
 */
enum GAAObject {
  TRAINING_EXPENSES = 'Training Expenses',
  OTHER_SUPPLIES = 'Other Supplies and Materials Expenses',
}

/**
 * Values related to the manner of release.
 */
enum ReleaseManner {
  FOR_DOWNLOAD_BOARD = 'For Downloading (Board and Lodging)',
  FOR_DOWNLOAD_PSF = 'For Downloading (Program Support Funds)',
  DIRECT_PAYMENT = 'Direct Payment',
  CASH_ADVANCE = 'Cash Advance',
}

export {
  BUDGET_ESTIMATE,
  EXPENDITURE_MATRIX,
  EXCEL_MIMETYPE,
  AUXILLIARY_SHEETS,
  YES,
  YES_NO_VALIDATION,
  MANNER_VALIDATION,
  VENUES_BY_AIR,
  ReleaseManner,
  GAAObject,
  ExpenseGroup,
  BOARD_LODGING_EXPENSE_PREFIX,
  HONORARIUM_EXPENSE_PREFIX,
  TRAVEL_EXPENSE_PREFIX,
};
