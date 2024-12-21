import { DataValidation } from 'exceljs';
import { OverheadTotalRowMap } from './types/globals.js';

/**
 * Value of yes.
 */
const YES = 'Y';

/**
 * Number of months in a year 'Y'.
 */
const MONTHS_IN_A_YEAR = 12;

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
  ACTIVITIES_COL: 'G',
  ACTIVITY_ROW_INDEX: 15,
  APP_SUPPLIES_COL: 'U',
  APP_TICKET_COL: 'V',
  COSTING_EXPENSE_ITEM_TOTAL_CELL: 'R16',
  COSTING_TOTAL_COL_INDEX: 18,
  COSTING_TOTAL_FORMULA_CELL: 'R15',
  CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX: 31,
  CURRENT_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL: 'AE14',
  DISBURSEMENT_MONTH_COL_INDEX: 58,
  DISBURSEMENT_MONTH_END_COL: 'BQ',
  DISBURSEMENT_MONTH_START_COL: 'BF',
  EXPENSE_GROUP_COL: 'J',
  EXPENSE_ITEM_COL: 'N',
  EXPENSE_ITEM_ROW_INDEX: 16,
  EXPENSE_OBJECT_FORMULA_CELL: 'K13',
  EXPENSE_OBJECT_FORMULA_COL: 11,
  EXTRA_ROWS_NUM_ROWS: 20,
  EXTRA_ROWS_START_INDEX: 20,
  FREQUENCY_COL: 'Q',
  GAA_OBJECT_COL: 'L',
  GAA_OBJECT_FORMULA_CELL: 'M16',
  GAA_OBJECT_FORMULA_COL: 13,
  HEADER_FIRST_ROW_INDEX: 1,
  HEADER_LAST_ROW_INDEX: 12,
  IS_BLANK_FORMULA_CELL1: 'BY13',
  IS_BLANK_FORMULA_CELL2: 'CA13',
  IS_BLANK_FORMULA_CELL3: 'CC13',
  IS_BLANK_FORMULA_START_COL: 77,
  MANNER_OF_RELEASE_COL: 'W',
  MILESTONES_NUM_ROWS: 2,
  MILESTONES_START_ROW: 15,
  MONTHLY_PROGRAM_NUM_ROWS: 26,
  OBLIGATION_EXPENSE_ITEM_TOTAL_CELL: 'AR16',
  OBLIGATION_MONTH_COL_INDEX: 45,
  OBLIGATION_MONTH_END_COL: 'BD',
  OBLIGATION_MONTH_START_COL: 'AS',
  OUTPUT_COL: 'D',
  OUTPUT_ROW_INDEX: 14,
  OVERHEAD_NUM_ROWS: 31,
  OVERHEAD_TOTAL_ROW_MAPPINGS: <OverheadTotalRowMap[]>[
    { rowsToAdd: 1 },
    {
      rowsToAdd: 2,
      expenseItemsCount: 2,
    },
    {
      rowsToAdd: 5,
      expenseItemsCount: 10,
    },
    {
      rowsToAdd: 16,
      expenseItemsCount: 4,
    },
    {
      rowsToAdd: 21,
      expenseItemsCount: 4,
    },
    {
      rowsToAdd: 26,
      expenseItemsCount: 4,
    },
  ],
  PERFORMANCE_INDICATOR_COL: 'H',
  PHYSICAL_TARGET_MONTH_COL_INDEX: 32,
  PHYSICAL_TARGET_MONTH_END_COL_INDEX: 'AQ',
  PHYSICAL_TARGET_MONTH_START_COL_INDEX: 'AF',
  PHYSICAL_TARGET_TOTAL_COL: 'AE',
  PPMP_COL: 'T',
  PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_COL_INDEX: 24,
  PREVIOUS_YEAR_PHYSICAL_TARGET_TOTAL_FORMULA_CELL: 'X14',
  PROGRAM_COL: 'C',
  PROGRAM_ROW_INDEX: 13,
  QUANTITY_COL: 'O',
  RANK_COL: 'E',
  TARGET_ROW_INDEX: 17,
  TEV_LOCATION_COL: 'S',
  TOTAL_COST_COL: 'R',
  TOTAL_DISBURSEMENT_COL: 'BE',
  TOTAL_OBLIGATION_COL_INDEX: 44,
  TOTAL_OBLIGATION_COL: 'AR',
  TOTAL_OBLIGATION_FORMULA_CELL: 'AR15',
  UNIT_COST_COL: 'P',
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
  ExpenseGroup,
  GAAObject,
  BOARD_LODGING_EXPENSE_PREFIX,
  HONORARIUM_EXPENSE_PREFIX,
  TRAVEL_EXPENSE_PREFIX,
  MONTHS_IN_A_YEAR,
};
