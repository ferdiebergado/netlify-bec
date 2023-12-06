import { DataValidation } from 'exceljs';

const CONVERT_URL = '/api/convert';

const BUDGET_ESTIMATE = {
  PROGRAM_CELL: 'F4',
  OUTPUT_CELL: 'F5',
  OUTPUT_INDICATOR_CELL: 'F6',
  ACTIVITY_CELL: 'F7',
  ACTIVITY_INDICATOR_CELL: 'F8',
  START_DATE_CELL: 'G13',
  VENUE_CELL: 'O13',
  TOTAL_PAX_CELL: 'H16',
  OUTPUT_PHYSICAL_TARGET_CELL: 'O6',
  ACTIVITY_PHYSICAL_TARGET_CELL: 'O8',
  EXPENSE_ITEM_COL_INDEX: 2,
  EXPENSE_ITEM_FIRST_COL_INDEX: 3,
  EXPENSE_ITEM_SECOND_COL_INDEX: 4,
  QUANTITY_CELL_INDEX: 8,
  FREQ_CELL_INDEX: 10,
  UNIT_COST_CELL_INDEX: 11,
  BOARD_LODGING_ROW_INDEX: 17,
  BOARD_LODGING_OTHER_ROW_INDEX: 24,
  TRAVEL_REGION_ROW_INDEX: 29,
  TRAVEL_CO_ROW_INDEX: 48,
  TRAVEL_OTHER_ROW_INDEX: 54,
  HONORARIUM_ROW_INDEX: 58,
  SUPPLIES_ROW_INDEX: 60,
} as const;

const EXPENDITURE_MATRIX = {
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
  PHYSICAL_TARGET_MONTH_COL_INDEX: 32,
  OBLIGATION_MONTH_COL_INDEX: 45,
  DISBURSEMENT_MONTH_COL_INDEX: 58,
} as const;

const EXPENSE_GROUP = {
  TRAINING_SCHOLARSHIPS_EXPENSES: 'Training and Scholarship Expenses',
  SUPPLIES_EXPENSES: 'Supplies and Materials Expenses',
} as const;

const GAA_OBJECT = {
  TRAINING_EXPENSES: 'Training Expenses',
  OTHER_SUPPLIES: 'Other Supplies and Materials Expenses',
} as const;

const MANNER_OF_RELEASE = {
  FOR_DOWNLOAD_BOARD: 'For Downloading (Board and Lodging)',
  FOR_DOWNLOAD_PSF: 'For Downloading (Program Support Funds)',
  DIRECT_PAYMENT: 'Direct Payment',
  CASH_ADVANCE: 'Cash Advance',
} as const;

const EXCEL_MIMETYPE =
  'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';

const AUXILLIARY_SHEETS = ['ContingencyMatrix', 'Venues', 'Honorarium'];

const YES = 'Y';

const YES_NO_VALIDATION: DataValidation = {
  type: 'list',
  formulae: ['links!$P$1:$P$2'],
};

const MANNER_VALIDATION: DataValidation = {
  type: 'list',
  formulae: ['links!$O$1:$O$5'],
};

const MAX_UPLOADS = 150;

const MAX_FILESIZE = 1024 * 1024 * 5;

const PLANE_VENUES = [
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

export {
  BUDGET_ESTIMATE,
  EXPENDITURE_MATRIX,
  GAA_OBJECT,
  MANNER_OF_RELEASE,
  EXPENSE_GROUP,
  CONVERT_URL,
  EXCEL_MIMETYPE,
  AUXILLIARY_SHEETS,
  YES,
  YES_NO_VALIDATION,
  MANNER_VALIDATION,
  MAX_UPLOADS,
  MAX_FILESIZE,
  PLANE_VENUES,
};
